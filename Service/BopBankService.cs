using System;
using System.IO;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using KF_WebAPI.BaseClass.Bop;

namespace KF_WebAPI.Service
{
    public class BopBankService
    {
        private readonly HttpClient _httpClient;
        private readonly string _baseUrl;
        private readonly string _keySn;

        // 動態還原後存於記憶體供第 6 步使用
        private readonly byte[] _activeRawPrivateKeyBytes;
        private readonly RSA _activeRsaPrivateKey;
        public BopBankService(HttpClient httpClient, IConfiguration config)
        {
            _httpClient = httpClient;
            _baseUrl = config["Bop:BaseUrl"].TrimEnd('/');
            _keySn = config["Bop:KeySN"];

            string savedKey1645 = config["Bop:SavedKey1645"];
            string savedTempKey = config["Bop:SavedTempKey"];
            string oldPrivateKeyStr = config["Bop:OldPrivateKey"];

            // 1. 還原舊私鑰 (支援 XML 或 PEM 格式)
            using var oldRsaPrivateKey = LoadRsaPrivateKey(oldPrivateKeyStr);

            // 2. 利用舊私鑰將 savedKey1645 與 savedTempKey 解密還原成新私鑰 Raw Bytes
            _activeRawPrivateKeyBytes = DecryptNewPrivateKey(savedKey1645, savedTempKey, oldRsaPrivateKey);

            // 3. 將還原後的 Raw Bytes 載入成新 RSA 物件 (供簽章使用)
            _activeRsaPrivateKey = RSA.Create();
            try
            {
                // 優先嘗試 PKCS#8
                _activeRsaPrivateKey.ImportPkcs8PrivateKey(_activeRawPrivateKeyBytes, out _);
            }
            catch
            {
                // 若為 PKCS#1 格式
                _activeRsaPrivateKey.ImportRSAPrivateKey(_activeRawPrivateKeyBytes, out _);
            }
        }

        #region 金鑰解密與還原核心邏輯

        /// <summary>
        /// 完整對應測試程式的 LoadNewPrivateKeyFromSavedKeys 解密流程
        /// </summary>
        private static byte[] DecryptNewPrivateKey(string encKeyBase64Url, string encTempKeyBase64Url, RSA oldRsaPrivateKey)
        {
            // 1. 解密 TempKey (OAEP SHA1) 得到 AES Key (32 bytes)
            byte[] encTempKeyBytes = FromBase64Url(encTempKeyBase64Url);
            byte[] aesKey = oldRsaPrivateKey.Decrypt(encTempKeyBytes, RSAEncryptionPadding.OaepSHA1);

            // 2. 取 AES Key 前 16 bytes 當作 IV
            byte[] encKeyBytes = FromBase64Url(encKeyBase64Url);
            byte[] iv = new byte[16];
            Array.Copy(aesKey, 0, iv, 0, 16);

            // 3. AES-CBC 解密出真正的新私鑰二進位內容
            return DecryptAesCbc(encKeyBytes, aesKey, iv);
        }

        /// <summary>
        /// 載入舊私鑰字串
        /// </summary>
        private static RSA LoadRsaPrivateKey(string keyString)
        {
            var rsa = RSA.Create();
            string trimmed = keyString.Trim();

            if (trimmed.StartsWith("<RSAKeyValue>"))
            {
                rsa.FromXmlString(trimmed);
            }
            else if (trimmed.Contains("PRIVATE KEY"))
            {
                rsa.ImportFromPem(trimmed);
            }
            else
            {
                // 假設是 Base64 PKCS#8 或 PKCS#1
                byte[] keyBytes = Convert.FromBase64String(trimmed);
                try
                {
                    rsa.ImportPkcs8PrivateKey(keyBytes, out _);
                }
                catch
                {
                    rsa.ImportRSAPrivateKey(keyBytes, out _);
                }
            }
            return rsa;
        }

        private static byte[] DecryptAesCbc(byte[] cipherText, byte[] key, byte[] iv)
        {
            using (Aes aes = Aes.Create())
            {
                aes.Mode = CipherMode.CBC;
                aes.Padding = PaddingMode.PKCS7;
                aes.Key = key;
                aes.IV = iv;

                using (var decryptor = aes.CreateDecryptor())
                using (var ms = new MemoryStream(cipherText))
                using (var cs = new CryptoStream(ms, decryptor, CryptoStreamMode.Read))
                using (var msOut = new MemoryStream())
                {
                    cs.CopyTo(msOut);
                    return msOut.ToArray();
                }
            }
        }

        #endregion


        /// <summary>
        /// 6. 虛擬帳號註冊 QRCode (企業 -> 板信)
        /// </summary>
        public async Task<RegisterAccountQRCodeResult> RegisterAccountQRCodeAsync(RegisterAccountQRCodeRequest req)
        {
            req.KeySN = _keySn;
            string url = $"{_baseUrl}/RegisterAccountQRCode";

            // 1. 簽章值 Sign = SHA256withRSA(VTACCT + SINGLELIMIT + TOTLIMIT + DEADLINE) -> Base64URL -> Hex
            string signSource = (req.VTACCT ?? "") + (req.SINGLELIMIT ?? "") + (req.TOTLIMIT ?? "") + (req.DEADLINE ?? "");
            byte[] signBytes = _activeRsaPrivateKey.SignData(Encoding.UTF8.GetBytes(signSource), HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
            string base64UrlSign = ToBase64Url(signBytes);
            string signHex = StringToHex(base64UrlSign);

            // 2. HmacSHA256 壓碼 (使用還原出的 _activeRawPrivateKeyBytes 前 12 bytes Hex + KeySN)
            string seed = GetSeed(_activeRawPrivateKeyBytes, req.KeySN);
            string hmacSource = (req.UUID ?? "") + (req.VTACCT ?? "") + (req.CHKREPEAT ?? "") +
                                (req.CHKSINGLE ?? "") + (req.SINGLELIMIT ?? "") + (req.CHKTOT ?? "") +
                                (req.TOTLIMIT ?? "") + (req.QRCODEAMOUNT ?? "") + (req.CHKDEADLINE ?? "") +
                                (req.DEADLINE ?? "") + (req.RECYCLEDATE ?? "") + req.KeySN;
            string sha256Hex = CalcHmacSha256(hmacSource, seed);

            // 3. 組裝 Request JSON
            var reqData = new JObject
            {
                ["UUID"] = req.UUID,
                ["VTACCT"] = req.VTACCT,
                ["CHKREPEAT"] = req.CHKREPEAT,
                ["CHKSINGLE"] = req.CHKSINGLE,
                ["SINGLELIMIT"] = req.SINGLELIMIT,
                ["CHKTOT"] = req.CHKTOT,
                ["TOTLIMIT"] = req.TOTLIMIT,
                ["QRCODEAMOUNT"] = req.QRCODEAMOUNT,
                ["QRCODENOTE"] = req.QRCODENOTE,
                ["CHKDEADLINE"] = req.CHKDEADLINE,
                ["DEADLINE"] = req.DEADLINE,
                ["RECYCLEDATE"] = req.RECYCLEDATE,
                ["Sign"] = signHex,
                ["KeySN"] = req.KeySN
            };

            var rootJson = new JObject
            {
                ["Data"] = reqData,
                ["SHA-256"] = sha256Hex
            };

            string jsonReq = rootJson.ToString(Formatting.None);

            // 4. HTTP 發送
            using var request = new HttpRequestMessage(HttpMethod.Post, url)
            {
                Content = new StringContent(jsonReq, Encoding.UTF8, "application/json")
            };
            request.Headers.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));

            var response = await _httpClient.SendAsync(request);
            string jsonResp = await response.Content.ReadAsStringAsync();

            var result = new RegisterAccountQRCodeResult
            {
                RequestJson = jsonReq,
                ResponseJson = $"[HTTP {(int)response.StatusCode} {response.StatusCode}]\r\n{jsonResp}"
            };

            if (string.IsNullOrWhiteSpace(jsonResp))
            {
                result.ReturnCode = $"HTTP_{(int)response.StatusCode}_EMPTY";
                result.ReturnMsg = "伺服器未回傳內容。";
                return result;
            }

            // 5. 解析 Response
            try
            {
                JObject respJObj = JObject.Parse(jsonResp);
                JToken dataToken = respJObj["Data"];
                if (dataToken != null)
                {
                    result.ReturnCode = dataToken.Value<string>("ReturnCode");
                    result.ReturnMsg = dataToken.Value<string>("ReturnMsg");
                    result.UUID = dataToken.Value<string>("UUID");
                    result.VTACCT = dataToken.Value<string>("VTACCT");
                    result.QRCode = dataToken.Value<string>("QRCode");
                }
                result.SHA256 = respJObj.Value<string>("SHA-256");
            }
            catch (Exception ex)
            {
                result.ReturnCode = "PARSE_ERROR";
                result.ReturnMsg = ex.Message;
            }

            return result;
        }
        #region 加解密小工具
        private static string ToBase64Url(byte[] input) =>
            Convert.ToBase64String(input).Replace("+", "-").Replace("/", "_").TrimEnd('=');
        private static byte[] FromBase64Url(string base64Url)
        {
            string base64 = base64Url.Replace("-", "+").Replace("_", "/");
            switch (base64.Length % 4)
            {
                case 2: base64 += "=="; break;
                case 3: base64 += "="; break;
            }
            return Convert.FromBase64String(base64);
        }
        private static string StringToHex(string input)
        {
            byte[] bytes = Encoding.ASCII.GetBytes(input);
            StringBuilder sb = new StringBuilder();
            foreach (byte b in bytes) sb.Append(b.ToString("X2"));
            return sb.ToString();
        }

        private static string GetSeed(byte[] rawPrivateKeyBytes, string keySn)
        {
            byte[] first12 = new byte[12];
            Array.Copy(rawPrivateKeyBytes, 0, first12, 0, 12);
            StringBuilder sb = new StringBuilder();
            foreach (byte b in first12) sb.Append(b.ToString("X2"));
            return sb.ToString() + keySn;
        }

        private static string CalcHmacSha256(string data, string seedHex)
        {
            byte[] keyBytes = Encoding.UTF8.GetBytes(seedHex);
            byte[] dataBytes = Encoding.UTF8.GetBytes(data);
            using var hmac = new HMACSHA256(keyBytes);
            byte[] hash = hmac.ComputeHash(dataBytes);
            StringBuilder sb = new StringBuilder();
            foreach (byte b in hash) sb.Append(b.ToString("X2"));
            return sb.ToString();
        }
        #endregion

      
    }
}
