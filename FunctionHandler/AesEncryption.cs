using System.Security.Cryptography;
using System.Text;
using System;
using KF_WebAPI.BaseClass;

namespace KF_WebAPI.FunctionHandler
{
    public class AesEncryption
    {
        private readonly string _aesKey = "g5-$@3%-t@3rW$By8Ywxsf6xEg-9p-zb";
        private readonly string _aesIV = "eE8xsCqug+zYc.WU";

        public  string EncryptAES256(string p_Source)
        {
            byte[] sourceBytes = Encoding.UTF8.GetBytes(p_Source);
            var rijndael = new RijndaelManaged
            {
                KeySize = 256,
                BlockSize = 128,
                Key = Encoding.UTF8.GetBytes(_aesKey),
                IV = Encoding.UTF8.GetBytes(_aesIV),
                Mode = CipherMode.CBC,
                Padding = PaddingMode.PKCS7
            };

            byte[]? outputBytes = null;
            using (MemoryStream memoryStream = new MemoryStream())
            {
                using CryptoStream encryptStream = new(memoryStream, rijndael.CreateEncryptor(), CryptoStreamMode.Write);
                MemoryStream inputStream = new(sourceBytes);
                inputStream.CopyTo(encryptStream);
                encryptStream.FlushFinalBlock();
                outputBytes = memoryStream.ToArray();
            }
            return Convert.ToBase64String(outputBytes);

        }

        public ResultClass<string> DecryptAES256(string p_Source)
        {
            ResultClass<string> resultClass = new();
           

            byte[]? outputBytes = null;
            try
            {
                byte[] sourceBytes = Convert.FromBase64String(p_Source);
                var rijndael = new RijndaelManaged
                {
                    KeySize = 256,
                    BlockSize = 128,
                    Key = Encoding.UTF8.GetBytes(_aesKey),
                    IV = Encoding.UTF8.GetBytes(_aesIV),
                    Mode = CipherMode.CBC,
                    Padding = PaddingMode.PKCS7
                };


                using (MemoryStream memoryStream = new(sourceBytes))
                {
                    using CryptoStream decryptStream = new(memoryStream, rijndael.CreateDecryptor(), CryptoStreamMode.Read);
                    MemoryStream outputStream = new();
                    decryptStream.CopyTo(outputStream);
                    outputBytes = outputStream.ToArray();
                }
                resultClass.ResultCode = "000";
                resultClass.objResult = Encoding.UTF8.GetString(outputBytes);
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "999";
                resultClass.ResultMsg = $" 解密失敗!!: {ex.Message}";

            }
            

            return resultClass;
        }


        public string EncryptSHA256(string str)
        {
            SHA256 sha256 = new SHA256CryptoServiceProvider();//建立一個 SHA256 
            byte[] source = Encoding.Default.GetBytes(str);//將字串轉為 Byte[] 
            byte[] crypto = sha256.ComputeHash(source);//進行 SHA256 加密 
            StringBuilder builder = new StringBuilder();
            for (Int32 i = 0; i < crypto.Length; i++)
            {
                builder.Append(crypto[i].ToString("x2"));
            }
            //輸出結果，把加密後的字串從 Byte[]轉為字串 
            return builder.ToString();
        }

    }


}
