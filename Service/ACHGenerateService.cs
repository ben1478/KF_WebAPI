namespace KF_WebAPI.Service
{
    using System;
    using System.IO;
    using System.Text;
    using System.Linq;
    using KF_WebAPI.BaseClass.AE;
    using KF_WebAPI.DataLogic;
    using System.Data;
    using Ionic.Zip;
    using Microsoft.VisualBasic.FileIO;
    using System.IO.Compression;
    using ZipFile = Ionic.Zip.ZipFile;

    public class AchGenerateService
    {
        // 固定參數
        private const string ORIG_ID = "52611690";          // 發動者公司統一編號
        private const string ORIG_ACC = "0012101800165758";    // 發動者帳號 (14碼)
        private const string TX_TYPE = "NSD";               // 交易型態 (樣張中明細為 NSD)
        private const string TX_CODE = "902";               // 交易代碼 (分期款代收)
        private const string SEND_ORG = "8070014";          // 發送單位代號 (永豐銀行通常為 8070014)
        private const string RECV_ORG = "9990250";          // 接收單位代號 (票交所 ACH 中心碼)

        public (byte[] FileBytes, string FileName) GenerateAchTextFile(string LaunchDate, string Ach_Bank)
        {
            // 1. 環境編碼註冊 (防止 Windows Server 環境無 Big5 字典)
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            Encoding big5 = Encoding.GetEncoding("big5");

            AE_Rpt _AE_Rpt = new AE_Rpt();
            DataTable m_dt = _AE_Rpt.GetRCMInfo(LaunchDate);
            string fileName = "";
            byte[] finalResultBytes = null;

            // 您原本的變數宣告 (請確保這些常數/欄位在類別內有定義)
            // string SEND_ORG = "..."; 
            // string RECV_ORG = "...";
            // string TX_CODE = "902";
            // string ORIG_ACC = "...";
            // string ORIG_ID = "52611690";

            string minguoDate = "0" + (Convert.ToInt16(LaunchDate.Substring(0, 4)) - 1911).ToString() + LaunchDate.Replace("/", "").Replace("-", "").Substring(4);
            string timeStr = DateTime.Now.ToString("HHmmss");

            // 建立暫存文字檔名稱（不論是否壓縮，都要定義壓縮檔內實體文字檔的名稱）
            string textFileName = "";
            if (Ach_Bank == "CTBC")
            {
                textFileName = "ACHP01_" + minguoDate + TX_CODE + ORIG_ID + ".txt";
                fileName = "ACHP01_" + minguoDate + TX_CODE + ORIG_ID + ".zip"; // 中信最終輸出檔名改為 .zip
            }
            else
            {
                textFileName = ORIG_ID + "_P01_" + LaunchDate.Replace("/", "").Replace("-", "") + ".txt";
                fileName = textFileName; // 永豐維持原本的 .txt 檔名
            }

            using (MemoryStream ms = new MemoryStream())
            {
                using (StreamWriter sw = new StreamWriter(ms, big5))
                {
                    // 換行符號設定為 Windows 的 \r\n
                    sw.NewLine = "\r\n";

                    // =========================================================================
                    // 2. 產出【控制首錄 (BOF)】 - 長度 250 Bytes
                    // =========================================================================
                    StringBuilder bof = new StringBuilder();
                    bof.Append("BOF");                                      // 1-3   首錄別
                    bof.Append("ACHP01");                                   // 4-9   資料代號 (代收為 ACHP01)
                    bof.Append(minguoDate);                                 // 10-17 處理日期
                    bof.Append(timeStr);                                    // 18-23 處理時間
                    bof.Append(SEND_ORG);                                   // 24-30 發送單位代號
                    bof.Append(RECV_ORG);                                   // 31-37 接收單位代號
                    bof.Append("V10");                                      // 38-40 版次固定 V10

                    // 41-250 補空白滿 210 個 Byte
                    string bofLine = bof.ToString().PadRightBytes(250);
                    sw.WriteLine(bofLine);

                    // =========================================================================
                    // 3. 產出【明細錄 (NSD)】 - 每行 250 Bytes
                    // =========================================================================
                    int seq = 1;
                    decimal totalAmount = 0;

                    foreach (DataRow dr in m_dt.Rows)
                    {
                        StringBuilder nsd = new StringBuilder();
                        nsd.Append("N");                                    // 1     交易型態
                        nsd.Append("SD");                                   // 2-3   交易類別 (SD=代收)
                        nsd.Append(TX_CODE);                                // 4-6   交易代號 (902)
                        nsd.Append(seq.ToString("D8"));                     // 7-14  交易序號 (靠右補零)
                        nsd.Append(SEND_ORG);                               // 15-21 提出行代號
                        nsd.Append(ORIG_ACC.PadRightBytes(16));             // 22-37 發動者帳號 (16碼，不滿補空白)
                        nsd.Append(dr["BankNo"].ToString().PadRightBytes(7));   // 38-44 提回行代號 (7碼)
                        nsd.Append(dr["AccountNo"].ToString().PadLeft(16, '0'));// 45-60 收受者帳號 (16碼)

                        // 金額處理：整數型態，10碼，靠右補零
                        long amtInt = (long)Math.Round(Convert.ToDecimal(dr["RC_amount"].ToString()), 0);
                        nsd.Append(amtInt.ToString("D10"));

                        if (Ach_Bank == "CTBC")
                        {
                            nsd.Append("  B" + ORIG_ID.PadRightBytes(10));
                        }
                        else
                        {
                            nsd.Append("00B" + ORIG_ID.PadRightBytes(10));
                        }

                        nsd.Append(dr["CS_PID"].ToString().PadRightBytes(10));

                        if (Ach_Bank == "CTBC")
                        {
                            nsd.Append("".PadRightBytes(22));
                        }
                        else
                        {
                            nsd.Append("".PadRightBytes(6));
                            nsd.Append(0.ToString("D16"));
                        }


                        nsd.Append(" ");
                        nsd.Append(dr["CS_PID"].ToString().PadRightBytes(10));
                        nsd.Append("".PadRightBytes(60));
                        nsd.Append(0.ToString("D5"));
                        string nsdLine = nsd.ToString().PadRightBytes(250);
                        sw.WriteLine(nsdLine);

                        totalAmount += amtInt;
                        seq++;
                    }

                    // =========================================================================
                    // 4. 產出【控制尾錄 (EOF)】 - 長度 250 Bytes
                    // =========================================================================
                    StringBuilder eof = new StringBuilder();
                    eof.Append("EOF");                                      // 1-3   尾錄別
                    eof.Append("ACHP01");                                   // 4-9   資料代號
                    eof.Append(minguoDate);                                 // 10-17 處理日期
                    eof.Append(SEND_ORG);                                   // 18-24 發送單位代號
                    eof.Append(RECV_ORG);                                   // 25-31 接收單位代號

                    int totalRecords = m_dt.Rows.Count;
                    eof.Append(totalRecords.ToString("D8"));                // 32-39 總筆數 (8碼，靠右補零)

                    long totalAmtLong = (long)Math.Round(totalAmount, 0);
                    eof.Append(totalAmtLong.ToString("D16"));               // 40-54 總金額 (16碼，靠右補零)

                    // 55-250 剩餘空間補空白 (196 Bytes)
                    string eofLine = eof.ToString().PadRightBytes(250);
                    sw.Write(eofLine); // 最後一行通常不加換行

                    sw.Flush();
                }

                // 取得純文字檔的 byte 陣列
                byte[] textFileBytes = ms.ToArray();

                // =========================================================================
                // 5. 判斷是否需要進行加密壓縮 (中信銀行情境)
                // =========================================================================
                if (Ach_Bank == "CTBC")
                {
                    using (MemoryStream zipMs = new MemoryStream())
                    {
                        using (ZipFile zip = new ZipFile())
                        {
                            // 指定 Zip 內部的檔名編碼支援，防止解壓後中文亂碼
                            zip.AlternateEncoding = big5;
                            zip.AlternateEncodingUsage = ZipOption.AsNecessary;

                            // 💡 安全防護更新：使用 AES 256 位元高強度加密演算法 (阻斷 Chrome 病毒誤報)
                            zip.Encryption = EncryptionAlgorithm.WinZipAes256;
                            zip.Password = "a52611690"; // 解壓縮密碼

                            // 將記憶體中的文字檔陣列注入 ZIP 結構中
                            zip.AddEntry(textFileName, textFileBytes);

                            // 儲存壓縮後的資料流
                            zip.Save(zipMs);
                        }
                        finalResultBytes = zipMs.ToArray();
                    }
                }
                else
                {
                    // 非中信銀行（永豐），直接回傳原始文字檔 Byte 陣列
                    finalResultBytes = textFileBytes;
                }

                return (finalResultBytes, fileName);
            }
        }




    }

    public static class StringByteExtensions
    {
        /// <summary>
        /// 將字串依 Big5 編碼轉換，並補齊或截斷至指定之位元組(Byte)長度
        /// </summary>
        public static string PadRightBytes(this string input, int totalBytes, char paddingChar = ' ')
        {
            if (input == null) input = string.Empty;

            // 取得 Big5 編碼器
            Encoding big5 = Encoding.GetEncoding("big5");
            byte[] sourceBytes = big5.GetBytes(input);

            byte[] resultBytes = new byte[totalBytes];

            if (sourceBytes.Length >= totalBytes)
            {
                // 超長則截斷
                Array.Copy(sourceBytes, resultBytes, totalBytes);
            }
            else
            {
                // 不足則補白
                Array.Copy(sourceBytes, resultBytes, sourceBytes.Length);
                byte paddingByte = (byte)paddingChar;
                for (int i = sourceBytes.Length; i < totalBytes; i++)
                {
                    resultBytes[i] = paddingByte;
                }
            }

            return big5.GetString(resultBytes);
        }
    }

}
