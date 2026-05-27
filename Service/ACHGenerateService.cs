namespace KF_WebAPI.Service
{
    using System;
    using System.IO;
    using System.Text;
    using System.Linq;
    using KF_WebAPI.BaseClass.AE;
    using KF_WebAPI.DataLogic;
    using System.Data;

    public class AchGenerateService
    {
        // 固定參數
        private const string ORIG_ID = "52611690";          // 發動者公司統一編號
        private const string ORIG_ACC = "0012101800165758";    // 發動者帳號 (14碼)
        private const string TX_TYPE = "NSD";               // 交易型態 (樣張中明細為 NSD)
        private const string TX_CODE = "902";               // 交易代碼 (分期款代收)
        private const string SEND_ORG = "8070014";          // 發送單位代號 (永豐銀行通常為 8070014)
        private const string RECV_ORG = "9990250";          // 接收單位代號 (票交所 ACH 中心碼)

        public byte[] GenerateAchTextFile(string LaunchDate)
        {
            // 1. 環境編碼註冊 (防止 Windows Server 環境無 Big5 字典)
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            Encoding big5 = Encoding.GetEncoding("big5");
            AE_Rpt _AE_Rpt = new AE_Rpt();
            DataTable m_dt = _AE_Rpt.GetRCMInfo(LaunchDate);
            string Ach_Bank = "";

            string minguoDate = "0"+(Convert.ToInt16( LaunchDate.Substring(0, 4)) - 1911).ToString()+ LaunchDate.Replace("/","").Substring(4);
            string timeStr = DateTime.Now.ToString("HHmmss");

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
                        // 轉換發動日期格式

                        Ach_Bank = dr["Ach_Bank"].ToString();
                        StringBuilder nsd = new StringBuilder();
                        nsd.Append("N");                                    // 1     交易型態
                        nsd.Append("SD");                                   // 2-3   交易類別 (SD=代收)
                        nsd.Append(TX_CODE);                                // 4-6   交易代號 (902)
                        nsd.Append(seq.ToString("D8"));                     // 7-14  交易序號 (靠右補零)
                        nsd.Append(SEND_ORG);                               // 15-21 提出行代號
                        nsd.Append(ORIG_ACC.PadRightBytes(16));             // 22-37 發動者帳號 (16碼，不滿補空白)
                        nsd.Append(dr["BankNo"].ToString().PadRightBytes(7));   // 38-44 提回行代號 (7碼)
                        nsd.Append(dr["AccountNo"].ToString().PadRightBytes(16));// 45-60 收受者帳號 (16碼)

                        // 金額處理：整數型態，10碼，靠右補零
                        long amtInt = (long)Math.Round(Convert.ToDecimal(dr["Amount"].ToString()), 0);
                        nsd.Append(amtInt.ToString("D10"));                 // 61-70 金額

                        //nsd.Append(" ");                                    // 71    退件理由代號 (發動時留空)
                        //nsd.Append(" ");                                    // 72    提示交換次序 (1碼空白)
                        nsd.Append("00B"+ORIG_ID.PadRightBytes(10));              // 73-82 發動者統一編號 (10碼)
                       // nsd.Append("  ");
                        nsd.Append(dr["CS_PID"].ToString().PadRightBytes(10));    // 83-92 收受者統一編號 (10碼)
                        nsd.Append("".PadRightBytes(6));
                        nsd.Append(0.ToString("D16"));
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
                    eof.Append(totalAmtLong.ToString("D16"));               // 40-54 總金額 (15碼，靠右補零)

                    // 55-250 剩餘空間補空白 (196 Bytes)
                    string eofLine = eof.ToString().PadRightBytes(250);
                    sw.Write(eofLine); // 最後一行通常不加換行

                    sw.Flush();
                }
                return ms.ToArray();
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
