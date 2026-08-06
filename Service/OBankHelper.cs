using System;
using System.Collections.Generic;

namespace KF_WebAPI.Service
{

    public class SupermarketBarcodeResult
    {
        public string Barcode1 { get; set; }
        public string Barcode2 { get; set; }
        public string Barcode3 { get; set; }
    }

    public static class OBankHelper
    {
        #region 1. 王道虛擬帳號生成與還原 (16碼)

        /// <summary>
        /// 產生王道銀行 16 碼虛擬帳號 (搭配 RCD_id)
        /// </summary>
        /// <param name="corpCode">5碼企業識別碼 (例如 "88052")</param>
        /// <param name="rcdId">Receivable_D 的 PK (RCD_id，如 10000001)</param>
        /// <returns>16碼完整虛擬帳號 (例如: 8805200100000013)</returns>
        public static string GenerateVirtualAccount(string corpCode, decimal rcdId)
        {
            if (string.IsNullOrWhiteSpace(corpCode) || corpCode.Length != 5)
                throw new ArgumentException("企業識別碼須為 5 碼");

            // 將 RCD_id 轉字串並靠左補零至 10 碼
            string billNo = Convert.ToInt64(rcdId).ToString().PadLeft(10, '0');
            if (billNo.Length > 10)
                throw new ArgumentOutOfRangeException(nameof(rcdId), "RCD_id 長度超出 10 碼上限");

            string account15 = corpCode + billNo;
            int[] weights = new int[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 0, 1, 2, 3, 4, 5 };

            int totalWeightSum = 0;
            for (int i = 0; i < 15; i++)
            {
                totalWeightSum += (account15[i] - '0') * weights[i];
            }

            int R = totalWeightSum % 11;
            int X;
            if (R == 1) X = 11;
            else if (R == 0) X = 10;
            else X = R;

            int checkDigit = 11 - X;

            return account15 + checkDigit.ToString();
        }

        /// <summary>
        /// 從王道 ATM 銷帳檔中的 16 碼收款帳號反推還原出 RCD_id
        /// </summary>
        /// <param name="virtualAccount">16 碼虛擬帳號 (例如: 8805200100000013)</param>
        /// <returns>還原後的 RCD_id (例如: 10000001)</returns>
        public static decimal ExtractRcdIdFromVirtualAccount(string virtualAccount)
        {
            if (string.IsNullOrWhiteSpace(virtualAccount) || virtualAccount.Length < 16)
                throw new ArgumentException("虛擬帳號長度不足 16 碼");

            // 取中間 10 碼銷帳編號 (索引 5 開始取 10 碼)
            string billNo = virtualAccount.Substring(5, 10);
            return decimal.Parse(billNo); // 自動去前導 0 轉回 RCD_id
        }

        #endregion

        #region 2. 超商三段式代收條碼生成 (9 + 16 + 15)

        // 字母轉數字對照表 (S 之後從 2 開始)
        private static readonly Dictionary<char, int> CharMap = new Dictionary<char, int>
    {
        {'A',1}, {'B',2}, {'C',3}, {'D',4}, {'E',5}, {'F',6}, {'G',7}, {'H',8}, {'I',9},
        {'J',1}, {'K',2}, {'L',3}, {'M',4}, {'N',5}, {'O',6}, {'P',7}, {'Q',8}, {'R',9},
        {'S',2}, {'T',3}, {'U',4}, {'V',5}, {'W',6}, {'X',7}, {'Y',8}, {'Z',9}
    };

        private static int GetCharValue(char c)
        {
            if (char.IsDigit(c)) return c - '0';
            char upper = char.ToUpper(c);
            return CharMap.ContainsKey(upper) ? CharMap[upper] : 0;
        }

        /// <summary>
        /// 產生超商三段式條碼 (搭配 RCD_id)
        /// </summary>
        /// <param name="dueDate">應繳日期</param>
        /// <param name="isFeeIncluded">true: XHM (手續費內扣), false: XHN (手續費外加)</param>
        /// <param name="corpCode">5 碼客戶編號 (例如 "88052")</param>
        /// <param name="rcdId">Receivable_D 的 PK (RCD_id)</param>
        /// <param name="amount">應繳金額 (單筆上限 20,000 元)</param>
        public static (string Barcode1, string Barcode2, string Barcode3) GenerateSupermarketBarcodes(
            DateTime dueDate,
            bool isFeeIncluded,
            string corpCode,
            decimal rcdId,
            decimal amount)
        {
            if (amount > 20000) throw new ArgumentOutOfRangeException(nameof(amount), "超商代收單筆金額上限為 20,000 元");
            if (string.IsNullOrWhiteSpace(corpCode) || corpCode.Length != 5)
                throw new ArgumentException("客戶編號須為 5 碼");

            // 1. 第一段條碼 (9碼): yymmdd(6碼) + 代收項目(3碼)
            string yymmdd = (dueDate.Year % 100).ToString("D2") + dueDate.ToString("MMdd");
            string itemCode = isFeeIncluded ? "XHM" : "XHN";
            string barcode1 = yymmdd + itemCode;

            // 2. 第二段條碼 (16碼): 5碼客戶編號 + 11碼銷帳編號 (將 RCD_id 補零至 11 碼)
            string custBillNo = Convert.ToInt64(rcdId).ToString().PadLeft(11, '0');
            string barcode2 = corpCode + custBillNo;

            // 3. 第三段條碼與校對碼計算 (15碼)
            string mmdd = dueDate.ToString("MMdd");
            long amtLong = (long)Math.Round(amount, 0);
            string amountStr = amtLong.ToString("D9"); // 靠右補 0 至 9 碼

            // 奇數位加總
            int oddSum = 0;
            for (int i = 0; i < 9; i += 2) oddSum += GetCharValue(barcode1[i]);
            for (int i = 0; i < 16; i += 2) oddSum += GetCharValue(barcode2[i]);
            oddSum += GetCharValue(mmdd[0]) + GetCharValue(mmdd[2]) + 0 +
                      GetCharValue(amountStr[1]) + GetCharValue(amountStr[3]) + GetCharValue(amountStr[5]) + GetCharValue(amountStr[7]);

            int r1 = oddSum % 11;
            char checkCode1 = r1 == 0 ? 'A' : (r1 == 10 ? 'B' : (char)('0' + r1));

            // 偶數位加總
            int evenSum = 0;
            for (int i = 1; i < 9; i += 2) evenSum += GetCharValue(barcode1[i]);
            for (int i = 1; i < 16; i += 2) evenSum += GetCharValue(barcode2[i]);
            evenSum += GetCharValue(mmdd[1]) + GetCharValue(mmdd[3]) + 0 + 0 +
                       GetCharValue(amountStr[0]) + GetCharValue(amountStr[2]) + GetCharValue(amountStr[4]) + GetCharValue(amountStr[6]) + GetCharValue(amountStr[8]);

            int r2 = evenSum % 11;
            char checkCode2 = r2 == 0 ? 'X' : (r2 == 10 ? 'Y' : (char)('0' + r2));

            string barcode3 = $"{mmdd}{checkCode1}{checkCode2}{amountStr}";

            return (barcode1, barcode2, barcode3);
        }

        /// <summary>
        /// 從王道超商銷帳檔中的 16 碼銷帳編號反推還原出 RCD_id
        /// </summary>
        /// <param name="cvsBillNo">16 碼銷帳編號 (如: 8805200010000001)</param>
        /// <returns>還原後的 RCD_id (例如: 10000001)</returns>
        public static decimal ExtractRcdIdFromCvsBillNo(string cvsBillNo)
        {
            if (string.IsNullOrWhiteSpace(cvsBillNo) || cvsBillNo.Length < 16)
                throw new ArgumentException("銷帳編號長度不足 16 碼");

            // 取後 11 碼 (索引 5 開始)
            string rcdStr = cvsBillNo.Substring(5, 11);
            return decimal.Parse(rcdStr); // 自動去前導 0 轉回 RCD_id
        }

        #endregion
    }
}
