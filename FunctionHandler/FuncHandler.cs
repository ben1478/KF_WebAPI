using KF_WebAPI.BaseClass;
using System.Data;

namespace KF_WebAPI.FunctionHandler
{
    public class FuncHandler
    {
        /// <summary>
        /// 分頁
        /// </summary>
        /// <returns></returns>
        public static DataTable GetPage(DataTable dt, int page, int pageSize)
        {
            if (dt == null || dt.Rows.Count == 0 || page <= 0 || pageSize <= 0)
            {
                return new DataTable();
            }

            int startIndex = (page - 1) * pageSize;
            int endIndex = startIndex + pageSize;

            if (startIndex >= dt.Rows.Count)
            {
                return new DataTable();
            }

            if (endIndex > dt.Rows.Count)
            {
                endIndex = dt.Rows.Count; 
            }

            DataTable pageTable = dt.Clone();

            for (int i = startIndex; i < endIndex; i++)
            {
                pageTable.ImportRow(dt.Rows[i]);
            }

            return pageTable;
        }
        /// <summary>
        /// 獲得Check_Num
        /// </summary>
        public string GetCheckNum()
        {
            Random random = new Random();

            // 生成兩個隨機數字
            string checkNum1 = GetRandomNumber(random);
            string checkNum2 = GetRandomNumber(random);

            // 獲取當前的日期和時間
            DateTime now = DateTime.Now;
            string year = now.Year.ToString();
            string month = now.Month.ToString("D2");
            string day = now.Day.ToString("D2");
            string hour = now.Hour.ToString("D2");
            string minute = now.Minute.ToString("D2");
            string second = now.Second.ToString("D2");

            // 拼接各部分
            string cknum = year + month + day + hour + minute + second + checkNum1 + checkNum2;

            return cknum;
        }

        private static string GetRandomNumber(Random random)
        {
            // 生成範圍為 [100001, 199999] 的隨機數字
            int number = random.Next(100001, 200000);
            // 提取中間的 5 位數字
            string numberString = number.ToString();
            return numberString.Substring(1, 5);
        }

    }
}
