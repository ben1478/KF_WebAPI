using KF_WebAPI.BaseClass;
using OfficeOpenXml;
using System.Data;
using System.Reflection;

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

        public static byte[] ExportToExcel<T>(List<T> items, Dictionary<string, string> headers)
        {
            if (items == null || items.Count == 0)
            {
                throw new ArgumentException("The list cannot be null or empty.", nameof(items));
            }

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add(typeof(T).Name);
                var properties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
                // 添加表頭
                worksheet.Cells[1, 1].Value = "序號"; 
                for (int i = 0; i < properties.Length; i++)
                {
                    var propertyName = properties[i].Name;
                    if (headers.TryGetValue(propertyName, out var header))
                    {
                        worksheet.Cells[1, i + 2].Value = header; 
                    }
                    else
                    {
                        worksheet.Cells[1, i + 2].Value = propertyName; 
                    }
                }
                // 添加表身
                for (int i = 0; i < items.Count; i++)
                {
                    var item = items[i];
                    worksheet.Cells[i + 2, 1].Value = i + 1; 
                    for (int j = 0; j < properties.Length; j++)
                    {
                        var value = properties[j].GetValue(item);
                        worksheet.Cells[i + 2, j + 2].Value = value?.ToString(); 
                    }
                }

                return package.GetAsByteArray();
            }
        }

    }
}
