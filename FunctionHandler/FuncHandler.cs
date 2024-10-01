using KF_WebAPI.BaseClass;
using KF_WebAPI.BaseClass.AE;
using OfficeOpenXml;
using System.Data;
using Microsoft.Data.SqlClient;
using System.Drawing;
using System.Reflection;
using System.Reflection.PortableExecutable;

namespace KF_WebAPI.FunctionHandler
{
    public class FuncHandler
    {
        /// <summary>
        /// DataTable分頁
        /// </summary>
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
        /// List分頁
        /// </summary>
        public static List<T> GetPagedList<T>(List<T> list, int pageNumber, int pageSize)
        {
            int totalItems = list.Count;

            int totalPages = (int)Math.Ceiling(totalItems / (double)pageSize);

            if (pageNumber < 1 || pageNumber > totalPages)
            {
                throw new ArgumentOutOfRangeException(nameof(pageNumber), "Page number is out of range.");
            }

            int skip = (pageNumber - 1) * pageSize;
            int take = pageSize;

            return list.Skip(skip).Take(take).ToList();
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
                worksheet.Cells[1, 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
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

                    worksheet.Cells[1, i + 2].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[1, i + 2].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);  
                }

                // 添加表頭邊框
                using (var range = worksheet.Cells[1, 1, 1, properties.Length + 1])
                {
                    range.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    range.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    range.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                }

                // 添加表身
                for (int i = 0; i < items.Count; i++)
                {
                    var item = items[i];
                    worksheet.Cells[i + 2, 1].Value = i + 1;

                    for (int j = 0; j < properties.Length; j++)
                    {
                        var value = properties[j].GetValue(item);
                        
                        // 檢查類型並設置值
                        if (value is int || value is long || value is float || value is double || value is decimal)
                        {
                            worksheet.Cells[i + 2, j + 2].Value = value; // 直接赋值
                        }
                        else
                        {
                            worksheet.Cells[i + 2, j + 2].Value = value?.ToString(); // 字符串赋值
                        }
                    }

                    // 添加表身邊框
                    using (var range = worksheet.Cells[i + 2, 1, i + 2, headers.Count + 1])
                    {
                        range.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        range.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        range.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    }
                }

                // 列寬自動調整
                for (int j = 0; j < properties.Length + 1; j++)
                {
                    worksheet.Column(j + 1).AutoFit();
                }
                return package.GetAsByteArray();
            }
        }

        public static SpecialClass CheckSpecial(string[] strings,string U_num)
        {
            SpecialClass sc=new SpecialClass();
            sc.special_check = "N";
            sc.BC_Strings = "zz";
            sc.U_num = U_num;   
            ADOData _adoData = new ADOData();

            foreach (string s in strings)
            {
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = "Select SP_type from Special_set Where U_num = @U_num AND SP_id = @SP_id AND del_tag='0'";
                parameters.Add(new SqlParameter("@U_num", U_num));
                parameters.Add(new SqlParameter("@SP_id", s));
                #endregion
                DataTable dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
                if (dtResult.Rows.Count > 0)
                {
                    DataRow row = dtResult.Rows[0];
                    string spType = row["SP_type"].ToString();
                    if (spType == "1")
                    {
                        sc.special_check = "Y";
                        switch (s)
                        {
                            case "7020":
                                sc.BC_Strings = sc.BC_Strings + ",BC0100";
                                break;
                            case "7021":
                                sc.BC_Strings = sc.BC_Strings + ",BC0200";
                                break;
                            case "7022":
                                sc.BC_Strings = sc.BC_Strings + ",BC0600";
                                break;
                            case "7023":
                                sc.BC_Strings = sc.BC_Strings + ",BC0300";
                                break;
                            case "7024":
                                sc.BC_Strings = sc.BC_Strings + ",BC0500";
                                break;
                            case "7025":
                                sc.BC_Strings = sc.BC_Strings + ",BC0400";
                                break;
                        }
                    }
                    
                }
            }
            return sc;
        }

        public static byte[] FeatDailyToExcel<T>(List<T> items, Dictionary<string, string> headers, string name, string datestring)
        {
            if (items == null || items.Count == 0)
            {
                throw new ArgumentException("列表不能為 null 或空。", nameof(items));
            }

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add(typeof(T).Name);
                var properties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
                int countperple = items.Count - 1;

                // 添加合併標題
                worksheet.Cells[1, 1].Value = "國峯租賃股份有限公司(" + name + ") 1+" + countperple + "人     " + datestring;
                worksheet.Cells[1, 1, 1, headers.Count].Merge = true; // 合併儲存格
                worksheet.Cells[1, 1, 1, headers.Count].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center; // 標題居中
                worksheet.Cells[1, 1, 1, headers.Count].Style.Font.Bold = true; // 標題加粗

                // 添加子標題
                int headerIndex = 1;
                foreach (var header in headers)
                {
                    var cell = worksheet.Cells[2, headerIndex++];
                    cell.Value = header.Value;
                    // 設置儲存格底色為淺藍色
                    cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    cell.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                }

                // 添加表頭邊框
                using (var range = worksheet.Cells[1, 1, 2, headers.Count])
                {
                    range.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    range.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    range.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                }

                // 添加表身
                int totalRows = items.Count;
                for (int i = 0; i < totalRows; i++)
                {
                    var item = items[i];
                    int columnIndex = 1;
                    foreach (var headerKey in headers.Keys)
                    {
                        var property = typeof(T).GetProperty(headerKey);
                        if (property != null)
                        {
                            var value = property.GetValue(item);
                            worksheet.Cells[i + 3, columnIndex++].Value = value; // 從第三行開始填寫數據
                        }
                        else
                        {
                            worksheet.Cells[i + 3, columnIndex++].Value = ""; // 如果找不到屬性，設置為空
                        }

                        if (headerKey == "AchievementRate")
                        {
                            int target_quota = (int)worksheet.Cells[i + 3, 17].Value;
                            if (target_quota > 0)
                            {
                                var formula = $"({worksheet.Cells[i + 3, 11].Value} + {worksheet.Cells[i + 3, 12].Value}) / {worksheet.Cells[i + 3, 17].Value}";
                                worksheet.Cells[i + 3, 18].Formula = formula;
                                worksheet.Cells[i + 3, 18].Style.Numberformat.Format = "0.00%";
                            }
                            else
                            {
                                worksheet.Cells[i + 3, 18].Value = "--";
                            }
                        }
                    }

                    // 添加表身邊框
                    using (var range = worksheet.Cells[i + 3, 1, i + 3, headers.Count])
                    {
                        range.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        range.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        range.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    }
                }

                // 添加總計行
                int totalRowIndex = items.Count + 3; // 調整以適應合併標題和子標題
                worksheet.Cells[totalRowIndex, 1].Value = "總計";
                worksheet.Cells[totalRowIndex, 1, totalRowIndex, 2].Merge = true; // 合併儲存格
                worksheet.Cells[totalRowIndex, 1, totalRowIndex, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center; // 標題居中
                int totalColumnCount = headers.Count;
                int totalColumnIndex = 3; //從第三列開始
                foreach (var headerKey in headers.Keys)
                {
                    if (totalColumnIndex <= totalColumnCount)
                    {
                        var formula = $"SUM({worksheet.Cells[3, totalColumnIndex].Address}:{worksheet.Cells[totalRows + 2, totalColumnIndex].Address})";
                        worksheet.Cells[totalRowIndex, totalColumnIndex].Formula = formula; // 設置總和公式
                    }

                    if (headerKey == "AchievementRate")
                    {
                        var formula = $"({worksheet.Cells[totalRowIndex, 11].Address} + {worksheet.Cells[totalRowIndex, 12].Address}) / {worksheet.Cells[totalRowIndex, 17].Address}";
                        worksheet.Cells[totalRowIndex, 18].Formula = formula;
                        worksheet.Cells[totalRowIndex, 18].Style.Numberformat.Format = "0.00%";
                    }

                    totalColumnIndex++;
                }

                // 添加總計邊框
                using (var range = worksheet.Cells[totalRowIndex, 1, totalRowIndex, headers.Count])
                {
                    range.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    range.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    range.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                }

                return package.GetAsByteArray();
            }
        }

        public static byte[] FeatDailyToExcelAgain<T>(byte[] existingFileBytes, List<T> items, Dictionary<string, string> headers, string name, string datestring)
        {
            using (var package = new ExcelPackage(new MemoryStream(existingFileBytes)))
            {
                var worksheet = package.Workbook.Worksheets[0];

                int existingRowCount = worksheet.Dimension.Rows; // 抓現有的行數
                int startRowIndex = existingRowCount + 2; // 留出一行空白
                int countperple = items.Count - 1;

                // 添加合併標題
                worksheet.Cells[startRowIndex, 1].Value = "國峯租賃股份有限公司(" + name + ") 1+" + countperple + "人     " + datestring;
                worksheet.Cells[startRowIndex, 1, startRowIndex, headers.Count].Merge = true; // 合併儲存格
                worksheet.Cells[startRowIndex, 1, startRowIndex, headers.Count].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center; // 標題居中
                worksheet.Cells[startRowIndex, 1, startRowIndex, headers.Count].Style.Font.Bold = true; // 標題加粗

                // 添加子標題
                startRowIndex++;
                int headerIndex = 1;
                foreach (var header in headers)
                {
                    var cell = worksheet.Cells[startRowIndex, headerIndex++];
                    cell.Value = header.Value;
                    // 設置儲存格底色為淺藍色
                    cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    cell.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                }

                // 添加表頭邊框
                using (var range = worksheet.Cells[startRowIndex - 1, 1, startRowIndex, headers.Count])
                {
                    range.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    range.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    range.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                }

                // 添加表身
                int totalRows = items.Count;
                for (int i = 0; i < totalRows; i++)
                {
                    var item = items[i];
                    int columnIndex = 1;
                    foreach (var headerKey in headers.Keys)
                    {
                        var property = typeof(T).GetProperty(headerKey);
                        if (property != null)
                        {
                            var value = property.GetValue(item);
                            worksheet.Cells[startRowIndex + i + 1, columnIndex++].Value = value; // 填入数据
                        }
                        else
                        {
                            worksheet.Cells[startRowIndex + i + 1, columnIndex++].Value = "";
                        }

                        if (headerKey == "AchievementRate")
                        {
                            int target_quota = (int)worksheet.Cells[startRowIndex + i + 1, 17].Value;
                            if (target_quota > 0)
                            {
                                var formula = $"({worksheet.Cells[startRowIndex + i + 1, 11].Value} + {worksheet.Cells[startRowIndex + i + 1, 12].Value}) / {worksheet.Cells[startRowIndex + i + 1, 17].Value}";
                                worksheet.Cells[startRowIndex + i + 1, 18].Formula = formula;
                                worksheet.Cells[startRowIndex + i + 1, 18].Style.Numberformat.Format = "0.00%";
                            }
                            else
                            {
                                worksheet.Cells[startRowIndex + i + 1, 18].Value = "--";
                            }
                        }
                    }

                    // 添加表身邊框
                    using (var range = worksheet.Cells[startRowIndex + i + 1, 1, startRowIndex + i + 1, headers.Count])
                    {
                        range.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        range.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        range.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    }
                }

                // 添加總計行
                int totalRowIndex = startRowIndex + totalRows + 1;
                worksheet.Cells[totalRowIndex, 1].Value = "總計";
                worksheet.Cells[totalRowIndex, 1, totalRowIndex, 2].Merge = true; // 合併儲存格
                worksheet.Cells[totalRowIndex, 1, totalRowIndex, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center; // 標題居中
                int totalColumnCount = headers.Count;
                int totalColumnIndex = 3; // 從第三列開始
                foreach (var headerKey in headers.Keys)
                {
                    if (totalColumnIndex <= totalColumnCount)
                    {
                        var formula = $"SUM({worksheet.Cells[startRowIndex + 1, totalColumnIndex].Address}:{worksheet.Cells[startRowIndex + totalRows, totalColumnIndex].Address})";
                        worksheet.Cells[totalRowIndex, totalColumnIndex].Formula = formula; // 設置總和公式
                    }

                    if (headerKey == "AchievementRate")
                    {
                        var formula = $"IF({worksheet.Cells[totalRowIndex, 17].Address} = 0, \"--\", ({worksheet.Cells[totalRowIndex, 11].Address} + {worksheet.Cells[totalRowIndex, 12].Address}) / {worksheet.Cells[totalRowIndex, 17].Address})";
                        worksheet.Cells[totalRowIndex, 18].Formula = formula;
                        worksheet.Cells[totalRowIndex, 18].Style.Numberformat.Format = "0.00%";
                    }

                    totalColumnIndex++;
                }

                // 添加總計邊框
                using (var range = worksheet.Cells[totalRowIndex, 1, totalRowIndex, headers.Count])
                {
                    range.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    range.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    range.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                }

                return package.GetAsByteArray();
            }
        }

        public static byte[] FeatDailyToExcelFooter(byte[] existingFileBytes, FeatDailyReport_excel_Total model, Dictionary<string, string> headers, string datestring)
        {
            using (var package = new ExcelPackage(new MemoryStream(existingFileBytes)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int existingRowCount = worksheet.Dimension.Rows;//抓現有的行數
                int startRowIndex = existingRowCount + 2; // 留出一行空白

                // 添加合併標題
                worksheet.Cells[startRowIndex, 1].Value = "國峯租賃股份有限公司(全區總計)     " + datestring;
                worksheet.Cells[startRowIndex, 1, startRowIndex, headers.Count + 1].Merge = true; // 合併儲存格
                worksheet.Cells[startRowIndex, 1, startRowIndex, headers.Count + 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center; // 標題居中
                worksheet.Cells[startRowIndex, 1, startRowIndex, headers.Count + 1].Style.Font.Bold = true; // 標題加粗

                // 添加子標題
                startRowIndex++;
                int headerIndex = 1;
                foreach (var header in headers)
                {
                    var cell = worksheet.Cells[startRowIndex, headerIndex++];
                    cell.Value = header.Value;
                    // 設置儲存格底色為淺藍色
                    cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    cell.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                }

                // 添加表頭邊框
                using (var range = worksheet.Cells[startRowIndex - 1, 1, startRowIndex, headers.Count])
                {
                    range.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    range.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    range.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                }

                // 添加表身
                int columnIndex = 3;
                worksheet.Cells[startRowIndex + 1, 1].Value = "總計";
                worksheet.Cells[startRowIndex + 1, 1, startRowIndex + 1, 2].Merge = true; // 合併儲存格
                worksheet.Cells[startRowIndex + 1, 1, startRowIndex + 1, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center; // 標題居中
                foreach (var headerKey in headers.Keys) 
                {
                    var property = model.GetType().GetProperty(headerKey);
                    if (property != null)
                    {
                        var value = property.GetValue(model);
                        worksheet.Cells[startRowIndex + 1, columnIndex++].Value = value; // 填入数据
                    }

                    if (headerKey == "AchievementRate")
                    {
                        int target_quota = (int)worksheet.Cells[startRowIndex + 1, 17].Value;
                        if (target_quota > 0)
                        {
                            var formula = $"({worksheet.Cells[startRowIndex + 1, 11].Value} + {worksheet.Cells[startRowIndex + 1, 12].Value}) / {worksheet.Cells[startRowIndex + 1, 17].Value}";
                            worksheet.Cells[startRowIndex + 1, 18].Formula = formula;
                            worksheet.Cells[startRowIndex + 1, 18].Style.Numberformat.Format = "0.00%";
                        }
                        else
                        {
                            worksheet.Cells[startRowIndex + 1, 18].Value = "--";
                        }
                    }
                }

                // 添加表身邊框
                using (var range = worksheet.Cells[startRowIndex + 1, 1, startRowIndex + 1, headers.Count])
                {
                    range.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    range.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    range.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                }

                return package.GetAsByteArray();
            }
        }

        public static byte[] FeatDailyToExcel<T>(List<T> items, Dictionary<string, string> headers, string name, string datestring, int people)
        {
            if (items == null || items.Count == 0)
            {
                throw new ArgumentException("列表不能為 null 或空。", nameof(items));
            }

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add(typeof(T).Name);
                var properties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);

                // 添加合併標題
                worksheet.Cells[1, 1].Value = "國峯租賃股份有限公司(" + name + ") 1+" + (people - 1) + "人     " + datestring;
                worksheet.Cells[1, 1, 1, headers.Count].Merge = true; // 合併儲存格
                worksheet.Cells[1, 1, 1, headers.Count].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center; // 標題居中
                worksheet.Cells[1, 1, 1, headers.Count].Style.Font.Bold = true; // 標題加粗

                // 添加子標題
                int headerIndex = 1;
                foreach (var header in headers)
                {
                    var cell = worksheet.Cells[2, headerIndex++];
                    cell.Value = header.Value;
                    // 設置儲存格底色為淺藍色
                    cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    cell.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                }

                // 添加表頭邊框
                using (var range = worksheet.Cells[1, 1, 2, headers.Count])
                {
                    range.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    range.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    range.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                }

                // 添加表身
                int totalRows = items.Count;
                for (int i = 0; i < totalRows; i++)
                {
                    var item = items[i];
                    int columnIndex = 1;
                    foreach (var headerKey in headers.Keys)
                    {
                        var property = typeof(T).GetProperty(headerKey);
                        if (property != null)
                        {
                            var value = property.GetValue(item);
                            worksheet.Cells[i + 3, columnIndex++].Value = value; // 從第三行開始填寫數據
                        }
                        else
                        {
                            worksheet.Cells[i + 3, columnIndex++].Value = ""; // 如果找不到屬性，設置為空
                        }

                        if (headerKey == "AchievementRate")
                        {
                            int target_quota = (int)worksheet.Cells[i + 3, 17].Value;
                            if (target_quota > 0)
                            {
                                var formula = $"({worksheet.Cells[i + 3, 11].Value} + {worksheet.Cells[i + 3, 12].Value}) / {worksheet.Cells[i + 3, 17].Value}";
                                worksheet.Cells[i + 3, 18].Formula = formula;
                                worksheet.Cells[i + 3, 18].Style.Numberformat.Format = "0.00%";
                            }
                            else
                            {
                                worksheet.Cells[i + 3, 18].Value = "--";
                            }
                        }
                    }

                    // 添加表身邊框
                    using (var range = worksheet.Cells[i + 3, 1, i + 3, headers.Count])
                    {
                        range.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        range.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        range.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    }

                    worksheet.Cells[3, 1].Value = "總計";
                    worksheet.Cells[3, 1, 3, 2].Merge = true; // 合併儲存格
                    worksheet.Cells[3, 1, 3, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center; // 標題居中
                }

                return package.GetAsByteArray();
            }
        }

        public static byte[] FeatDailyToExcelAgain<T>(byte[] existingFileBytes, List<T> items, Dictionary<string, string> headers, string name, string datestring,int people)
        {
            using (var package = new ExcelPackage(new MemoryStream(existingFileBytes)))
            {
                var worksheet = package.Workbook.Worksheets[0];

                int existingRowCount = worksheet.Dimension.Rows; // 抓現有的行數
                int startRowIndex = existingRowCount + 2; // 留出一行空白

                // 添加合併標題
                worksheet.Cells[startRowIndex, 1].Value = "國峯租賃股份有限公司(" + name + ") 1+" + (people-1) + "人     " + datestring;
                worksheet.Cells[startRowIndex, 1, startRowIndex, headers.Count].Merge = true; // 合併儲存格
                worksheet.Cells[startRowIndex, 1, startRowIndex, headers.Count].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center; // 標題居中
                worksheet.Cells[startRowIndex, 1, startRowIndex, headers.Count].Style.Font.Bold = true; // 標題加粗

                // 添加子標題
                startRowIndex++;
                int headerIndex = 1;
                foreach (var header in headers)
                {
                    var cell = worksheet.Cells[startRowIndex, headerIndex++];
                    cell.Value = header.Value;
                    // 設置儲存格底色為淺藍色
                    cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    cell.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                }

                // 添加表頭邊框
                using (var range = worksheet.Cells[startRowIndex - 1, 1, startRowIndex, headers.Count])
                {
                    range.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    range.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    range.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                }

                // 添加表身
                int totalRows = items.Count;
                for (int i = 0; i < totalRows; i++)
                {
                    var item = items[i];
                    int columnIndex = 1;
                    foreach (var headerKey in headers.Keys)
                    {
                        var property = typeof(T).GetProperty(headerKey);
                        if (property != null)
                        {
                            var value = property.GetValue(item);
                            worksheet.Cells[startRowIndex + i + 1, columnIndex++].Value = value; // 填入数据
                        }
                        else
                        {
                            worksheet.Cells[startRowIndex + i + 1, columnIndex++].Value = "";
                        }

                        if (headerKey == "AchievementRate")
                        {
                            int target_quota = (int)worksheet.Cells[startRowIndex + i + 1, 17].Value;
                            if (target_quota > 0)
                            {
                                var formula = $"({worksheet.Cells[startRowIndex + i + 1, 11].Value} + {worksheet.Cells[startRowIndex + i + 1, 12].Value}) / {worksheet.Cells[startRowIndex + i + 1, 17].Value}";
                                worksheet.Cells[startRowIndex + i + 1, 18].Formula = formula;
                                worksheet.Cells[startRowIndex + i + 1, 18].Style.Numberformat.Format = "0.00%";
                            }
                            else
                            {
                                worksheet.Cells[startRowIndex + i + 1, 18].Value = "--";
                            }
                        }
                    }

                    // 添加表身邊框
                    using (var range = worksheet.Cells[startRowIndex + i + 1, 1, startRowIndex + i + 1, headers.Count])
                    {
                        range.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        range.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        range.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    }

                    worksheet.Cells[startRowIndex + 1, 1].Value = "總計";
                    worksheet.Cells[startRowIndex + 1, 1, startRowIndex + 1, 2].Merge = true; // 合併儲存格
                    worksheet.Cells[startRowIndex + 1, 1, startRowIndex + 1, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center; // 標題居中
                }

                return package.GetAsByteArray();
            }
        }
        
        /// <summary>
        /// 民國年月日換成西元
        /// </summary>
        /// <param name="rocDate">113/9/25</param>
        public static string ConvertROCToGregorian(string rocDate)
        {
            var parts = rocDate.Split('/');
            int rocYear = int.Parse(parts[0]);
            int month = int.Parse(parts[1]);
            int day = int.Parse(parts[2]);

            int gregorianYear = rocYear + 1911;

            string formattedMonth = month.ToString("D2");
            string formattedDay = day.ToString("D2");

            string gregorianDate = $"{gregorianYear}/{formattedMonth}/{formattedDay}";

            return gregorianDate;
        }

        public static byte[] ApprovalLoanSalesExcelFooter(byte[] existingFileBytes)
        {
            using (var package = new ExcelPackage(new MemoryStream(existingFileBytes)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int existingRowCount = worksheet.Dimension.Rows;//抓現有的行數
                int startRowIndex = existingRowCount + 1;


                worksheet.Cells[startRowIndex, 1].Value = "總計:" + (existingRowCount-1);
                worksheet.Cells[startRowIndex, 1, startRowIndex, 10].Merge = true; // 合併儲存格

                worksheet.Cells[startRowIndex, 11].Value = "合計:";

                var formula12 = $"SUM({worksheet.Cells[2, 12].Address}:{worksheet.Cells[existingRowCount, 12].Address})";
                worksheet.Cells[startRowIndex, 12].Formula = formula12;
                worksheet.Cells[startRowIndex, 12].Style.Numberformat.Format = "#,##0";

                var formula17 = $"SUM({worksheet.Cells[2, 17].Address}:{worksheet.Cells[existingRowCount, 17].Address})";
                worksheet.Cells[startRowIndex, 17].Formula = formula17;
                worksheet.Cells[startRowIndex, 17].Style.Numberformat.Format = "#,##0";

                var formula18 = $"SUM({worksheet.Cells[2, 18].Address}:{worksheet.Cells[existingRowCount, 18].Address})";
                worksheet.Cells[startRowIndex, 18].Formula = formula18;
                worksheet.Cells[startRowIndex, 18].Style.Numberformat.Format = "#,##0";

                var formula19 = $"SUM({worksheet.Cells[2, 19].Address}:{worksheet.Cells[existingRowCount, 19].Address})";
                worksheet.Cells[startRowIndex, 19].Formula = formula19;
                worksheet.Cells[startRowIndex, 19].Style.Numberformat.Format = "#,##0";
                
                var formula20 = $"SUM({worksheet.Cells[2, 20].Address}:{worksheet.Cells[existingRowCount, 20].Address})";
                worksheet.Cells[startRowIndex, 20].Formula = formula20;
                worksheet.Cells[startRowIndex, 20].Style.Numberformat.Format = "#,##0";
                
                using (var range = worksheet.Cells[startRowIndex, 1, startRowIndex, 23])
                {
                    range.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    range.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    range.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                }
                return package.GetAsByteArray();
            }
        }
    }
}
