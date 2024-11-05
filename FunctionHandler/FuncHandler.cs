using KF_WebAPI.BaseClass;
using KF_WebAPI.BaseClass.AE;
using OfficeOpenXml;
using System.Data;
using Microsoft.Data.SqlClient;
using System.Drawing;
using System.Reflection;
using System.Reflection.PortableExecutable;
using Microsoft.AspNetCore.Http;
using System.Globalization;
using System;
using Microsoft.Identity.Client;

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
        /// <summary>
        /// 西元年月日換成民國
        /// </summary>
        /// <param name="rocDate">2024/09/25</param>
        public static string ConvertGregorianToROC(string rocDate)
        {
            var parts = rocDate.Split('/');
            int rocYear = int.Parse(parts[0]);
            int month = int.Parse(parts[1]);
            int day = int.Parse(parts[2]);

            int gregorianYear = rocYear - 1911;

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

        public static byte[] CustomerQtyCountExcel(List<customer_qty_count_Excel> items, Dictionary<string, string> headers, customer_qty_count_req model)
        {
            if (items == null || items.Count == 0)
            {
                throw new ArgumentException("列表不能為 null 或空。", nameof(items));
            }

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Sheet0");

                // 添加合併標題
                worksheet.Cells[1, 1].Value = "報表日期區間" + model.Qty_Date_S + " ~ " + model.Qty_Date_E;
                worksheet.Cells[1, 1, 1, headers.Count + 1].Merge = true;
                worksheet.Cells[1, 1, 1, headers.Count + 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                worksheet.Cells[1, 1, 1, headers.Count + 1].Style.Font.Bold = true;

                // 添加表頭行
                int colIndex = 1;
                worksheet.Cells[2, colIndex++].Value = "序號";
                foreach (var header in headers)
                {
                    worksheet.Cells[2, colIndex++].Value = header.Value;
                }

                int rowIndex = 3; // 從第三行開始填充數據
                var groupedItems = items.GroupBy(x => x.com_name);
                int serialNumber = 1; 

                int subtotal = 0;

                // 先填充台北的數據
                foreach (var group in groupedItems)
                {
                    if (group.Key == "台北")
                    {
                        foreach (var item in group)
                        {
                            colIndex = 1;
                            worksheet.Cells[rowIndex, colIndex++].Value = serialNumber++; // 填充序號
                            foreach (var header in headers)
                            {
                                var propertyValue = item.GetType().GetProperty(header.Key)?.GetValue(item);
                                worksheet.Cells[rowIndex, colIndex++].Value = propertyValue;
                            }
                            subtotal += item.count; // 累加預估件數
                            rowIndex++;
                        }
                    }
                }

                // 填充數位行銷部的數據
                foreach (var group in groupedItems)
                {
                    if (group.Key == "數位行銷部")
                    {
                        foreach (var item in group)
                        {
                            colIndex = 1;
                            worksheet.Cells[rowIndex, colIndex++].Value = serialNumber++; // 填充序號
                            foreach (var header in headers)
                            {
                                var propertyValue = item.GetType().GetProperty(header.Key)?.GetValue(item);
                                worksheet.Cells[rowIndex, colIndex++].Value = propertyValue;
                            }
                            subtotal += item.count; // 累加預估件數
                            rowIndex++;
                        }
                    }
                }

                // 添加小計行
                worksheet.Cells[rowIndex, 1].Value = "台北 小計";
                worksheet.Cells[rowIndex, 1, rowIndex, colIndex - 2].Merge = true;
                worksheet.Cells[rowIndex, colIndex - 1].Value = subtotal;
                rowIndex++;

                // 填充其他部門
                foreach (var group in groupedItems)
                {
                    if (group.Key != "台北" && group.Key != "數位行銷部")
                    {
                        subtotal = group.Sum(x => x.count); // 計算該組的總和
                        foreach (var item in group)
                        {
                            colIndex = 1;
                            worksheet.Cells[rowIndex, colIndex++].Value = serialNumber++; // 填充序號
                            foreach (var header in headers)
                            {
                                var propertyValue = item.GetType().GetProperty(header.Key)?.GetValue(item);
                                worksheet.Cells[rowIndex, colIndex++].Value = propertyValue;
                            }
                            rowIndex++;
                        }

                        // 添加該分公司的小計行
                        worksheet.Cells[rowIndex, 1].Value = group.Key + " 小計";
                        worksheet.Cells[rowIndex, 1, rowIndex, colIndex - 2].Merge = true;
                        worksheet.Cells[rowIndex, colIndex - 1].Value = subtotal;
                        rowIndex++;
                    }
                }

                // 添加總計行
                worksheet.Cells[rowIndex, 1].Value = "總計";
                worksheet.Cells[rowIndex, 1, rowIndex, colIndex - 2].Merge = true;
                worksheet.Cells[rowIndex, colIndex - 1].Value = items.Sum(x => x.count);

                // 添加框線
                var range = worksheet.Cells[1, 1, rowIndex, headers.Count + 1];
                range.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                range.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                range.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                // 自動調整列寬
                worksheet.Cells.AutoFitColumns();

                return package.GetAsByteArray();
            }
        }

        public static byte[] IncomingPartsExcel(List<Incoming_Part_res> itemsA, List<Incoming_Part_res> itemsB, string[] daysArray, Incoming_parts_req req)
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Sheet0");
                var leagh = daysArray.Count();
                string[] headers = new string[] { "序號", "區域", "職稱", "業務", "統計", "工作天數" };
                DateTime date;

                // 添加合併標題
                worksheet.Cells[1, 1].Value = "進件預估核准撥款表 報表日期區間" + req.Inc_Date_S + " ~ " + req.Inc_Date_E;
                worksheet.Cells[1, 1, 1, leagh + 6].Merge = true;
                worksheet.Cells[1, 1, 1, leagh + 6].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                worksheet.Cells[1, 1, 1, leagh + 6].Style.Font.Bold = true;

                // 添加表頭行
                int colIndex = 1;
                foreach (var header in headers)
                {
                    worksheet.Cells[2, colIndex++].Value = header;
                }
                foreach (var strday in daysArray)
                {
                    if (DateTime.TryParse(strday, out date))
                    {
                        string formattedDate = date.ToString("M月d日", CultureInfo.InvariantCulture);
                        worksheet.Cells[2, colIndex++].Value = formattedDate;
                    }
                }

                //添加表身
                int serialNumber = 1;
                int rowIndex = 3;
                int subtotal = 0;
                var groupedItemsA = itemsA.GroupBy(x => x.com_name);
                foreach (var group in groupedItemsA)
                {
                    subtotal = group.Sum(x => x.totalcount);
                    foreach (var item in group)
                    {
                        colIndex = 1;
                        worksheet.Cells[rowIndex, colIndex++].Value = serialNumber++; // 填充序號
                        worksheet.Cells[rowIndex, colIndex++].Value = item.com_name;
                        worksheet.Cells[rowIndex, colIndex++].Value = item.ti_name;
                        worksheet.Cells[rowIndex, colIndex++].Value = item.U_name;
                        worksheet.Cells[rowIndex, colIndex++].Value = item.totalcount;
                        worksheet.Cells[rowIndex, colIndex++].Value = leagh;
                        foreach (var day in item.DateValues)
                        {
                            worksheet.Cells[rowIndex, colIndex++].Value = day.Value;
                        }
                        rowIndex++;
                    }
                    // 添加該分公司的小計行
                    worksheet.Cells[rowIndex, 1].Value = group.Key + " 小計";
                    worksheet.Cells[rowIndex, 1, rowIndex, 4].Merge = true;
                    worksheet.Cells[rowIndex, 5].Value = subtotal;
                    worksheet.Cells[rowIndex, 6, rowIndex, leagh + 6].Merge = true;
                    rowIndex++;
                }

                //添加總計行
                worksheet.Cells[rowIndex, 1].Value = "總計";
                worksheet.Cells[rowIndex, 1, rowIndex, 4].Merge = true;
                worksheet.Cells[rowIndex, 5].Value = itemsA.Sum(x => x.totalcount);
                worksheet.Cells[rowIndex, 6, rowIndex, leagh + 6].Merge = true;

                rowIndex++;
                rowIndex++;

                // 添加合併標題
                worksheet.Cells[rowIndex, 1].Value = "業務件數統計表 報表日期區間" + req.Inc_Date_S + " ~ " + req.Inc_Date_E;
                worksheet.Cells[rowIndex, 1, rowIndex, leagh + 6].Merge = true;
                worksheet.Cells[rowIndex, 1, rowIndex, leagh + 6].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                worksheet.Cells[rowIndex, 1, rowIndex, leagh + 6].Style.Font.Bold = true;
                rowIndex++;

                // 添加表頭行
                colIndex = 1;
                int serialNumberB = 1;
                foreach (var header in headers)
                {
                    worksheet.Cells[rowIndex, colIndex++].Value = header;
                }
                foreach (var strday in daysArray)
                {
                    if (DateTime.TryParse(strday, out date))
                    {
                        string formattedDate = date.ToString("M月d日", CultureInfo.InvariantCulture);
                        worksheet.Cells[rowIndex, colIndex++].Value = formattedDate;
                    }
                }
                rowIndex++;

                //添加表身
                var groupedItemsB = itemsB.GroupBy(x => x.com_name);
                foreach (var group in groupedItemsB)
                {
                    subtotal = group.Sum(x => x.totalcount);
                    foreach (var item in group)
                    {
                        colIndex = 1;
                        worksheet.Cells[rowIndex, colIndex++].Value = serialNumberB++; // 填充序號
                        worksheet.Cells[rowIndex, colIndex++].Value = item.com_name;
                        worksheet.Cells[rowIndex, colIndex++].Value = item.ti_name;
                        worksheet.Cells[rowIndex, colIndex++].Value = item.U_name;
                        worksheet.Cells[rowIndex, colIndex++].Value = item.totalcount;
                        worksheet.Cells[rowIndex, colIndex++].Value = leagh;
                        foreach (var day in item.DateValues)
                        {
                            worksheet.Cells[rowIndex, colIndex++].Value = day.Value;
                        }
                        rowIndex++;
                    }
                    // 添加該分公司的小計行
                    worksheet.Cells[rowIndex, 1].Value = group.Key + " 小計";
                    worksheet.Cells[rowIndex, 1, rowIndex, 4].Merge = true;
                    worksheet.Cells[rowIndex, 5].Value = subtotal;
                    worksheet.Cells[rowIndex, 6, rowIndex, leagh + 6].Merge = true;
                    rowIndex++;
                }

                //添加總計行
                worksheet.Cells[rowIndex, 1].Value = "總計";
                worksheet.Cells[rowIndex, 1, rowIndex, 4].Merge = true;
                worksheet.Cells[rowIndex, 5].Value = itemsB.Sum(x => x.totalcount);
                worksheet.Cells[rowIndex, 6, rowIndex, leagh + 6].Merge = true;

                // 自動調整列寬
                worksheet.Cells.AutoFitColumns();
                return package.GetAsByteArray();
            }
        }

        public static byte[] AttendanceExcelAllMonth(List<Attendance_report_excel> modelList,string yyyy,string mm, List<Flow_rest_HR_excel> flowRestList)
        {
            using (var package = new ExcelPackage())
            {
                #region 各公司打卡資料
                var bcOrder = new List<string> { "BC0800", "BC0900", "BC0100", "BC0200", "BC0600", "BC0300", "BC0500", "BC0400", "BC0700" };
                var bcGroups = modelList.GroupBy(x => x.U_BC).OrderBy(g => bcOrder.IndexOf(g.Key)).ToList();
                string[] headers = { "名稱", "日期", "上班", "下班", "遲到", "外出時間" };
             
                foreach (var bcGroup in bcGroups)
                {
                    ExcelWorksheet worksheet;
                    switch (bcGroup.Key)
                    {
                        case "BC0800":
                            worksheet = package.Workbook.Worksheets.Add("總公司");
                            break;
                        case "BC0900":
                            worksheet = package.Workbook.Worksheets.Add("行銷部");
                            break;
                        case "BC0100":
                            worksheet = package.Workbook.Worksheets.Add("台北");
                            break;
                        case "BC0200":
                            worksheet = package.Workbook.Worksheets.Add("板橋");
                            break;
                        case "BC0600":
                            worksheet = package.Workbook.Worksheets.Add("桃園");
                            break;
                        case "BC0300":
                            worksheet = package.Workbook.Worksheets.Add("台中");
                            break;
                        case "BC0500":
                            worksheet = package.Workbook.Worksheets.Add("台南");
                            break;
                        case "BC0400":
                            worksheet = package.Workbook.Worksheets.Add("高雄");
                            break;
                        case "BC0700":
                            worksheet = package.Workbook.Worksheets.Add("湧立");
                            break;
                        default:
                            throw new InvalidOperationException("未知的 U_BC: " + bcGroup.Key);
                    }

                    int intcount = 0;
                    int rowIndex = 1;
                    int colIndex = 1;

                    foreach (var userIDGroup in bcGroup.GroupBy(x => x.userID))
                    {
                        if (intcount > 0 && intcount % 4 == 0)
                        {
                            rowIndex += 33;
                            intcount = 0;
                        }

                        // 寫入標題
                        for (int i = 0; i < headers.Length; i++)
                        {
                            worksheet.Cells[rowIndex, colIndex + (intcount * 6) + i].Value = headers[i];
                        }
                        worksheet.Cells[rowIndex + 1, colIndex + (intcount * 6)].Value = $"{userIDGroup.First().userID}: {userIDGroup.First().user_name}";

                        var items = userIDGroup.OrderBy(x => x.attendance_date).ToList();

                        int j;
                        for (j = 0; j < 32; j++)
                        {
                            if (j < items.Count)
                            {
                                worksheet.Cells[rowIndex + j + 1, colIndex + (intcount * 6) + 1].Value = items[j].attendance_week;
                                worksheet.Cells[rowIndex + j + 1, colIndex + (intcount * 6) + 2].Value = items[j].work_time;
                                worksheet.Cells[rowIndex + j + 1, colIndex + (intcount * 6) + 3].Value = items[j].getoffwork_time;
                                //worksheet.Cells[rowIndex + j + 1, colIndex + (intcount * 6) + 4].Value = items[j].Late;
                                //worksheet.Cells[rowIndex + j + 1, colIndex + (intcount * 6) + 5].Value = items[j].typename;
                                #region 請假事由判定(颱風假、請假、到職日、離職日、早退
                                DateTime attendanceDate;
                                DateTime.TryParse(items[j].attendance_date, out attendanceDate);
                                bool isRest = false;
                                //颱風假
                                if (items[j].type== "Hk_04")
                                {
                                    isRest = true;
                                    items[j].typename += " ";
                                }
                                //請假
                                if (flowRestList.Any(x => x.FR_U_num == items[j].userID && attendanceDate.Date >= x.FR_Date_S.Date
                                && attendanceDate.Date <= x.FR_Date_E.Date && x.FR_kind != "FRK021"))
                                {
                                    isRest = true;

                                    var RestList = flowRestList.Where(x => x.FR_U_num.Equals(items[j].userID) && attendanceDate.Date >= x.FR_Date_S.Date
                                    && attendanceDate.Date <= x.FR_Date_E.Date).ToList();

                                    if( RestList.Count > 0)
                                    {
                                        foreach( var rest in RestList) 
                                        {
                                            string displayValue = (rest.FR_total_hour % 1 == 0) ? ((int)rest.FR_total_hour).ToString() : rest.FR_total_hour.ToString();
                                            if (rest.FR_kind == "FRK017") //忘打卡FRK017
                                            {
                                                if(rest.FR_Date_S.Hour == 9)
                                                {
                                                    items[j].typename += rest.FR_Kind_name + " 09:00上班";
                                                    items[j].Late = 0;
                                                }
                                                else
                                                {
                                                    items[j].typename += rest.FR_Kind_name + " 18:00下班";
                                                    items[j].early = 0;
                                                }
                                            }
                                            else if(rest.FR_kind == "FRK016") //外出 FRK016
                                            {
                                                items[j].typename += rest.FR_note;
                                                items[j].Late = 0;
                                                items[j].early = 0;
                                            }
                                            else if(rest.FR_total_hour >= 8)
                                            {
                                                items[j].typename += rest.FR_Kind_name + "8H ";
                                                items[j].Late = 0;
                                                items[j].early = 0;
                                            }
                                            else
                                            {
                                                items[j].typename += rest.FR_Kind_name + displayValue + "H " + rest.FR_Date_S.ToString("HH:mm") + " ~ " + rest.FR_Date_E.ToString("HH:mm");
                                                var totalHour = Convert.ToInt32(rest.FR_total_hour * 60);
                                                if (!string.IsNullOrEmpty(items[j].work_time) && DateTime.Parse(items[j].work_time).TimeOfDay > TimeSpan.Parse("09:00")
                                                    && items[j].Late > 0 && rest.FR_Date_E.TimeOfDay != TimeSpan.Parse("18:00"))
                                                {
                                                    items[j].Late = Math.Max(0, Convert.ToInt32(items[j].Late) - totalHour);
                                                    continue;
                                                }
                                                if (!string.IsNullOrEmpty(items[j].getoffwork_time) && DateTime.Parse(items[j].getoffwork_time).TimeOfDay <= TimeSpan.Parse("18:00"))
                                                {
                                                    items[j].early = Math.Max(0, Convert.ToInt32(items[j].early) - totalHour);
                                                }
                                            }
                                        }
                                    }
                                }
                                //整天曠職
                                if (!isRest)
                                {
                                    if (string.IsNullOrEmpty(items[j].work_time) && !items[j].attendance_week.Contains("*")
                                        && items[j].arrive_date.Value.Date <= attendanceDate.Date
                                        && attendanceDate <= DateTime.Now.Date)
                                    {
                                        if (!(items[j].leave_date.HasValue && items[j].leave_date.Value.Date <= attendanceDate.Date))
                                        {
                                            items[j].absenteeism = "Y";
                                            items[j].typename += "曠職8H";
                                            items[j].early = 480;
                                            isRest = true;
                                        }
                                    }
                                }
                                //早退
                                if (!isRest)
                                {
                                    if(items[j].early > 0)
                                    {
                                        decimal earlyHour = Convert.ToInt32(items[j].early) / 60m;
                                        decimal calculatedEarlyHour = Math.Ceiling(earlyHour / 0.5m) * 0.5m;
                                        if (calculatedEarlyHour > 8)
                                            calculatedEarlyHour = 8;
                                        items[j].typename += "曠職" + calculatedEarlyHour + " H";
                                    }
                                }
                                //到職日
                                if (items[j].arrive_date.HasValue && items[j].attendance_date == items[j].arrive_date.Value.ToString("yyyy/MM/dd"))
                                {
                                    items[j].typename += "到職日";
                                    items[j].Late = 0;
                                    items[j].early = 0;
                                    worksheet.Cells[rowIndex + j + 1, colIndex + (intcount * 6) + 4].Value = 0;
                                }
                                //離職日
                                if (items[j].leave_date.HasValue && items[j].attendance_date == items[j].leave_date.Value.ToString("yyyy/MM/dd"))
                                    items[j].typename += "離職日";

                                //先將早退的值加入到遲到(暫定)
                                items[j].Late = items[j].Late + items[j].early;
                                worksheet.Cells[rowIndex + j + 1, colIndex + (intcount * 6) + 4].Value = items[j].Late;
                                worksheet.Cells[rowIndex + j + 1, colIndex + (intcount * 6) + 5].Value = items[j].typename;
                                worksheet.Cells[rowIndex + j + 1, colIndex + (intcount * 6) + 5].Style.Font.Color.SetColor(Color.Red);
                                #endregion
                            }
                            else
                            {
                                for (int col = 1; col <= 5; col++)
                                {
                                    worksheet.Cells[rowIndex + j + 1, colIndex + (intcount * 6) + col].Value = string.Empty;
                                }
                            }
                        }
                        var lateTotal = Convert.ToInt32(userIDGroup.Sum(x => x.Late) - 15);
                        worksheet.Cells[rowIndex + j, colIndex + (intcount * 6)].Value = "合計";
                        worksheet.Cells[rowIndex + j, colIndex + (intcount * 6) + 1].Value = "扣";
                        decimal Hour = 0;
                        if (lateTotal > 0)
                            Hour = lateTotal / 60m;
                        worksheet.Cells[rowIndex + j, colIndex + (intcount * 6) + 2].Value = Math.Ceiling(Hour / 0.5m) * 0.5m;
                        worksheet.Cells[rowIndex + j, colIndex + (intcount * 6) + 3].Value = "小時";
                        worksheet.Cells[rowIndex + j, colIndex + (intcount * 6) + 4].Value = lateTotal;

                        worksheet.Cells[rowIndex + j, colIndex + (intcount * 6) + 1].Style.Font.Color.SetColor(Color.Red);
                        worksheet.Cells[rowIndex + j, colIndex + (intcount * 6) + 2].Style.Font.Color.SetColor(Color.Red);
                        worksheet.Cells[rowIndex + j, colIndex + (intcount * 6) + 3].Style.Font.Color.SetColor(Color.Red);
                        worksheet.Cells[rowIndex + j, colIndex + (intcount * 6) + 4].Style.Font.Color.SetColor(Color.Red);

                        // 設置框線
                        var range = worksheet.Cells[rowIndex, colIndex + (intcount * 6), rowIndex + j, colIndex + (intcount * 6) + 5];
                        range.Style.Border.Top.Style = range.Style.Border.Bottom.Style =
                        range.Style.Border.Left.Style = range.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                        intcount++;
                    }

                    // 自動調整列寬
                    worksheet.Cells.AutoFitColumns();
                }
                #endregion

                #region 遲到、曠職
                var worksheet_ly = package.Workbook.Worksheets.Add("遲到、曠職");

                // 添加合併標題 遲到
                worksheet_ly.Cells[1, 1].Value = Convert.ToInt32(yyyy) - 1911 + "年" + mm + "月  國峯遲到";
                worksheet_ly.Cells[1, 1, 1, 4].Merge = true;
                worksheet_ly.Cells[1, 1, 1, 4].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                worksheet_ly.Cells[1, 1, 1, 4].Style.Font.Bold = true;

                worksheet_ly.Cells[2, 1].Value = "名稱";
                worksheet_ly.Cells[2, 4].Value = "遲到";

                int rowindex_ly = 2;
                int colindex_ly = 1;
                var userResult = modelList.GroupBy(x => x.userID).Select(g => new { UserID = g.Key, Totalval = g.Sum(x => x.Late)-15 }).OrderBy(x=>x.UserID);
                foreach ( var user in userResult)
                {
                    if(user.Totalval > 0)
                    {
                        var name = modelList.Where(x => x.userID.Equals(user.UserID)).FirstOrDefault().user_name;
                        rowindex_ly++;
                        worksheet_ly.Cells[rowindex_ly, colindex_ly++].Value = user.UserID + ":" + name;
                        worksheet_ly.Cells[rowindex_ly, colindex_ly++].Value = "合計";
                        worksheet_ly.Cells[rowindex_ly, colindex_ly++].Value = "扣";
                        decimal lateHour = Convert.ToInt32(user.Totalval) / 60m;
                        worksheet_ly.Cells[rowindex_ly, colindex_ly++].Value = Math.Ceiling(lateHour / 0.5m) * 0.5m + "小時";
                        colindex_ly = 1;
                    }
                }
                // 添加框線
                var range_ly1 = worksheet_ly.Cells[1, 1, rowindex_ly, 4];
                range_ly1.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                range_ly1.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                range_ly1.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                range_ly1.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                #region 曠職註解
                //// 添加合併標題 曠職
                //worksheet_ly.Cells[1, 6].Value = Convert.ToInt32(yyyy) - 1911 + "年" + mm + "月  國峯曠職";
                //worksheet_ly.Cells[1, 6, 1, 10].Merge = true;
                //worksheet_ly.Cells[1, 6, 1, 10].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                //worksheet_ly.Cells[1, 6, 1, 10].Style.Font.Bold = true;

                //worksheet_ly.Cells[2, 6].Value = "名稱";
                //worksheet_ly.Cells[2, 7].Value = "日期";
                //worksheet_ly.Cells[2, 10].Value = "曠職";

                //rowindex_ly = 2;
                //colindex_ly = 6;

                //foreach (var item in modelList.Where(x => x.early > 0).OrderBy(x => x.userID))
                //{
                //    rowindex_ly++;
                //    worksheet_ly.Cells[rowindex_ly, colindex_ly++].Value = $"{item.userID}: {item.user_name}";
                //    worksheet_ly.Cells[rowindex_ly, colindex_ly++].Value = item.attendance_week;
                //    colindex_ly++;
                //    colindex_ly++;
                //    decimal earlyHour = Convert.ToInt32(item.early) / 60m;
                //    worksheet_ly.Cells[rowindex_ly, colindex_ly++].Value = Math.Ceiling(earlyHour / 0.5m) * 0.5m + "H";
                //    colindex_ly = 6;
                //}

                //// 添加框線
                //var range_ly2 = worksheet_ly.Cells[1, 6, rowindex_ly, 10];
                //range_ly2.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                //range_ly2.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                //range_ly2.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                //range_ly2.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                #endregion

                //颱風假&&無請假紀錄&&在職
                if (modelList?.Any(x => x.type != null && x.type.Equals("Hk_04")) == true)
                {
                    // 添加合併標題 曠職
                    worksheet_ly.Cells[1, 6].Value = Convert.ToInt32(yyyy) - 1911 + "年" + mm + "月  國峯颱風假";
                    worksheet_ly.Cells[1, 6, 1, 7].Merge = true;
                    worksheet_ly.Cells[1, 6, 1, 7].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    worksheet_ly.Cells[1, 6, 1, 7].Style.Font.Bold = true;

                    worksheet_ly.Cells[2, 6].Value = "名稱";
                    worksheet_ly.Cells[2, 7].Value = "日期";

                    rowindex_ly = 2;
                    colindex_ly = 6;

                    foreach (var bcGroup in modelList.Where(x => x.type != null && x.type.Equals("Hk_04")).GroupBy(x => x.U_BC).OrderBy(g => bcOrder.IndexOf(g.Key)).ToList())
                    {
                        rowindex_ly++;
                        switch (bcGroup.Key)
                        {
                            case "BC0800":
                                worksheet_ly.Cells[rowindex_ly, 6].Value = "總公司";
                                worksheet_ly.Cells[rowindex_ly, 6, rowindex_ly, 7].Merge = true;
                                worksheet_ly.Cells[rowindex_ly, 6, rowindex_ly, 7].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                worksheet_ly.Cells[rowindex_ly, 6, rowindex_ly, 7].Style.Font.Bold = true;
                                break;
                            case "BC0900":
                                worksheet_ly.Cells[rowindex_ly, 6].Value = "行銷部";
                                worksheet_ly.Cells[rowindex_ly, 6, rowindex_ly, 7].Merge = true;
                                worksheet_ly.Cells[rowindex_ly, 6, rowindex_ly, 7].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                worksheet_ly.Cells[rowindex_ly, 6, rowindex_ly, 7].Style.Font.Bold = true;
                                break;
                            case "BC0100":
                                worksheet_ly.Cells[rowindex_ly, 6].Value = "台北";
                                worksheet_ly.Cells[rowindex_ly, 6, rowindex_ly, 7].Merge = true;
                                worksheet_ly.Cells[rowindex_ly, 6, rowindex_ly, 7].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                worksheet_ly.Cells[rowindex_ly, 6, rowindex_ly, 7].Style.Font.Bold = true;
                                break;
                            case "BC0200":
                                worksheet_ly.Cells[rowindex_ly, 6].Value = "板橋";
                                worksheet_ly.Cells[rowindex_ly, 6, rowindex_ly, 7].Merge = true;
                                worksheet_ly.Cells[rowindex_ly, 6, rowindex_ly, 7].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                worksheet_ly.Cells[rowindex_ly, 6, rowindex_ly, 7].Style.Font.Bold = true;
                                break;
                            case "BC0600":
                                worksheet_ly.Cells[rowindex_ly, 6].Value = "桃園";
                                worksheet_ly.Cells[rowindex_ly, 6, rowindex_ly, 7].Merge = true;
                                worksheet_ly.Cells[rowindex_ly, 6, rowindex_ly, 7].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                worksheet_ly.Cells[rowindex_ly, 6, rowindex_ly, 7].Style.Font.Bold = true;
                                break;
                            case "BC0300":
                                worksheet_ly.Cells[rowindex_ly, 6].Value = "台中";
                                worksheet_ly.Cells[rowindex_ly, 6, rowindex_ly, 7].Merge = true;
                                worksheet_ly.Cells[rowindex_ly, 6, rowindex_ly, 7].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                worksheet_ly.Cells[rowindex_ly, 6, rowindex_ly, 7].Style.Font.Bold = true;
                                break;
                            case "BC0500":
                                worksheet_ly.Cells[rowindex_ly, 6].Value = "台南";
                                worksheet_ly.Cells[rowindex_ly, 6, rowindex_ly, 7].Merge = true;
                                worksheet_ly.Cells[rowindex_ly, 6, rowindex_ly, 7].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                worksheet_ly.Cells[rowindex_ly, 6, rowindex_ly, 7].Style.Font.Bold = true;
                                break;
                            case "BC0400":
                                worksheet_ly.Cells[rowindex_ly, 6].Value = "高雄";
                                worksheet_ly.Cells[rowindex_ly, 6, rowindex_ly, 7].Merge = true;
                                worksheet_ly.Cells[rowindex_ly, 6, rowindex_ly, 7].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                worksheet_ly.Cells[rowindex_ly, 6, rowindex_ly, 7].Style.Font.Bold = true;
                                break;
                            case "BC0700":
                                worksheet_ly.Cells[rowindex_ly, 6].Value = "湧立";
                                worksheet_ly.Cells[rowindex_ly, 6, rowindex_ly, 7].Merge = true;
                                worksheet_ly.Cells[rowindex_ly, 6, rowindex_ly, 7].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                worksheet_ly.Cells[rowindex_ly, 6, rowindex_ly, 7].Style.Font.Bold = true;
                                break;
                        }

                        foreach (var item in bcGroup.Where(x => x.type != null && x.type.Equals("Hk_04")).OrderBy(x => x.attendance_date))
                        {
                            DateTime attendanceDate;
                            DateTime.TryParse(item.attendance_date, out attendanceDate);

                            if (item.arrive_date.HasValue && item.arrive_date <= attendanceDate)
                            {
                                if (!(item.leave_date.HasValue && item.leave_date < attendanceDate))
                                {
                                    if (!flowRestList.Any(x => x.FR_U_num == item.userID && attendanceDate.Date >= x.FR_Date_S.Date
                                    && attendanceDate.Date <= x.FR_Date_E.Date))
                                    {
                                        //除(國峯助理&&業務&&業務主管)外只要有來就不會扣颱風假
                                        if (!(item.Role_num != "1008" && item.Role_num != "1009" && item.Role_num != "1011" && !string.IsNullOrEmpty(item.work_time)))
                                        {
                                            rowindex_ly++;
                                            worksheet_ly.Cells[rowindex_ly, colindex_ly++].Value = $"{item.userID}: {item.user_name}";
                                            worksheet_ly.Cells[rowindex_ly, colindex_ly++].Value = item.attendance_week;
                                            colindex_ly = 6;
                                        }
                                    }
                                }
                            }
                        }
                    }

                    // 添加框線
                    var range_ly2 = worksheet_ly.Cells[1, 6, rowindex_ly, 7];
                    range_ly2.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    range_ly2.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    range_ly2.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    range_ly2.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                }

                // 自動調整列寬
                worksheet_ly.Cells.AutoFitColumns();
                #endregion

                #region 請假
                var worksheet_af = package.Workbook.Worksheets.Add("請假");
                string[] headers_af = new string[] { "序號", "假單編號", "姓名", "假別", "請假起迄日", "請假時數", "簽核結果" };

                int colIndex_af = 1;
                int rowIndex_af = 1;
                foreach (var header in headers_af)
                {
                    worksheet_af.Cells[rowIndex_af, colIndex_af++].Value = header;
                }
                rowIndex_af++;
                int iaf = 1;
                foreach (var item in flowRestList.Where(x=>x.FR_kind != "FRK016" && x.FR_kind != "FRK017" && x.FR_kind != "FRK021").OrderBy(x=>x.FR_Date_S).ToList()) 
                {
                    colIndex_af = 1;
                    worksheet_af.Cells[rowIndex_af, colIndex_af].Value = iaf++;
                    worksheet_af.Cells[rowIndex_af, colIndex_af, rowIndex_af + 1, colIndex_af].Merge = true;
                    worksheet_af.Cells[rowIndex_af, colIndex_af, rowIndex_af + 1, colIndex_af].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    worksheet_af.Cells[rowIndex_af, colIndex_af, rowIndex_af + 1, colIndex_af].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                    colIndex_af++;
                    worksheet_af.Cells[rowIndex_af, colIndex_af].Value = item.FR_id;
                    worksheet_af.Cells[rowIndex_af, colIndex_af, rowIndex_af + 1, colIndex_af].Merge = true;
                    worksheet_af.Cells[rowIndex_af, colIndex_af, rowIndex_af + 1, colIndex_af].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    worksheet_af.Cells[rowIndex_af, colIndex_af, rowIndex_af + 1, colIndex_af].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                    colIndex_af++;
                    worksheet_af.Cells[rowIndex_af, colIndex_af].Value = item.U_name;
                    worksheet_af.Cells[rowIndex_af, colIndex_af, rowIndex_af + 1, colIndex_af].Merge = true;
                    worksheet_af.Cells[rowIndex_af, colIndex_af, rowIndex_af + 1, colIndex_af].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    worksheet_af.Cells[rowIndex_af, colIndex_af, rowIndex_af + 1, colIndex_af].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                    colIndex_af++;
                    worksheet_af.Cells[rowIndex_af, colIndex_af].Value = item.FR_Kind_name;
                    worksheet_af.Cells[rowIndex_af, colIndex_af, rowIndex_af + 1, colIndex_af].Merge = true;
                    worksheet_af.Cells[rowIndex_af, colIndex_af, rowIndex_af + 1, colIndex_af].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    worksheet_af.Cells[rowIndex_af, colIndex_af, rowIndex_af + 1, colIndex_af].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                    colIndex_af++;
                    
                    worksheet_af.Cells[rowIndex_af, colIndex_af].Value = item.FR_Date_S.ToString("yyyy-MM-dd HH:mm");
                    worksheet_af.Cells[rowIndex_af + 1, colIndex_af++].Value = item.FR_Date_E.ToString("yyyy-MM-dd HH:mm");

                    worksheet_af.Cells[rowIndex_af, colIndex_af].Value = item.FR_total_hour;
                    worksheet_af.Cells[rowIndex_af, colIndex_af, rowIndex_af + 1, colIndex_af].Merge = true;
                    worksheet_af.Cells[rowIndex_af, colIndex_af, rowIndex_af + 1, colIndex_af].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    worksheet_af.Cells[rowIndex_af, colIndex_af, rowIndex_af + 1, colIndex_af].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                    colIndex_af++;
                    worksheet_af.Cells[rowIndex_af, colIndex_af].Value = item.Sign_name;
                    worksheet_af.Cells[rowIndex_af, colIndex_af, rowIndex_af + 1, colIndex_af].Merge = true;
                    worksheet_af.Cells[rowIndex_af, colIndex_af, rowIndex_af + 1, colIndex_af].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    worksheet_af.Cells[rowIndex_af, colIndex_af, rowIndex_af + 1, colIndex_af].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                    colIndex_af++;
                    rowIndex_af++;
                    rowIndex_af++;
                }

                // 添加框線
                var range_af = worksheet_af.Cells[1, 1, rowIndex_af - 1, colIndex_af - 1];
                range_af.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                range_af.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                range_af.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                range_af.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                // 自動調整列寬
                worksheet_af.Cells.AutoFitColumns();
                #endregion

                #region 加班
                var worksheet_wo = package.Workbook.Worksheets.Add("加班");
                string[] headers_wo = new string[] { "序號", "名稱", "加班起迄日", "加班時數", "簽核結果" };         

                int rowIndex_wo = 1;
                int colIndex_wo = 1;
                
                foreach (var header in headers_wo)
                {
                    worksheet_wo.Cells[rowIndex_wo, colIndex_wo++].Value = header;
                }

                rowIndex_wo++;
                int wo_index = 1;


                foreach (var item in flowRestList.Where(x=>x.FR_kind.Equals("FRK021")).ToList())
                {
                    colIndex_wo = 1;
                    worksheet_wo.Cells[rowIndex_wo, colIndex_wo].Value = wo_index++;
                    worksheet_wo.Cells[rowIndex_wo, colIndex_wo, rowIndex_wo + 1, colIndex_wo].Merge = true;
                    worksheet_wo.Cells[rowIndex_wo, colIndex_wo, rowIndex_wo + 1, colIndex_wo].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    worksheet_wo.Cells[rowIndex_wo, colIndex_wo, rowIndex_wo + 1, colIndex_wo].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                    colIndex_wo++;
                    worksheet_wo.Cells[rowIndex_wo, colIndex_wo].Value = $"{item.FR_U_num}: {item.U_name}";
                    worksheet_wo.Cells[rowIndex_wo, colIndex_wo, rowIndex_wo + 1, colIndex_wo].Merge = true;
                    worksheet_wo.Cells[rowIndex_wo, colIndex_wo, rowIndex_wo + 1, colIndex_wo].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    worksheet_wo.Cells[rowIndex_wo, colIndex_wo, rowIndex_wo + 1, colIndex_wo].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                    colIndex_wo++;

                    worksheet_wo.Cells[rowIndex_wo, colIndex_wo].Value = item.FR_Date_S.ToString("yyyy-MM-dd HH:mm");
                    worksheet_wo.Cells[rowIndex_wo + 1, colIndex_wo++].Value = item.FR_Date_E.ToString("yyyy-MM-dd HH:mm");

                    worksheet_wo.Cells[rowIndex_wo, colIndex_wo].Value = item.FR_total_hour;
                    worksheet_wo.Cells[rowIndex_wo, colIndex_wo, rowIndex_wo + 1, colIndex_wo].Merge = true;
                    worksheet_wo.Cells[rowIndex_wo, colIndex_wo, rowIndex_wo + 1, colIndex_wo].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    worksheet_wo.Cells[rowIndex_wo, colIndex_wo, rowIndex_wo + 1, colIndex_wo].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                    colIndex_wo++;
                    worksheet_wo.Cells[rowIndex_wo, colIndex_wo].Value = item.Sign_name;
                    worksheet_wo.Cells[rowIndex_wo, colIndex_wo, rowIndex_wo + 1, colIndex_wo].Merge = true;
                    worksheet_wo.Cells[rowIndex_wo, colIndex_wo, rowIndex_wo + 1, colIndex_wo].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    worksheet_wo.Cells[rowIndex_wo, colIndex_wo, rowIndex_wo + 1, colIndex_wo].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                    colIndex_wo++;
                    rowIndex_wo++;
                    rowIndex_wo++;
                }

                // 添加框線
                var range_wo = worksheet_wo.Cells[1, 1, rowIndex_wo - 1, colIndex_wo - 1];
                range_wo.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                range_wo.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                range_wo.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                range_wo.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                // 自動調整列寬
                worksheet_wo.Cells.AutoFitColumns();
                #endregion

                return package.GetAsByteArray();
            }
        }

        public static byte[] ReceivableExcel(List<Receivable_Excel> items, Dictionary<string, string> headers)
        {
            if (items == null || items.Count == 0)
            {
                throw new ArgumentException("列表不能為 null 或空。", nameof(items));
            }

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Sheet0");

                // 添加表頭行
                int headerIndex = 1;
                foreach (var header in headers)
                {
                    var cell = worksheet.Cells[1, headerIndex++];
                    cell.Value = header.Value;
                    // 設置儲存格底色為淺藍色
                    cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    cell.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                }

                int colIndex = 1;
                int rowIndex = 1;
                int index = 1;
                foreach (var item in items)
                {
                    colIndex = 1;
                    rowIndex++; // 從第二行開始填充數據
                    worksheet.Cells[rowIndex, colIndex++].Value = index++;
                    worksheet.Cells[rowIndex, colIndex++].Value = item.CS_name;
                    worksheet.Cells[rowIndex, colIndex].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells[rowIndex, colIndex++].Value = item.amount_total;
                    worksheet.Cells[rowIndex, colIndex++].Value = item.month_total;
                    worksheet.Cells[rowIndex, colIndex++].Value = item.RC_count;
                    worksheet.Cells[rowIndex, colIndex++].Value = item.RC_date;
                    worksheet.Cells[rowIndex, colIndex].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells[rowIndex, colIndex++].Value = item.RC_amount;
                    worksheet.Cells[rowIndex, colIndex].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells[rowIndex, colIndex++].Value = item.interest;
                    worksheet.Cells[rowIndex, colIndex].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells[rowIndex, colIndex++].Value = item.Rmoney;
                    worksheet.Cells[rowIndex, colIndex].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells[rowIndex, colIndex++].Value = item.RemainingAmount;
                    worksheet.Cells[rowIndex, colIndex].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells[rowIndex, colIndex++].Value = item.PartiallySettled;
                    worksheet.Cells[rowIndex, colIndex++].Value = item.DelayDay;
                    worksheet.Cells[rowIndex, colIndex++].Value = item.Delaymoney;
                    if (item.check_pay_type == "N")
                    {
                        worksheet.Cells[rowIndex, colIndex++].Value = "未沖銷";
                        worksheet.Cells[rowIndex, colIndex-1].Style.Font.Color.SetColor(Color.Red);
                    }
                    if(item.check_pay_type == "Y")
                    {
                        worksheet.Cells[rowIndex, colIndex++].Value = "已沖銷";
                        worksheet.Cells[rowIndex, colIndex-1].Style.Font.Color.SetColor(Color.Blue);
                    }
                    if (item.check_pay_type == "S")
                    {
                        worksheet.Cells[rowIndex, colIndex++].Value = "已結清";
                        worksheet.Cells[rowIndex, colIndex-1].Style.Font.Color.SetColor(Color.Black);
                    }
                    worksheet.Cells[rowIndex, colIndex++].Value = item.check_pay_date;
                    worksheet.Cells[rowIndex, colIndex++].Value = item.check_pay_name;
                    worksheet.Cells[rowIndex, colIndex++].Value = item.RC_note;
                    if (item.bad_debt_type == "Y" && item.check_pay_type != "S")
                    {
                        worksheet.Cells[rowIndex, colIndex++].Value = "已轉呆";
                        worksheet.Cells[rowIndex, colIndex-1].Style.Font.Color.SetColor(Color.Black);
                    }
                    if (item.bad_debt_type == "N" && item.check_pay_type != "S")
                    {
                        worksheet.Cells[rowIndex, colIndex++].Value = "未轉呆";
                        worksheet.Cells[rowIndex, colIndex-1].Style.Font.Color.SetColor(Color.Blue);
                    }
                    if (item.check_pay_type == "S")
                    {
                        worksheet.Cells[rowIndex, colIndex++].Value = "-";
                    }
                    worksheet.Cells[rowIndex, colIndex++].Value = item.bad_debt_date;
                    worksheet.Cells[rowIndex, colIndex++].Value = item.bad_debt_name;
                    worksheet.Cells[rowIndex, colIndex++].Value = item.invoice_no;
                    worksheet.Cells[rowIndex, colIndex++].Value = item.invoice_date;
                }

                worksheet.Cells[rowIndex + 1, 8].Style.Numberformat.Format = "#,##0";
                worksheet.Cells[rowIndex + 1, 8].Value = items.Sum(x => x.interest);
                worksheet.Cells[rowIndex + 1, 9].Style.Numberformat.Format = "#,##0";
                worksheet.Cells[rowIndex + 1, 9].Value = items.Sum(x => x.Rmoney);
                worksheet.Cells[rowIndex + 1, 10].Style.Numberformat.Format = "#,##0";
                worksheet.Cells[rowIndex + 1, 10].Value = items.Sum(x => x.RemainingAmount);
                worksheet.Cells[rowIndex + 1, 11].Style.Numberformat.Format = "#,##0";
                worksheet.Cells[rowIndex + 1, 11].Value = items.Sum(x => x.PartiallySettled);
                worksheet.Cells[rowIndex + 1, 12].Style.Numberformat.Format = "#,##0";
                worksheet.Cells[rowIndex + 1, 12].Value = items.Sum(x => x.DelayDay);

                // 添加框線
                var range = worksheet.Cells[1, 1, rowIndex, headers.Count];
                range.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                range.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                range.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                // 自動調整列寬
                worksheet.Cells.AutoFitColumns();

                return package.GetAsByteArray();

            }
        }

        public static byte[] ReceivableNewExcel(List<Receivable_New_Excel> items, Dictionary<string, string> headers)
        {
            if (items == null || items.Count == 0)
            {
                throw new ArgumentException("列表不能為 null 或空。", nameof(items));
            }

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Sheet0");

                // 添加表頭行
                int headerIndex = 1;
                foreach (var header in headers)
                {
                    var cell = worksheet.Cells[1, headerIndex++];
                    cell.Value = header.Value;
                    // 設置儲存格底色為淺藍色
                    cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    cell.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                }

                int rowIndex = 1;
                int colIndex = 1;
                int index = 1;
                foreach (var item in items)
                {
                    colIndex = 1;
                    rowIndex++; // 從第二行開始填充數據
                    worksheet.Cells[rowIndex, colIndex++].Value = index++;
                    worksheet.Cells[rowIndex, colIndex++].Value = item.CS_name;
                    worksheet.Cells[rowIndex, colIndex].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells[rowIndex, colIndex++].Value = item.amount_total;
                    worksheet.Cells[rowIndex, colIndex++].Value = item.month_total;
                    worksheet.Cells[rowIndex, colIndex++].Value = item.RC_count;
                    worksheet.Cells[rowIndex, colIndex++].Value = item.RC_date;
                    worksheet.Cells[rowIndex, colIndex++].Value = item.DiffDay;
                    worksheet.Cells[rowIndex, colIndex].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells[rowIndex, colIndex++].Value = item.RC_amount;
                    worksheet.Cells[rowIndex, colIndex].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells[rowIndex, colIndex++].Value = item.interest;
                    worksheet.Cells[rowIndex, colIndex].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells[rowIndex, colIndex++].Value = item.Rmoney;
                    worksheet.Cells[rowIndex, colIndex].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells[rowIndex, colIndex++].Value = item.RemainingPrincipal;
                    worksheet.Cells[rowIndex, colIndex].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells[rowIndex, colIndex++].Value = item.RemainingPrincipal_1;
                }

                worksheet.Cells[rowIndex + 1, 8].Style.Numberformat.Format = "#,##0";
                worksheet.Cells[rowIndex + 1, 8].Value = items.Sum(x => x.RC_amount);
                worksheet.Cells[rowIndex + 1, 9].Style.Numberformat.Format = "#,##0";
                worksheet.Cells[rowIndex + 1, 9].Value = items.Sum(x => x.interest);
                worksheet.Cells[rowIndex + 1, 10].Style.Numberformat.Format = "#,##0";
                worksheet.Cells[rowIndex + 1, 10].Value = items.Sum(x => x.Rmoney);
                worksheet.Cells[rowIndex + 1, 11].Style.Numberformat.Format = "#,##0";
                worksheet.Cells[rowIndex + 1, 11].Value = items.Sum(x => x.RemainingPrincipal);

                // 添加框線
                var range = worksheet.Cells[1, 1, rowIndex, headers.Count];
                range.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                range.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                range.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                // 自動調整列寬
                worksheet.Cells.AutoFitColumns();

                return package.GetAsByteArray();
            }
        }
    }
}
