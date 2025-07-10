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
using Newtonsoft.Json.Linq;
using System.Collections;
using KF_WebAPI.Controllers;
using System.Net.Http;
using Microsoft.AspNetCore.Mvc;
using System.Text;
using System.Linq;
using OfficeOpenXml.Style;
using System.IO;
using static OfficeOpenXml.ExcelErrorValue;
using KF_WebAPI.DataLogic;
using System.Diagnostics;

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
        public static string GetCheckNum()
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
                            worksheet.Cells[i + 2, j + 2].Style.Numberformat.Format = "#,##0";
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

        public static SpecialClass CheckSpecial(string[] strings, string U_num)
        {
            var sc = new SpecialClass
            {
                special_check = "N",
                BC_Strings = "zz",
                U_num = U_num
            };

            if (strings == null || strings.Length == 0)
                return sc;

            ADOData adoData = new ADOData();

            var parameters = new List<SqlParameter>();
            var spIdParams = new List<string>();
            for (int i = 0; i < strings.Length; i++)
            {
                string paramName = "@SP_id" + i;
                spIdParams.Add(paramName);
                parameters.Add(new SqlParameter(paramName, strings[i]));
            }
            parameters.Add(new SqlParameter("@U_num", U_num));

            string query = $@"SELECT SP_id, SP_type FROM Special_set WHERE U_num = @U_num 
                              AND SP_id IN ({string.Join(",", spIdParams)}) AND del_tag = '0'";

            DataTable dtResult = adoData.ExecuteQuery(query, parameters);

            var specialMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (DataRow row in dtResult.Rows)
            {
                var sp_id = row["SP_id"].ToString();
                var sp_type = row["SP_type"].ToString();
                specialMap[sp_id] = sp_type;
            }

            var spToBCMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                {"7020", ",BC0100"},
                {"7021", ",BC0200"},
                {"7022", ",BC0600"},
                {"7023", ",BC0300"},
                {"7024", ",BC0500"},
                {"7025", ",BC0400"}
            };

            foreach (var s in strings)
            {
                if (specialMap.TryGetValue(s, out string spType) && spType == "1")
                {
                    sc.special_check = "Y";
                    if (spToBCMapping.TryGetValue(s, out string bcValue))
                    {
                        sc.BC_Strings += bcValue;
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

        public static byte[] FeatDailyToExcelAgain<T>(byte[] existingFileBytes, List<T> items, Dictionary<string, string> headers, string name, string datestring, int people)
        {
            using (var package = new ExcelPackage(new MemoryStream(existingFileBytes)))
            {
                var worksheet = package.Workbook.Worksheets[0];

                int existingRowCount = worksheet.Dimension.Rows; // 抓現有的行數
                int startRowIndex = existingRowCount + 2; // 留出一行空白

                // 添加合併標題
                worksheet.Cells[startRowIndex, 1].Value = "國峯租賃股份有限公司(" + name + ") 1+" + (people - 1) + "人     " + datestring;
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
            var parts = rocDate.Split('-');
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

            string gregorianDate = $"{gregorianYear}-{formattedMonth}-{formattedDay}";

            return gregorianDate;
        }

        public static byte[] ApprovalLoanSalesExcelFooter(byte[] existingFileBytes)
        {
            using (var package = new ExcelPackage(new MemoryStream(existingFileBytes)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int existingRowCount = worksheet.Dimension.Rows;//抓現有的行數
                int startRowIndex = existingRowCount + 1;


                worksheet.Cells[startRowIndex, 1].Value = "總計:" + (existingRowCount - 1);
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

        public static byte[] AttendanceExcelAllMonth(List<Attendance_report_excel> modelList, string yyyy, string mm, List<Flow_rest_HR_excel> flowRestList)
        {
            using (var package = new ExcelPackage())
            {
                #region 各公司打卡資料
                var bcOrder = new List<string> { "BC0800", "BC0801", "BC0802", "BC0803", "BC0900", "BC0100", "BC0200", "BC0600", "BC0300", "BC0500", "BC0400", "BC0700" };
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
                        case "BC0801":
                            worksheet = package.Workbook.Worksheets.Add("資訊部");
                            break;
                        case "BC0802":
                            worksheet = package.Workbook.Worksheets.Add("審查部");
                            break;
                        case "BC0803":
                            worksheet = package.Workbook.Worksheets.Add("財會部");
                            break;
                        case "BC0900":
                            worksheet = package.Workbook.Worksheets.Add("行銷部");
                            break;
                        case "BC0100":
                            worksheet = package.Workbook.Worksheets.Add("台北");
                            break;
                        case "BC0200":
                            worksheet = package.Workbook.Worksheets.Add("新北");
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
                                if (items[j].type == "Hk_04")
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

                                    if (RestList.Count > 0)
                                    {
                                        foreach (var rest in RestList.Where(x => x.Sign_name.Equals("同意")).ToList())
                                        {
                                            string displayValue = (rest.FR_total_hour % 1 == 0) ? ((int)rest.FR_total_hour).ToString() : rest.FR_total_hour.ToString();
                                            if (rest.FR_kind == "FRK017") //忘打卡FRK017
                                            {
                                                if (rest.FR_Date_S.Hour == 9)
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
                                            else if (rest.FR_kind == "FRK016") //外出 FRK016
                                            {
                                                items[j].typename += rest.FR_note;
                                                TimeSpan LateTime = new TimeSpan(9, 0, 0);
                                                TimeSpan EarlyTime = new TimeSpan(18, 0, 0);
                                                if (rest.FR_Date_S.TimeOfDay == LateTime)
                                                {
                                                    items[j].Late = 0;
                                                }
                                                if (rest.FR_Date_E.TimeOfDay == EarlyTime)
                                                {
                                                    items[j].early = 0;
                                                }
                                            }
                                            else if (rest.FR_total_hour >= 8)
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
                                    #region 因為消毒所以06/13 5F人員提早下班不算曠職
                                    if (bcGroup.Key == "BC0800" || bcGroup.Key == "BC0801" || bcGroup.Key == "BC0802" || bcGroup.Key == "BC0803"
                                        || bcGroup.Key == "BC0100")
                                    {
                                        if (items[j].attendance_week == "06/13 (五)")
                                        {
                                            items[j].early = 0;
                                        }
                                    }
                                    #endregion

                                    if (items[j].early > 0)
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
                        var lateTotal = 0;
                        if (userIDGroup.Where(x => x.userID == "K0330").Count() > 0)
                        {
                            lateTotal = Convert.ToInt32(userIDGroup.Sum(x => x.Late));
                        }
                        else
                        {
                            lateTotal = Convert.ToInt32(userIDGroup.Sum(x => x.Late) - 15);
                        }
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
                worksheet_ly.Cells[1, 1].Value = Convert.ToInt32(yyyy) - 1911 + "年" + mm + "月  國峯遲到、早退";
                worksheet_ly.Cells[1, 1, 1, 4].Merge = true;
                worksheet_ly.Cells[1, 1, 1, 4].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                worksheet_ly.Cells[1, 1, 1, 4].Style.Font.Bold = true;

                worksheet_ly.Cells[2, 1].Value = "名稱";
                worksheet_ly.Cells[2, 4].Value = "遲到、早退";

                int rowindex_ly = 2;
                int colindex_ly = 1;
                var userResult = modelList.Where(x => (x.absenteeism ?? "").Equals("Y") == false).GroupBy(x => x.userID).Select(g => new { UserID = g.Key, Totalval = g.Sum(x => x.Late) }).Where(y => y.Totalval > 0).OrderBy(x => x.UserID);
                foreach (var user in userResult)
                {
                    if (user.Totalval > 15 && user.UserID != "K0330")
                    {
                        var name = modelList.Where(x => x.userID.Equals(user.UserID)).FirstOrDefault().user_name;
                        rowindex_ly++;
                        worksheet_ly.Cells[rowindex_ly, colindex_ly++].Value = user.UserID + ":" + name;
                        worksheet_ly.Cells[rowindex_ly, colindex_ly++].Value = "合計";
                        worksheet_ly.Cells[rowindex_ly, colindex_ly++].Value = "扣";
                        decimal lateHour = Convert.ToInt32(user.Totalval - 15) / 60m;
                        worksheet_ly.Cells[rowindex_ly, colindex_ly++].Value = Math.Ceiling(lateHour / 0.5m) * 0.5m + "小時";
                        colindex_ly = 1;
                    }
                    else if (user.UserID == "K0330")
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

                #region 曠職
                // 添加合併標題 曠職
                worksheet_ly.Cells[1, 6].Value = Convert.ToInt32(yyyy) - 1911 + "年" + mm + "月  國峯曠職";
                worksheet_ly.Cells[1, 6, 1, 10].Merge = true;
                worksheet_ly.Cells[1, 6, 1, 10].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                worksheet_ly.Cells[1, 6, 1, 10].Style.Font.Bold = true;

                worksheet_ly.Cells[2, 6].Value = "名稱";
                worksheet_ly.Cells[2, 7].Value = "日期";
                worksheet_ly.Cells[2, 10].Value = "曠職";

                rowindex_ly = 2;
                colindex_ly = 6;

                foreach (var item in modelList.Where(x => (x.absenteeism ?? "").Equals("Y") == true).OrderBy(x => x.userID))
                {
                    rowindex_ly++;
                    worksheet_ly.Cells[rowindex_ly, colindex_ly++].Value = $"{item.userID}: {item.user_name}";
                    worksheet_ly.Cells[rowindex_ly, colindex_ly++].Value = item.attendance_week;
                    colindex_ly++;
                    colindex_ly++;
                    decimal earlyHour = Convert.ToInt32(item.early) / 60m;
                    worksheet_ly.Cells[rowindex_ly, colindex_ly++].Value = Math.Ceiling(earlyHour / 0.5m) * 0.5m + "H";
                    colindex_ly = 6;
                }

                // 添加框線
                var range_ly2 = worksheet_ly.Cells[1, 6, rowindex_ly, 10];
                range_ly2.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                range_ly2.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                range_ly2.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                range_ly2.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                #endregion

                //颱風假&&無請假紀錄&&在職
                if (modelList?.Any(x => x.type != null && x.type.Equals("Hk_04")) == true)
                {
                    // 添加合併標題 曠職
                    worksheet_ly.Cells[1, 12].Value = Convert.ToInt32(yyyy) - 1911 + "年" + mm + "月  國峯颱風假";
                    worksheet_ly.Cells[1, 12, 1, 13].Merge = true;
                    worksheet_ly.Cells[1, 12, 1, 13].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    worksheet_ly.Cells[1, 12, 1, 13].Style.Font.Bold = true;

                    worksheet_ly.Cells[2, 12].Value = "名稱";
                    worksheet_ly.Cells[2, 13].Value = "日期";

                    rowindex_ly = 2;
                    colindex_ly = 12;

                    foreach (var bcGroup in modelList.Where(x => x.type != null && x.type.Equals("Hk_04")).GroupBy(x => x.U_BC).OrderBy(g => bcOrder.IndexOf(g.Key)).ToList())
                    {
                        rowindex_ly++;
                        switch (bcGroup.Key)
                        {
                            case "BC0800":
                                worksheet_ly.Cells[rowindex_ly, 12].Value = "總公司";
                                worksheet_ly.Cells[rowindex_ly, 12, rowindex_ly, 13].Merge = true;
                                worksheet_ly.Cells[rowindex_ly, 12, rowindex_ly, 13].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                worksheet_ly.Cells[rowindex_ly, 12, rowindex_ly, 13].Style.Font.Bold = true;
                                break;
                            case "BC0801":
                                worksheet_ly.Cells[rowindex_ly, 12].Value = "資訊部";
                                worksheet_ly.Cells[rowindex_ly, 12, rowindex_ly, 13].Merge = true;
                                worksheet_ly.Cells[rowindex_ly, 12, rowindex_ly, 13].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                worksheet_ly.Cells[rowindex_ly, 12, rowindex_ly, 13].Style.Font.Bold = true;
                                break;
                            case "BC0802":
                                worksheet_ly.Cells[rowindex_ly, 12].Value = "審查部";
                                worksheet_ly.Cells[rowindex_ly, 12, rowindex_ly, 13].Merge = true;
                                worksheet_ly.Cells[rowindex_ly, 12, rowindex_ly, 13].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                worksheet_ly.Cells[rowindex_ly, 12, rowindex_ly, 13].Style.Font.Bold = true;
                                break;
                            case "BC0803":
                                worksheet_ly.Cells[rowindex_ly, 12].Value = "財會部";
                                worksheet_ly.Cells[rowindex_ly, 12, rowindex_ly, 13].Merge = true;
                                worksheet_ly.Cells[rowindex_ly, 12, rowindex_ly, 13].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                worksheet_ly.Cells[rowindex_ly, 12, rowindex_ly, 13].Style.Font.Bold = true;
                                break;
                            case "BC0900":
                                worksheet_ly.Cells[rowindex_ly, 12].Value = "行銷部";
                                worksheet_ly.Cells[rowindex_ly, 12, rowindex_ly, 13].Merge = true;
                                worksheet_ly.Cells[rowindex_ly, 12, rowindex_ly, 13].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                worksheet_ly.Cells[rowindex_ly, 12, rowindex_ly, 13].Style.Font.Bold = true;
                                break;
                            case "BC0100":
                                worksheet_ly.Cells[rowindex_ly, 12].Value = "台北";
                                worksheet_ly.Cells[rowindex_ly, 12, rowindex_ly, 13].Merge = true;
                                worksheet_ly.Cells[rowindex_ly, 12, rowindex_ly, 13].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                worksheet_ly.Cells[rowindex_ly, 12, rowindex_ly, 13].Style.Font.Bold = true;
                                break;
                            case "BC0200":
                                worksheet_ly.Cells[rowindex_ly, 12].Value = "新北";
                                worksheet_ly.Cells[rowindex_ly, 12, rowindex_ly, 13].Merge = true;
                                worksheet_ly.Cells[rowindex_ly, 12, rowindex_ly, 13].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                worksheet_ly.Cells[rowindex_ly, 12, rowindex_ly, 13].Style.Font.Bold = true;
                                break;
                            case "BC0600":
                                worksheet_ly.Cells[rowindex_ly, 12].Value = "桃園";
                                worksheet_ly.Cells[rowindex_ly, 12, rowindex_ly, 13].Merge = true;
                                worksheet_ly.Cells[rowindex_ly, 12, rowindex_ly, 13].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                worksheet_ly.Cells[rowindex_ly, 12, rowindex_ly, 13].Style.Font.Bold = true;
                                break;
                            case "BC0300":
                                worksheet_ly.Cells[rowindex_ly, 12].Value = "台中";
                                worksheet_ly.Cells[rowindex_ly, 12, rowindex_ly, 13].Merge = true;
                                worksheet_ly.Cells[rowindex_ly, 12, rowindex_ly, 13].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                worksheet_ly.Cells[rowindex_ly, 12, rowindex_ly, 13].Style.Font.Bold = true;
                                break;
                            case "BC0500":
                                worksheet_ly.Cells[rowindex_ly, 12].Value = "台南";
                                worksheet_ly.Cells[rowindex_ly, 12, rowindex_ly, 13].Merge = true;
                                worksheet_ly.Cells[rowindex_ly, 12, rowindex_ly, 13].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                worksheet_ly.Cells[rowindex_ly, 12, rowindex_ly, 13].Style.Font.Bold = true;
                                break;
                            case "BC0400":
                                worksheet_ly.Cells[rowindex_ly, 12].Value = "高雄";
                                worksheet_ly.Cells[rowindex_ly, 12, rowindex_ly, 13].Merge = true;
                                worksheet_ly.Cells[rowindex_ly, 12, rowindex_ly, 13].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                worksheet_ly.Cells[rowindex_ly, 12, rowindex_ly, 13].Style.Font.Bold = true;
                                break;
                            case "BC0700":
                                worksheet_ly.Cells[rowindex_ly, 12].Value = "湧立";
                                worksheet_ly.Cells[rowindex_ly, 12, rowindex_ly, 13].Merge = true;
                                worksheet_ly.Cells[rowindex_ly, 12, rowindex_ly, 13].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                worksheet_ly.Cells[rowindex_ly, 12, rowindex_ly, 13].Style.Font.Bold = true;
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
                                            colindex_ly = 12;
                                        }
                                    }
                                }
                            }
                        }
                    }

                    // 添加框線
                    var range_ly3 = worksheet_ly.Cells[1, 12, rowindex_ly, 13];
                    range_ly3.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    range_ly3.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    range_ly3.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    range_ly3.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
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
                foreach (var item in flowRestList.Where(x => x.FR_kind != "FRK016" && x.FR_kind != "FRK017" && x.FR_kind != "FRK021").OrderBy(x => x.FR_Date_S).ToList())
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
                string[] headers_wo = new string[] { "序號", "名稱", "加班起迄日", "加班時數", "選擇項目", "簽核結果" };

                int rowIndex_wo = 1;
                int colIndex_wo = 1;

                foreach (var header in headers_wo)
                {
                    worksheet_wo.Cells[rowIndex_wo, colIndex_wo++].Value = header;
                }

                rowIndex_wo++;
                int wo_index = 1;


                foreach (var item in flowRestList.Where(x => x.FR_kind.Equals("FRK021")).ToList())
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

                    string compensationType = item.FR_ot_compen;
                    string compensationDescription = compensationType switch
                    {
                        "0" => "補休",
                        "1" => "加班費",
                        "2" => "其他"
                    };
                    worksheet_wo.Cells[rowIndex_wo, colIndex_wo].Value = compensationDescription;
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
                        worksheet.Cells[rowIndex, colIndex - 1].Style.Font.Color.SetColor(Color.Red);
                    }
                    if (item.check_pay_type == "Y")
                    {
                        worksheet.Cells[rowIndex, colIndex++].Value = "已沖銷";
                        worksheet.Cells[rowIndex, colIndex - 1].Style.Font.Color.SetColor(Color.Blue);
                    }
                    if (item.check_pay_type == "S")
                    {
                        worksheet.Cells[rowIndex, colIndex++].Value = "已結清";
                        worksheet.Cells[rowIndex, colIndex - 1].Style.Font.Color.SetColor(Color.Black);
                    }
                    worksheet.Cells[rowIndex, colIndex++].Value = item.check_pay_date;
                    worksheet.Cells[rowIndex, colIndex++].Value = item.check_pay_name;
                    worksheet.Cells[rowIndex, colIndex++].Value = item.RC_note;
                    if (item.bad_debt_type == "Y" && item.check_pay_type != "S")
                    {
                        worksheet.Cells[rowIndex, colIndex++].Value = "已轉呆";
                        worksheet.Cells[rowIndex, colIndex - 1].Style.Font.Color.SetColor(Color.Black);
                    }
                    if (item.bad_debt_type == "N" && item.check_pay_type != "S")
                    {
                        worksheet.Cells[rowIndex, colIndex++].Value = "未轉呆";
                        worksheet.Cells[rowIndex, colIndex - 1].Style.Font.Color.SetColor(Color.Blue);
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

        public static byte[] ReceivableCollExcel(List<Receivable_Coll_Excel> items, Dictionary<string, string> headers)
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

        public static byte[] ReceivableLatePayExcel(List<Receivable_Late_Pay_Excel> items, Dictionary<string, string> headers)
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
                    worksheet.Cells[rowIndex, colIndex].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells[rowIndex, colIndex++].Value = item.RC_amount;
                    worksheet.Cells[rowIndex, colIndex].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells[rowIndex, colIndex++].Value = item.interest;
                    worksheet.Cells[rowIndex, colIndex].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells[rowIndex, colIndex++].Value = item.Rmoney;
                    worksheet.Cells[rowIndex, colIndex].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells[rowIndex, colIndex++].Value = item.RemainingPrincipal;
                    worksheet.Cells[rowIndex, colIndex++].Value = item.DelayDay;
                }

                worksheet.Cells[rowIndex + 1, 7].Style.Numberformat.Format = "#,##0";
                worksheet.Cells[rowIndex + 1, 7].Value = items.Sum(x => x.RC_amount);
                worksheet.Cells[rowIndex + 1, 8].Style.Numberformat.Format = "#,##0";
                worksheet.Cells[rowIndex + 1, 8].Value = items.Sum(x => x.interest);
                worksheet.Cells[rowIndex + 1, 9].Style.Numberformat.Format = "#,##0";
                worksheet.Cells[rowIndex + 1, 9].Value = items.Sum(x => x.Rmoney);
                worksheet.Cells[rowIndex + 1, 10].Style.Numberformat.Format = "#,##0";
                worksheet.Cells[rowIndex + 1, 10].Value = items.Sum(x => x.RemainingPrincipal);

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

        public static byte[] ReceivableOverRelExcel(List<Receivable_Over_Rel_Excel> items, Dictionary<string, string> headers)
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
                    worksheet.Cells[rowIndex, colIndex++].Value = item.BC_name;
                    worksheet.Cells[rowIndex, colIndex++].Value = item.u_name;
                    worksheet.Cells[rowIndex, colIndex++].Value = item.ToT_Count;
                    worksheet.Cells[rowIndex, colIndex].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells[rowIndex, colIndex++].Value = item.amount_total;
                    worksheet.Cells[rowIndex, colIndex++].Value = item.OV_Count;
                    worksheet.Cells[rowIndex, colIndex].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells[rowIndex, colIndex++].Value = item.OV_total;
                    double rate = double.Parse(item.OV_Rate.TrimEnd('%')) / 100;
                    worksheet.Cells[rowIndex, colIndex].Style.Numberformat.Format = "0.00%";
                    worksheet.Cells[rowIndex, colIndex++].Value = rate;
                }

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

        public static byte[] ReceivableRepayExcel(List<Receivable_Repay_Excel> items, Dictionary<string, string> headers)
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
                    worksheet.Cells[rowIndex, colIndex++].Value = item.yyyMM;
                    worksheet.Cells[rowIndex, colIndex++].Value = item.ToCount;
                    worksheet.Cells[rowIndex, colIndex++].Value = item.YCount;
                    worksheet.Cells[rowIndex, colIndex++].Value = item.NCount;
                    worksheet.Cells[rowIndex, colIndex++].Value = item.BCount;
                    worksheet.Cells[rowIndex, colIndex++].Value = item.SCount;
                    worksheet.Cells[rowIndex, colIndex].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells[rowIndex, colIndex++].Value = item.interest_T;
                    worksheet.Cells[rowIndex, colIndex].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells[rowIndex, colIndex++].Value = item.interest;
                    worksheet.Cells[rowIndex, colIndex].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells[rowIndex, colIndex++].Value = item.interest_U;
                    worksheet.Cells[rowIndex, colIndex].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells[rowIndex, colIndex++].Value = item.Rmoney_T;
                    worksheet.Cells[rowIndex, colIndex].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells[rowIndex, colIndex++].Value = item.Rmoney;
                    worksheet.Cells[rowIndex, colIndex].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells[rowIndex, colIndex++].Value = item.Rmoney_U;
                    worksheet.Cells[rowIndex, colIndex].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells[rowIndex, colIndex++].Value = item.RemainingPrincipal_BB;
                    worksheet.Cells[rowIndex, colIndex].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells[rowIndex, colIndex++].Value = item.S_AMT;
                    worksheet.Cells[rowIndex, colIndex].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells[rowIndex, colIndex++].Value = item.RemainingPrincipal;
                }

                //合併儲存格
                rowIndex++;
                worksheet.Cells[rowIndex, 1].Value = "總計：";
                worksheet.Cells[rowIndex, 1, rowIndex, 7].Merge = true;
                worksheet.Cells[rowIndex, 1, rowIndex, 7].Style.Font.Bold = true;

                worksheet.Cells[rowIndex, 8].Style.Numberformat.Format = "#,##0";
                worksheet.Cells[rowIndex, 8].Value = items.Sum(x => x.interest_T);
                worksheet.Cells[rowIndex, 9].Style.Numberformat.Format = "#,##0";
                worksheet.Cells[rowIndex, 9].Value = items.Sum(x => x.interest);
                worksheet.Cells[rowIndex, 10].Style.Numberformat.Format = "#,##0";
                worksheet.Cells[rowIndex, 10].Value = items.Sum(x => x.interest_U);
                worksheet.Cells[rowIndex, 11].Style.Numberformat.Format = "#,##0";
                worksheet.Cells[rowIndex, 11].Value = items.Sum(x => x.Rmoney_T);
                worksheet.Cells[rowIndex, 12].Style.Numberformat.Format = "#,##0";
                worksheet.Cells[rowIndex, 12].Value = items.Sum(x => x.Rmoney);
                worksheet.Cells[rowIndex, 13].Style.Numberformat.Format = "#,##0";
                worksheet.Cells[rowIndex, 13].Value = items.Sum(x => x.Rmoney_U);
                worksheet.Cells[rowIndex, 14].Style.Numberformat.Format = "#,##0";
                worksheet.Cells[rowIndex, 14].Value = items.Sum(x => x.RemainingPrincipal_BB);
                worksheet.Cells[rowIndex, 15].Style.Numberformat.Format = "#,##0";
                worksheet.Cells[rowIndex, 15].Value = items.Sum(x => x.S_AMT);

                // 添加框線
                var range = worksheet.Cells[1, 1, rowIndex - 1, headers.Count];
                range.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                range.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                range.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                // 自動調整列寬
                worksheet.Cells.AutoFitColumns();

                return package.GetAsByteArray();
            }
        }

        public static byte[] ReceivableExcessExcel(List<Receivable_Excess_Excel> items, Dictionary<string, string> headers)
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


                int m1 = items.Where(x => x.diffType == "M1").Count();
                int m2 = items.Where(x => x.diffType == "M2").Count();
                int m3 = items.Where(x => x.diffType == "M3").Count();

                worksheet.Cells[2, 1].Value = "M1";
                worksheet.Cells[2, 1, m1 + 1, 1].Merge = true; // 合併儲存格
                worksheet.Cells[2, 1, m1 + 1, 1].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                worksheet.Cells[m1 + 2, 1].Value = "M2";
                worksheet.Cells[m1 + 2, 1, m2 + m1 + 1, 1].Merge = true; // 合併儲存格
                worksheet.Cells[m1 + 2, 1, m2 + m1 + 1, 1].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                worksheet.Cells[m2 + m1 + 2, 1].Value = "M3";
                worksheet.Cells[m2 + m1 + 2, 1, m3 + m2 + m1 + 1, 1].Merge = true; // 合併儲存格
                worksheet.Cells[m2 + m1 + 2, 1, m3 + m2 + m1 + 1, 1].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                int rowIndex = 1;
                int colIndex = 2;
                foreach (var item in items)
                {
                    colIndex = 2;
                    rowIndex++;
                    worksheet.Cells[rowIndex, colIndex++].Value = item.AmtTypeDesc;
                    worksheet.Cells[rowIndex, colIndex++].Value = item.Count;
                    worksheet.Cells[rowIndex, colIndex].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells[rowIndex, colIndex++].Value = item.amount_total;
                    worksheet.Cells[rowIndex, colIndex].Style.Numberformat.Format = "0.00%";
                    worksheet.Cells[rowIndex, colIndex++].Value = item.Rate / 100;
                }

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

        public static byte[] ReceivableExcessDetailExcel(List<Receivable_Excess_Detail_Excel> items, Dictionary<string, string> headers)
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

                int m1 = items.Where(x => x.diffType == "M1").Count();
                int m2 = items.Where(x => x.diffType == "M2").Count();
                int m3 = items.Where(x => x.diffType == "M3").Count();

                worksheet.Cells[2, 1].Value = "M1";
                worksheet.Cells[2, 1, m1 + 1, 1].Merge = true; // 合併儲存格
                worksheet.Cells[2, 1, m1 + 1, 1].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                worksheet.Cells[m1 + 2, 1].Value = "M2";
                worksheet.Cells[m1 + 2, 1, m2 + m1 + 1, 1].Merge = true; // 合併儲存格
                worksheet.Cells[m1 + 2, 1, m2 + m1 + 1, 1].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                worksheet.Cells[m2 + m1 + 2, 1].Value = "M3";
                worksheet.Cells[m2 + m1 + 2, 1, m3 + m2 + m1 + 1, 1].Merge = true; // 合併儲存格
                worksheet.Cells[m2 + m1 + 2, 1, m3 + m2 + m1 + 1, 1].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                int rowIndex = 1;
                int colIndex = 2;
                foreach (var item in items)
                {
                    colIndex = 2;
                    rowIndex++;
                    worksheet.Cells[rowIndex, colIndex++].Value = item.AmtTypeDesc;
                    worksheet.Cells[rowIndex, colIndex++].Value = item.Cs_name;
                    worksheet.Cells[rowIndex, colIndex++].Value = item.DiffDay;
                    worksheet.Cells[rowIndex, colIndex].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells[rowIndex, colIndex++].Value = item.amount_total;
                    worksheet.Cells[rowIndex, colIndex++].Value = item.RCM_note;
                }

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

        public static byte[] PerformanceExcel(List<Performance_Excel> items, Dictionary<string, string> headers)
        {
            if (items == null || items.Count == 0)
            {
                throw new ArgumentException("The list cannot be null or empty.", nameof(items));
            }

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Sheet0");


                // 添加表頭
                worksheet.Cells[1, 1].Value = "排名";
                worksheet.Cells[1, 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                int headerIndex = 2;
                foreach (var header in headers)
                {
                    var cell = worksheet.Cells[1, headerIndex++];
                    cell.Value = header.Value;
                    // 設置儲存格底色為淺藍色
                    cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    cell.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                }


                // 添加表身
                int rowIndex = 2;
                foreach (var item in items)
                {
                    worksheet.Cells[rowIndex, 1].Value = rowIndex - 1;
                    worksheet.Cells[rowIndex, 2].Value = item.U_BC_name;
                    worksheet.Cells[rowIndex, 3].Value = item.title;
                    worksheet.Cells[rowIndex, 4].Value = item.U_name;
                    worksheet.Cells[rowIndex, 5].Value = item.U_arrive_date;
                    worksheet.Cells[rowIndex, 6].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells[rowIndex, 6].Value = item.Jan;
                    worksheet.Cells[rowIndex, 7].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells[rowIndex, 7].Value = item.Feb;
                    worksheet.Cells[rowIndex, 8].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells[rowIndex, 8].Value = item.Mar;
                    worksheet.Cells[rowIndex, 9].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells[rowIndex, 9].Value = item.Apr;
                    worksheet.Cells[rowIndex, 10].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells[rowIndex, 10].Value = item.Mar;
                    worksheet.Cells[rowIndex, 11].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells[rowIndex, 11].Value = item.Jun;
                    worksheet.Cells[rowIndex, 12].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells[rowIndex, 12].Value = item.Jul;
                    worksheet.Cells[rowIndex, 13].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells[rowIndex, 13].Value = item.Aug;
                    worksheet.Cells[rowIndex, 14].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells[rowIndex, 14].Value = item.Sep;
                    worksheet.Cells[rowIndex, 15].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells[rowIndex, 15].Value = item.Oct;
                    worksheet.Cells[rowIndex, 16].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells[rowIndex, 16].Value = item.Nov;
                    worksheet.Cells[rowIndex, 17].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells[rowIndex, 17].Value = item.Dec;
                    worksheet.Cells[rowIndex, 18].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells[rowIndex, 18].Value = item.Totle;
                    worksheet.Cells[rowIndex, 19].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells[rowIndex, 19].Value = item.MonAVG;
                    worksheet.Cells[rowIndex, 20].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells[rowIndex, 20].Value = item.YearAVG;
                    rowIndex++;
                }


                //合併儲存格
                worksheet.Cells[rowIndex, 1].Value = $"筆數：{items.Count}";
                worksheet.Cells[rowIndex, 1, rowIndex, 4].Merge = true;
                worksheet.Cells[rowIndex, 1, rowIndex, 4].Style.Font.Bold = true;
                worksheet.Cells[rowIndex, 5].Value = "合計:";
                worksheet.Cells[rowIndex, 6].Style.Numberformat.Format = "#,##0";
                worksheet.Cells[rowIndex, 6].Value = items.Sum(x => x.Jan);
                worksheet.Cells[rowIndex, 7].Style.Numberformat.Format = "#,##0";
                worksheet.Cells[rowIndex, 7].Value = items.Sum(x => x.Feb);
                worksheet.Cells[rowIndex, 8].Style.Numberformat.Format = "#,##0";
                worksheet.Cells[rowIndex, 8].Value = items.Sum(x => x.Mar);
                worksheet.Cells[rowIndex, 9].Style.Numberformat.Format = "#,##0";
                worksheet.Cells[rowIndex, 9].Value = items.Sum(x => x.Apr);
                worksheet.Cells[rowIndex, 10].Style.Numberformat.Format = "#,##0";
                worksheet.Cells[rowIndex, 10].Value = items.Sum(x => x.May);
                worksheet.Cells[rowIndex, 11].Style.Numberformat.Format = "#,##0";
                worksheet.Cells[rowIndex, 11].Value = items.Sum(x => x.Jun);
                worksheet.Cells[rowIndex, 12].Style.Numberformat.Format = "#,##0";
                worksheet.Cells[rowIndex, 12].Value = items.Sum(x => x.Jul);
                worksheet.Cells[rowIndex, 13].Style.Numberformat.Format = "#,##0";
                worksheet.Cells[rowIndex, 13].Value = items.Sum(x => x.Aug);
                worksheet.Cells[rowIndex, 14].Style.Numberformat.Format = "#,##0";
                worksheet.Cells[rowIndex, 14].Value = items.Sum(x => x.Sep);
                worksheet.Cells[rowIndex, 15].Style.Numberformat.Format = "#,##0";
                worksheet.Cells[rowIndex, 15].Value = items.Sum(x => x.Oct);
                worksheet.Cells[rowIndex, 16].Style.Numberformat.Format = "#,##0";
                worksheet.Cells[rowIndex, 16].Value = items.Sum(x => x.Nov);
                worksheet.Cells[rowIndex, 17].Style.Numberformat.Format = "#,##0";
                worksheet.Cells[rowIndex, 17].Value = items.Sum(x => x.Dec);
                worksheet.Cells[rowIndex, 18].Style.Numberformat.Format = "#,##0";
                worksheet.Cells[rowIndex, 18].Value = items.Sum(x => x.Totle);

                // 添加表身邊框
                using (var range = worksheet.Cells[1, 1, rowIndex, headers.Count + 1])
                {
                    range.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    range.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    range.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                }

                // 列寬自動調整
                worksheet.Cells.AutoFitColumns();

                return package.GetAsByteArray();
            }
        }

        public static ArrayList ConvertJsonToArrayList(string json)
        {
            ArrayList result = new ArrayList();
            DataTable dataTable = new DataTable();

            JArray jsonArray = JArray.Parse(json);

            result.Add(jsonArray.Count);

            if (jsonArray.Count > 0)
            {
                foreach (var property in jsonArray[0].ToObject<JObject>().Properties())
                {
                    dataTable.Columns.Add(property.Name);
                }

                // 填充 DataTable
                foreach (var item in jsonArray)
                {
                    DataRow row = dataTable.NewRow();
                    foreach (var property in item.ToObject<JObject>().Properties())
                    {
                        row[property.Name] = property.Value.ToString();
                    }
                    dataTable.Rows.Add(row);
                }

                dataTable.TableName = "Table";

                result.Add(dataTable);
            }

            return result;
        }

        public static bool AuditFlow(string U_num, string U_BC, string AF_ID, string FM_Source_ID, string IP)
        {
            bool result = false;

            ADOData _adoData = new ADOData();
            var T_SQL_AUDIT = @"select * from AuditFlow where AF_ID=@AF_ID";
            var parameters_audit = new List<SqlParameter>
            {
                new SqlParameter("@AF_ID", AF_ID)
            };
            var dtResultAudit = _adoData.ExecuteQuery(T_SQL_AUDIT, parameters_audit);
            int FM_step = Convert.ToInt32(dtResultAudit.Rows[0]["AF_Step"]); //流程步驟
            var FM_Step_Caption = dtResultAudit.Rows[0]["AF_Step_Caption"].ToString(); //流程步驟說明
            var AF_Step_Person = dtResultAudit.Rows[0]["AF_Step_Person"].ToString(); //流程人員設定

            var stepDescriptions = FM_Step_Caption.Split(',');
            var stepPersontions = AF_Step_Person.Split(',');

            //新增主流程
            var T_SQL_M = @"Insert into AuditFlow_M(AF_ID,FM_Source_ID,FM_BC,FM_Step,FM_Step_SignType,add_date,add_num,add_ip,edit_date,edit_num,edit_ip) 
                Values (@AF_ID,@FM_Source_ID,@FM_BC,'1','FSIGN001',GETDATE(),@add_num,@add_ip,GETDATE(),@edit_num,@edit_ip) ";
            var parameters_m = new List<SqlParameter>
            {
                new SqlParameter("@AF_ID", AF_ID),
                new SqlParameter("@FM_Source_ID", FM_Source_ID),
                new SqlParameter("@FM_BC",U_BC),
                new SqlParameter("@add_num", U_num),
                new SqlParameter("@add_ip", IP),
                new SqlParameter("@edit_num", U_num),
                new SqlParameter("@edit_ip", IP)
            };
            int dtResult_m = _adoData.ExecuteNonQuery(T_SQL_M, parameters_m);
            if (dtResult_m == 0)
            {
                return result;
            }
            else
            {
                var MsgNum = "";
                for (int i = 1; i <= FM_step; i++)
                {

                    var T_SQL_D = @"Insert into AuditFlow_D (AF_ID,FM_Source_ID,FD_Step,FD_Sign_Countersign,FD_Step_title,FD_Step_num,FD_Step_SignType,add_date,add_num,add_ip)
                        Values (@AF_ID,@FM_Source_ID,@FD_Step,'S',@FD_Step_title,@FD_Step_num,'FSIGN001',GETDATE(),@add_num,@add_ip)";
                    var parameters_d = new List<SqlParameter>
                    {
                        new SqlParameter("@AF_ID", AF_ID),
                        new SqlParameter("@FM_Source_ID", FM_Source_ID),
                        new SqlParameter("@FD_Step", i),
                        new SqlParameter("@FD_Step_title", stepDescriptions[i - 1]),
                        new SqlParameter("@add_num", U_num),
                        new SqlParameter("@add_ip", IP)
                    };

                    switch (stepPersontions[i - 1])
                    {
                        case "U0000": //雙副總
                            parameters_d.Add(new SqlParameter("@FD_Step_num", "K0002"));
                            _adoData.ExecuteNonQuery(T_SQL_D, parameters_d);
                            var parameters_dx = new List<SqlParameter>
                            {
                                new SqlParameter("@AF_ID", AF_ID),
                                new SqlParameter("@FM_Source_ID", FM_Source_ID),
                                new SqlParameter("@FD_Step", i),
                                new SqlParameter("@FD_Step_title", stepDescriptions[i - 1]),
                                new SqlParameter("@add_num", U_num),
                                new SqlParameter("@add_ip", IP),
                                new SqlParameter("@FD_Step_num", "K0001")
                            };
                            _adoData.ExecuteNonQuery(T_SQL_D, parameters_dx);
                            break;
                        case "U0001": //部門主管
                            var T_SQL_USER = @"select * from User_M where U_num = @U_num";
                            var parameters_user = new List<SqlParameter>
                            {
                                new SqlParameter("@U_num", U_num)
                            };
                            var dtResultUser = _adoData.ExecuteQuery(T_SQL_USER, parameters_user);
                            var leaderNum = dtResultUser.Rows[0]["U_leader_2_num"].ToString();
                            parameters_d.Add(new SqlParameter("@FD_Step_num", leaderNum));
                            _adoData.ExecuteNonQuery(T_SQL_D, parameters_d);
                            if (i == 1)
                                MsgNum = leaderNum;
                            break;
                        default:
                            if (stepPersontions[i - 1] == U_num) // 當申請人為審核者自動變更為代理人
                            {
                                var T_SQL_AGENT = @"select * from User_M where U_num = @U_num";
                                var parameters_agent = new List<SqlParameter>
                                {
                                    new SqlParameter("@U_num", U_num)
                                };
                                var dtResultAgent = _adoData.ExecuteQuery(T_SQL_AGENT, parameters_agent);
                                var agentNum = dtResultAgent.Rows[0]["U_agent_num"].ToString();
                                parameters_d.Add(new SqlParameter("@FD_Step_num", agentNum));
                            }
                            else
                            {
                                parameters_d.Add(new SqlParameter("@FD_Step_num", stepPersontions[i - 1]));
                            }
                            _adoData.ExecuteNonQuery(T_SQL_D, parameters_d);
                            if (i == 1)
                                MsgNum = stepPersontions[i - 1];
                            break;
                    }
                }
                result = true;
                if (result)
                {
                    var Fun = new FuncHandler();
                    //訊息通知
                    if (new[] { "PO", "PP", "PS", "PA" }.Contains(AF_ID))
                        Fun.MsgIns("MSGK0005", U_num, MsgNum, "請採購單或請款單簽核通知,請前往處理!!", IP);


                }
            }
            return result;
        }

        public void MsgIns(string Msg_Kind, string Add_User, string Msg_User, string Msg_Title, string IP)
        {
            ADOData _adoData = new ADOData();
            var parameters = new List<SqlParameter>();
            var T_SQL = @"Insert into Msg (Msg_cknum,Msg_source,Msg_kind,Msg_show_date,Msg_title,Msg_note,Msg_to_num,Msg_read_type,Msg_reply_note,Msg_reply_type
                ,del_tag,show_tag,add_date,add_num,add_ip) values (@Msg_cknum,'sys',@Msg_kind,GETDATE(),@Msg_title,'',@Msg_to_num,'N',''
                ,'N','0','0',GETDATE(),@add_num,@add_ip)";
            parameters.Add(new SqlParameter("@Msg_cknum", GetCheckNum()));
            parameters.Add(new SqlParameter("@Msg_kind", Msg_Kind));
            parameters.Add(new SqlParameter("@Msg_title", Msg_Title));
            parameters.Add(new SqlParameter("@Msg_to_num", Msg_User));
            parameters.Add(new SqlParameter("@add_num", Add_User));
            parameters.Add(new SqlParameter("@add_ip", IP));

            try
            {
                _adoData.ExecuteNonQuery(T_SQL, parameters);
            }
            catch (Exception)
            {
                throw;
            }

        }

        public void ExtAPILogIns(string API_CODE, string API_NAME, string API_KEY, string ACCESS_TOKEN, string PARAM_JSON, string RESULT_CODE, string RESULT_MSG)
        {
            ExtAPILogIns(API_CODE, API_NAME, API_KEY, ACCESS_TOKEN, PARAM_JSON, RESULT_CODE, RESULT_MSG, "sys");
        }

        public void ExtAPILogIns(string API_CODE, string API_NAME, string API_KEY, string ACCESS_TOKEN, string PARAM_JSON, string RESULT_CODE, string RESULT_MSG, string Add_User)
        {
            ADOData _adoData = new ADOData();
            var parameters = new List<SqlParameter>();
            var T_SQL = @"Insert into External_API_Log (API_CODE,API_NAME,API_KEY,ACCESS_TOKEN,PARAM_JSON,RESULT_CODE,RESULT_MSG,
                Add_date,Add_User) values (@API_CODE,@API_NAME,@API_KEY,@ACCESS_TOKEN,@PARAM_JSON,@RESULT_CODE,
                @RESULT_MSG,GETDATE(),@Add_User)";
            parameters.Add(new SqlParameter("@API_CODE", API_CODE));
            parameters.Add(new SqlParameter("@API_NAME", API_NAME));
            parameters.Add(new SqlParameter("@API_KEY", API_KEY));
            parameters.Add(new SqlParameter("@ACCESS_TOKEN", ACCESS_TOKEN));
            parameters.Add(new SqlParameter("@PARAM_JSON", PARAM_JSON));
            parameters.Add(new SqlParameter("@RESULT_CODE", RESULT_CODE));
            parameters.Add(new SqlParameter("@RESULT_MSG", RESULT_MSG));
            parameters.Add(new SqlParameter("@Add_User", Add_User));
            try
            {
                _adoData.ExecuteNonQuery(T_SQL, parameters);
            }
            catch (Exception)
            {
                throw;
            }
        }

        #region 難字相關處理
        /// <summary>
        /// 取得HKSCS對應的Unicode字元
        /// </summary>
        /// <param name="hkscs"></param>
        /// <returns></returns>
        public ResultClass<string> GetUniCode(string hkscs)
        {
            ResultClass<string> result = new ResultClass<string>();
            try
            {
                ADOData _adoData = new ADOData(); // 測試:"Test" / 正式:""
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"select [ISO10646] from [HkscsMapping] where [hkscs2008] = @hkscs";
                parameters.Add(new SqlParameter("@hkscs", hkscs));
                #endregion
                var dtResult = _adoData.ExecuteQuery(T_SQL, parameters);

                if (dtResult.Rows.Count > 0)
                {
                    result.ResultCode = "000";
                    result.objResult = dtResult.AsEnumerable().ToList().Select(row => row.Field<string>("ISO10646")).FirstOrDefault();
                }
                else
                {
                    result.ResultCode = "400";
                    result.ResultMsg = "查無資料";
                }
            }
            catch (Exception ex)
            {
                result.ResultCode = "500";
                result.ResultMsg = $" response: {ex.Message}";
            }
            return result;
        }
        public string DeCodeBig5Words(string strBig5)
        {
            string decodeWords = string.Empty;
            var words = CheckRareChineseCharacters(strBig5);
            foreach (var word in words)
            {
                if (word.IsRare)
                {
                    string HtmlEncode = ConvertDecimalToUrlEncoded(word.Character);
                    ResultClass<string> resUniCode = GetUniCode(HtmlEncode.Replace("%", ""));
                    if (resUniCode.ResultCode != "000")
                    {
                        continue;
                    }
                    string result = ConvertHexToUnicodeChar(resUniCode.objResult);
                    decodeWords += result; // 將解碼後的字元累加
                }
                else
                {
                    decodeWords += word.Character; // 將解碼後的字元累加
                }
            }
            return decodeWords;
        }
        public static string ConvertDecimalToUrlEncoded(string str)
        {
            StringBuilder sb = new StringBuilder();

            for (int i = 0; i < str.Length; i += Char.IsSurrogatePair(str, i) ? 2 : 1)
            {
                int codePoint = Char.ConvertToUtf32(str, i);
                var decimalInput = Char.ConvertFromUtf32(codePoint); //Converts the specified Unicode code point into a UTF-16 encoded string
                char c = (char)codePoint;                          // U+E1D0
                Encoding big5 = Encoding.GetEncoding(950);         // Big5 編碼
                byte[] bytes = big5.GetBytes(new[] { c });         // 取得原始 Big5 bytes

                // 轉成 %XX%YY
                foreach (byte b in bytes)
                {
                    sb.Append('%');
                    sb.Append(b.ToString("X2"));
                }
            }
            return sb.ToString();
        }
        public static string ConvertHexToUnicodeChar(string hex)
        {
            // 將十六進位字串轉成整數
            int codePoint = int.Parse(hex, System.Globalization.NumberStyles.HexNumber);

            // 將 codePoint 轉為 Unicode 字元
            string result = char.ConvertFromUtf32(codePoint);

            return result;
        }
        public static List<(string Character, bool IsRare)> CheckRareChineseCharacters(string text)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var result = new List<(string Character, bool IsRare)>();
            for (int i = 0; i < text.Length; i++)
            {
                int codePoint = char.ConvertToUtf32(text, i);

                // 如果是 surrogate pair，要跳過下個 char（它已包含在 codePoint 中）
                if (char.IsHighSurrogate(text[i]) && i + 1 < text.Length && char.IsLowSurrogate(text[i + 1]))
                    i++; // skip next
                bool isRare = IsRareChineseCodePoint(codePoint) || IsDifficultChineseChar(text[i].ToString());
                result.Add((text[i].ToString(), isRare));
            }
            return result;
        }
        public static bool IsRareChineseCodePoint(int codePoint)
        {
            // 常見中文字區：U+4E00 ~ U+9FFF
            // 擴展A區：U+3400 ~ U+4DBF
            // 超出這範圍的中文通常屬於難字

            return !((codePoint >= 0x3400 && codePoint <= 0x4DBF) ||
                      (codePoint >= 0x4E00 && codePoint <= 0x9FFF));
        }
        public static bool IsDifficultChineseChar(string s)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            var big5 = Encoding.GetEncoding(950, new EncoderExceptionFallback(), new DecoderExceptionFallback());

            foreach (var rune in s.EnumerateRunes()) // 支援 UTF-32 字元
            {
                int codePoint = rune.Value;

                // 若超出 CJK 常見範圍
                if (!((codePoint >= 0x3400 && codePoint <= 0x4DBF) ||
                       (codePoint >= 0x4E00 && codePoint <= 0x9FFF)))
                {
                    return true; // 難字
                }

                // 嘗試 Big5 編碼（失敗則為難字）
                try
                {
                    big5.GetBytes(rune.ToString());
                }
                catch
                {
                    return true; // 難字
                }
            }

            return false; // 全部是常用可編碼字
        }
        /// <summary>
        /// 將NCR轉換為字串
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public string fromNCR(string value)
        {
            return System.Text.RegularExpressions.Regex.Replace(
                value,
                @"&#(\d+)",
                m => char.ConvertFromUtf32(int.Parse(m.Groups[1].Value))
            );
        }
        public string toNCR(string rawString)
        {
            StringBuilder sb = new StringBuilder();
            Encoding big5 = CodePagesEncodingProvider.Instance.GetEncoding(name: "big5");
            //Encoding big5 = Encoding.GetEncoding("big5");
            foreach (char c in rawString)
            {
                //強迫轉碼成Big5，看會不會變成問號
                string cInBig5 = big5.GetString(big5.GetBytes(new char[] { c }));
                //原來不是問號，轉碼後變問號，判定為難字
                if (c != '?' && cInBig5 == "?")
                    sb.AppendFormat("&#{0};", Convert.ToInt32(c));
                else
                    sb.Append(c);
            }
            return sb.ToString();
        }
        #endregion
        public List<Per_Achieve> GetTargetAchieveList(string YYYY,string? U_BC)
        {
            ADOData _adoData = new ADOData();

            try
            {
                List<Per_Achieve> achievesList = new List<Per_Achieve>();
                for (int i = 1; i <= 12; i++)
                {
                    string month = i.ToString("D2");
                    string dateYM = $"{YYYY}-{month}";
                    var T_SQL = $"EXEC sp_GetTargetPerfByMonth '{dateYM}'";
                    var dtResult = _adoData.ExecuteSQuery(T_SQL);

                    var achieves = dtResult.AsEnumerable().Select(row => new Per_Achieve
                    {
                        month = row.Field<string>("month"),
                        U_BC_NEW = row.Field<string>("U_BC_NEW"),
                        total_target = row.Field<double?>("total_target") ?? 0,
                        total_perf = row.Field<double?>("total_perf") ?? 0,
                        total_perf_after_discount = row.Field<double?>("total_perf_after_discount") ?? 0,
                        Subord = row.Field<int>("Subord"),
                        Leader = row.Field<int>("Leader")
                    }).ToList();

                    achievesList.AddRange(achieves);
                }

                if (!string.IsNullOrEmpty(U_BC))
                {
                    if (U_BC == "BC0100")
                    {
                        string[] strBc = new string[] { "BC0100-1", "BC0100-2" };
                        achievesList = achievesList.Where(q => strBc.Contains(q.U_BC_NEW)).GroupBy(q => q.month)
                            .Select(g =>
                            {
                                double totalPerf = g.Sum(x => x.total_perf);
                                double totalTarget = g.Sum(x => x.total_target);
                                double totalPerfAfterDiscount = g.Sum(x => x.total_perf_after_discount);

                                string achieveRate = totalPerf != 0
                                    ? (totalPerf / totalTarget * 100).ToString("F2") + "%"
                                    : "0.00%";

                                string achieveRateAfterDiscount = totalPerfAfterDiscount != 0
                                    ? (totalPerfAfterDiscount / totalTarget * 100).ToString("F2") + "%"
                                    : "0.00%";

                                return new Per_Achieve
                                {
                                    month = g.Key,
                                    U_BC_NEW = "BC0100",
                                    total_target = totalTarget,
                                    total_perf = totalPerf,
                                    total_perf_after_discount = totalPerfAfterDiscount,
                                    achieve_rate = achieveRate,
                                    achieve_rate_after_discount = achieveRateAfterDiscount,
                                    Subord = g.Sum(x => x.Subord),
                                    Leader = g.Sum(x => x.Leader)
                                };
                            }).ToList();
                    }
                    else
                    {
                        achievesList = achievesList.Where(q => q.U_BC_NEW.Equals(U_BC)).OrderBy(q => q.month).Select(g =>
                        {
                            string achieveRate = g.total_perf != 0
                                    ? (g.total_perf / g.total_target * 100).ToString("F2") + "%"
                                    : "0.00%";

                            string achieveRateAfterDiscount = g.total_perf_after_discount != 0
                                ? (g.total_perf_after_discount / g.total_target * 100).ToString("F2") + "%"
                                : "0.00%";

                            return new Per_Achieve
                            {
                                month = g.month,
                                U_BC_NEW = g.U_BC_NEW,
                                total_target = g.total_target,
                                total_perf = g.total_perf,
                                total_perf_after_discount = g.total_perf_after_discount,
                                achieve_rate = achieveRate,
                                achieve_rate_after_discount = achieveRateAfterDiscount,
                                Subord = g.Subord,
                                Leader = g.Leader
                            };
                        }).ToList();
                    }
                }

                return achievesList;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static byte[] TargetAchieveToExcel(List<Per_Achieve> achieveList, Dictionary<string, string> headers)
        {
            var fieldsWithWan = new HashSet<string> { "total_target", "total_perf", "total_perf_after_discount" };
            var percentFields = new HashSet<string> { "achieve_rate", "achieve_rate_after_discount" };

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("責任額達成率");

                #region 台北
                string[] taipeiBC = { "BC0100-1", "BC0100-2" };
                var taipeiList = achieveList
                    .Where(q => taipeiBC.Contains(q.U_BC_NEW))
                    .GroupBy(q => q.month)
                    .Select(g =>
                    {
                        double totalTarget = g.Sum(x => x.total_target);
                        double totalPerf = g.Sum(x => x.total_perf);
                        double totalPerfAfter = g.Sum(x => x.total_perf_after_discount);

                        return new Per_Achieve
                        {
                            month = g.Key,
                            U_BC_NEW = "BC0100",
                            total_target = totalTarget,
                            total_perf = totalPerf,
                            total_perf_after_discount = totalPerfAfter,
                            achieve_rate = totalTarget != 0 ? $"{(totalPerf / totalTarget * 100):F2}%" : "0.00%",
                            achieve_rate_after_discount = totalTarget != 0 ? $"{(totalPerfAfter / totalTarget * 100):F2}%" : "0.00%",
                            Subord = g.Sum(x => x.Subord),
                            Leader = g.Sum(x => x.Leader)
                        };
                    }).ToList();
                WriteRegionData(worksheet, taipeiList, 1, 1, "台北", headers, fieldsWithWan, percentFields);
                #endregion

                #region 台中
                var taichungList = achieveList
                    .Where(x => x.U_BC_NEW == "BC0300")
                    .OrderBy(x => x.month)
                    .Select(g =>
                    {
                        double totalTarget = g.total_target;
                        double totalPerf = g.total_perf;
                        double totalPerfAfter = g.total_perf_after_discount;

                        return new Per_Achieve
                        {
                            month = g.month,
                            U_BC_NEW = g.U_BC_NEW,
                            total_target = totalTarget,
                            total_perf = totalPerf,
                            total_perf_after_discount = totalPerfAfter,
                            achieve_rate = totalTarget != 0 ? $"{(totalPerf / totalTarget * 100):F2}%" : "0.00%",
                            achieve_rate_after_discount = totalTarget != 0 ? $"{(totalPerfAfter / totalTarget * 100):F2}%" : "0.00%",
                            Subord = g.Subord,
                            Leader = g.Leader
                        };
                    }).ToList();
                WriteRegionData(worksheet, taichungList, 1, 8, "台中", headers, fieldsWithWan, percentFields);
                #endregion

                #region 台北一區(覃學華)
                var taipei_1_List = achieveList
                    .Where(x => x.U_BC_NEW.Equals("BC0100-1"))
                    .OrderBy(q => q.month)
                    .Select(g =>
                    {
                        double totalTarget = g.total_target;
                        double totalPerf = g.total_perf;
                        double totalPerfAfter = g.total_perf_after_discount;

                        return new Per_Achieve
                        {
                            month = g.month,
                            U_BC_NEW = g.U_BC_NEW,
                            total_target = totalTarget,
                            total_perf = totalPerf,
                            total_perf_after_discount = totalPerfAfter,
                            achieve_rate = totalTarget != 0 ? $"{(totalPerf / totalTarget * 100):F2}%" : "0.00%",
                            achieve_rate_after_discount = totalTarget != 0 ? $"{(totalPerfAfter / totalTarget * 100):F2}%" : "0.00%",
                            Subord = g.Subord,
                            Leader = g.Leader
                        };
                    }).ToList();
                WriteRegionData(worksheet, taipei_1_List, 1, 15, "台北一區(覃學華)", headers, fieldsWithWan, percentFields);
                #endregion

                #region 新北
                var newTaipeiList = achieveList
                    .Where(x => x.U_BC_NEW.Equals("BC0200"))
                    .OrderBy(q => q.month)
                    .Select(g =>
                    {
                        double totalTarget = g.total_target;
                        double totalPerf = g.total_perf;
                        double totalPerfAfter = g.total_perf_after_discount;

                        return new Per_Achieve
                        {
                            month = g.month,
                            U_BC_NEW = g.U_BC_NEW,
                            total_target = totalTarget,
                            total_perf = totalPerf,
                            total_perf_after_discount = totalPerfAfter,
                            achieve_rate = totalTarget != 0 ? $"{(totalPerf / totalTarget * 100):F2}%" : "0.00%",
                            achieve_rate_after_discount = totalTarget != 0 ? $"{(totalPerfAfter / totalTarget * 100):F2}%" : "0.00%",
                            Subord = g.Subord,
                            Leader = g.Leader
                        };
                    }).ToList();
                WriteRegionData(worksheet, newTaipeiList, 11, 1, "新北", headers, fieldsWithWan, percentFields);
                #endregion

                #region 台南
                var tainanList = achieveList
                    .Where(x => x.U_BC_NEW.Equals("BC0500"))
                    .OrderBy(q => q.month)
                    .Select(g =>
                    {
                        double totalTarget = g.total_target;
                        double totalPerf = g.total_perf;
                        double totalPerfAfter = g.total_perf_after_discount;

                        return new Per_Achieve
                        {
                            month = g.month,
                            U_BC_NEW = g.U_BC_NEW,
                            total_target = totalTarget,
                            total_perf = totalPerf,
                            total_perf_after_discount = totalPerfAfter,
                            achieve_rate = totalTarget != 0 ? $"{(totalPerf / totalTarget * 100):F2}%" : "0.00%",
                            achieve_rate_after_discount = totalTarget != 0 ? $"{(totalPerfAfter / totalTarget * 100):F2}%" : "0.00%",
                            Subord = g.Subord,
                            Leader = g.Leader
                        };
                    }).ToList();
                WriteRegionData(worksheet, tainanList, 11, 8, "台南", headers, fieldsWithWan, percentFields);
                #endregion

                #region 台北二區(李詩慧)
                var taipei_2_List = achieveList
                    .Where(x => x.U_BC_NEW.Equals("BC0100-2"))
                    .OrderBy(q => q.month)
                    .Select(g =>
                    {
                        double totalTarget = g.total_target;
                        double totalPerf = g.total_perf;
                        double totalPerfAfter = g.total_perf_after_discount;

                        return new Per_Achieve
                        {
                            month = g.month,
                            U_BC_NEW = g.U_BC_NEW,
                            total_target = totalTarget,
                            total_perf = totalPerf,
                            total_perf_after_discount = totalPerfAfter,
                            achieve_rate = totalTarget != 0 ? $"{(totalPerf / totalTarget * 100):F2}%" : "0.00%",
                            achieve_rate_after_discount = totalTarget != 0 ? $"{(totalPerfAfter / totalTarget * 100):F2}%" : "0.00%",
                            Subord = g.Subord,
                            Leader = g.Leader
                        };
                    }).ToList();
                WriteRegionData(worksheet, taipei_2_List, 11, 15, "台北二區(李詩慧)", headers, fieldsWithWan, percentFields);
                #endregion

                #region 桃園
                var taoyuanList = achieveList
                    .Where(x => x.U_BC_NEW.Equals("BC0600"))
                    .OrderBy(q => q.month)
                    .Select(g =>
                    {
                        double totalTarget = g.total_target;
                        double totalPerf = g.total_perf;
                        double totalPerfAfter = g.total_perf_after_discount;

                        return new Per_Achieve
                        {
                            month = g.month,
                            U_BC_NEW = g.U_BC_NEW,
                            total_target = totalTarget,
                            total_perf = totalPerf,
                            total_perf_after_discount = totalPerfAfter,
                            achieve_rate = totalTarget != 0 ? $"{(totalPerf / totalTarget * 100):F2}%" : "0.00%",
                            achieve_rate_after_discount = totalTarget != 0 ? $"{(totalPerfAfter / totalTarget * 100):F2}%" : "0.00%",
                            Subord = g.Subord,
                            Leader = g.Leader
                        };
                    }).ToList();
                WriteRegionData(worksheet, taoyuanList, 21, 1, "桃園", headers, fieldsWithWan, percentFields);
                #endregion

                #region 高雄
                var kaohsiungList = achieveList
                    .Where(x => x.U_BC_NEW.Equals("BC0400"))
                    .OrderBy(q => q.month)
                    .Select(g =>
                    {
                        double totalTarget = g.total_target;
                        double totalPerf = g.total_perf;
                        double totalPerfAfter = g.total_perf_after_discount;

                        return new Per_Achieve
                        {
                            month = g.month,
                            U_BC_NEW = g.U_BC_NEW,
                            total_target = totalTarget,
                            total_perf = totalPerf,
                            total_perf_after_discount = totalPerfAfter,
                            achieve_rate = totalTarget != 0 ? $"{(totalPerf / totalTarget * 100):F2}%" : "0.00%",
                            achieve_rate_after_discount = totalTarget != 0 ? $"{(totalPerfAfter / totalTarget * 100):F2}%" : "0.00%",
                            Subord = g.Subord,
                            Leader = g.Leader
                        };
                    }).ToList();
                WriteRegionData(worksheet, kaohsiungList, 21, 8, "高雄", headers, fieldsWithWan, percentFields);
                #endregion

                return package.GetAsByteArray();
            }
        }

        private static void WriteRegionData(ExcelWorksheet worksheet, List<Per_Achieve> dataList, int startRow, int startCol, string regionTitle,
            Dictionary<string, string> headers, HashSet<string> fieldsWithWan, HashSet<string> percentFields)
        {
            worksheet.Cells[startRow, startCol].Value = regionTitle;
            worksheet.Cells[startRow, startCol, startRow, startCol + headers.Count - 1].Merge = true;
            worksheet.Cells[startRow, startCol, startRow, startCol + headers.Count - 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[startRow, startCol, startRow, startCol + headers.Count - 1].Style.Font.Bold = true;

            int rowIndex = startRow + 1;
            int colIndex = startCol;

            // 表頭
            foreach (var header in headers)
            {
                var cell = worksheet.Cells[rowIndex, colIndex];
                cell.Value = header.Value;
                cell.Style.Font.Bold = true;
                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                cell.AutoFitColumns();
                colIndex++;
            }

            // 資料列
            rowIndex++;
            foreach (var item in dataList)
            {
                colIndex = startCol;
                foreach (var key in headers.Keys)
                {
                    var prop = typeof(Per_Achieve).GetProperty(key, BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);
                    if (prop == null) continue;

                    var value = prop.GetValue(item);
                    var cell = worksheet.Cells[rowIndex, colIndex];

                    if (fieldsWithWan.Contains(key))
                        cell.Value = $"{value}萬";
                    else if (percentFields.Contains(key))
                    {
                        double percent = Convert.ToDouble(value.ToString().Replace("%", "")) / 100;
                        cell.Value = percent;
                        cell.Style.Numberformat.Format = "0.00%";
                    }
                    else if (key == "month")
                    {
                        int month = int.Parse(value.ToString().Split('-')[1]);
                        cell.Value = $"{month}月({item.Leader}+{item.Subord})";
                    }
                    else
                        cell.Value = value;

                    cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    colIndex++;
                }
                rowIndex++;
            }

            // 總和列
            int summaryRow = rowIndex;
            colIndex = startCol;
            worksheet.Cells[summaryRow, colIndex].Value = "總和";
            worksheet.Cells[summaryRow, colIndex].Style.Font.Bold = true;
            worksheet.Cells[summaryRow, colIndex].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            double totalTargetSum = dataList.Sum(x => x.total_target);
            double totalPerfSum = dataList.Sum(x => x.total_perf);
            double totalPerfAfterSum = dataList.Sum(x => x.total_perf_after_discount);

            colIndex++;
            foreach (var key in headers.Keys.Skip(1))
            {
                var prop = typeof(Per_Achieve).GetProperty(key, BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);
                if (prop == null) continue;

                var cell = worksheet.Cells[summaryRow, colIndex];
                cell.Style.Font.Bold = true;
                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                if (fieldsWithWan.Contains(key))
                {
                    double sum = dataList.Sum(x => Convert.ToDouble(prop.GetValue(x) ?? 0));
                    cell.Value = $"{sum}萬";
                    cell.Style.Numberformat.Format = "#,##";
                }
                else if (key == "achieve_rate")
                {
                    double rate = totalTargetSum != 0 ? totalPerfSum / totalTargetSum : 0;
                    cell.Value = rate;
                    cell.Style.Numberformat.Format = "0.00%";
                }
                else if (key == "achieve_rate_after_discount")
                {
                    double rateAfter = totalTargetSum != 0 ? totalPerfAfterSum / totalTargetSum : 0;
                    cell.Value = rateAfter;
                    cell.Style.Numberformat.Format = "0.00%";
                }

                colIndex++;
            }

            // 邊框
            var dataRange = worksheet.Cells[startRow, startCol, summaryRow, startCol + headers.Count - 1];
            dataRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            dataRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            dataRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            dataRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
        }

        /// <summary>
        /// 產生用於偵錯和記錄的完整 T-SQL 語法。
        /// </summary>
        /// <param name="sqlTemplate">包含參數預留位置的 SQL 樣板。</param>
        /// <param name="parameters">SQL 參數集合。</param>
        /// <returns>組合完成的 T-SQL 字串。</returns>
        public static string GenerateDebugSql(string sqlTemplate, IEnumerable<SqlParameter> parameters)
        {
            // 使用 StringBuilder 提高字串操作效率
            var sb = new StringBuilder(sqlTemplate);

            // 先依據參數名稱長度由長到短排序，避免參數名稱是另一參數的子字串時發生錯誤
            // 例如：先替換 @p10 再替換 @p1
            foreach (var p in parameters.OrderByDescending(p => p.ParameterName.Length))
            {
                string valueToReplace;
                object pValue = p.Value;

                if (pValue == null || pValue == DBNull.Value)
                {
                    valueToReplace = "NULL";
                }
                else
                {
                    var type = pValue.GetType();

                    // 處理需要加單引號的類型
                    if (type == typeof(string) || type == typeof(Guid))
                    {
                        // 將字串中的單引號逸出，變成兩個單引號
                        valueToReplace = $"'{pValue.ToString().Replace("'", "''")}'";
                    }
                    else if (type == typeof(DateTime))
                    {
                        // 使用標準格式，避免地區設定問題
                        valueToReplace = $"'{((DateTime)pValue).ToString("yyyy-MM-dd HH:mm:ss.fff")}'";
                    }
                    else if (type == typeof(bool))
                    {
                        valueToReplace = (bool)pValue ? "1" : "0";
                    }
                    else // 處理數字等其他不需要單引號的類型
                    {
                        valueToReplace = pValue.ToString();
                    }
                }

                // 將 SQL 樣板中的參數名稱替換為格式化後的值
                sb.Replace(p.ParameterName, valueToReplace);
            }

            return sb.ToString();
        }
    }
}