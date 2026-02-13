using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using OfficeOpenXml; // 需引用 EPPlus
using OfficeOpenXml.Style;
namespace KF_WebAPI.FunctionHandler
{
    public static class ExcelExportHelper
    {
        public static byte[] ExportSalesDataToExcel(string BaseDate, DataTable dataTable, ArrayList sheetNames)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // EPPlus 5.0+ 需設定

            using (var package = new ExcelPackage())
            {
                DateTime currentMonth = DateTime.Parse(BaseDate);
                DateTime prevMonth = currentMonth.AddMonths(-1);

                foreach (string sheetName in sheetNames)
                {
                    var ws = package.Workbook.Worksheets.Add(sheetName);

                    // 1. 根據工作表名稱決定欄位前綴 (Prefix)
                    string prefix = "";
                    if (sheetName == "機車貸") prefix = "Engine_";
                    else if (sheetName == "汽車貸") prefix = "Car_";
                    // 房貸不需要前綴，保持空字串

                    // 2. 準備欄位映射 (與您提供的邏輯一致)
                    var mapping = new Dictionary<string, (string current, string pre)>
                {
                    { "估價", (prefix + "Rate", prefix + "Pre_Rate") },
                    { "進件", (prefix + "month_incase", prefix + "Pre_month_incase") },
                    { "撥款件數", (prefix + "month_get_amount_num", prefix + "Pre_month_get_amount_num") },
                    { "已撥金額", (prefix + "month_get_amount", prefix + "Pre_month_get_amount") },
                    { "未撥金額", (prefix + "month_pass_amount", prefix + "Pre_month_pass_amount") }
                };

                    int currentRow = 1;

                    // 3. 處理「全區」(加總)
                    WriteRegionBlock(ws, ref currentRow, "全區", currentMonth, prevMonth, mapping, dataTable, isTotal: true);

                    // 4. 處理各別地區 (台北, 新北, ...)
                    foreach (DataRow row in dataTable.Rows)
                    {
                        string regionName = row["UC_Na"]?.ToString();
                        if (string.IsNullOrEmpty(regionName)) continue;

                        WriteRegionBlock(ws, ref currentRow, regionName, currentMonth, prevMonth, mapping, dataTable, isTotal: false, specificRow: row);
                    }

                    // 自動調整欄寬
                    ws.Cells.AutoFitColumns();
                }

                return package.GetAsByteArray();
            }
        }

        /// <summary>
        /// 寫入一個地區的數據區塊 (標題+兩列數據)
        /// </summary>
        private static void WriteRegionBlock(ExcelWorksheet ws, ref int startRow, string regionName,
            DateTime currDate, DateTime preDate, Dictionary<string, (string current, string pre)> mapping,
            DataTable dt, bool isTotal, DataRow specificRow = null)
        {
            // A. 寫入地區名稱 (Header)
            ws.Cells[startRow, 1].Value = regionName;
            ws.Cells[startRow, 1, startRow, 6].Merge = true;
            ws.Cells[startRow, 1].Style.Font.Bold = true;
            ws.Cells[startRow, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[startRow, 1].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
            startRow++;

            // B. 寫入欄位標題
            ws.Cells[startRow, 1].Value = "日期";
            int col = 2;
            foreach (var key in mapping.Keys)
            {
                ws.Cells[startRow, col++].Value = key;
            }
            ws.Cells[startRow, 1, startRow, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[startRow, 1, startRow, 6].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(235, 235, 235));
            startRow++;

            // C. 寫入數據列 (本月與上月)
            for (int i = 0; i < 2; i++)
            {
                bool isCurrent = (i == 0);
                ws.Cells[startRow, 1].Value = isCurrent ? currDate.ToString("yyyy-MM-dd") : preDate.ToString("yyyy-MM-dd");

                int dataCol = 2;
                foreach (var map in mapping.Values)
                {
                    string fieldName = isCurrent ? map.current : map.pre;

                    if (isTotal)
                    {
                        // 全區：加總該欄位
                        ws.Cells[startRow, dataCol++].Value = dt.AsEnumerable().Sum(r => Convert.ToDecimal(r[fieldName] == DBNull.Value ? 0 : r[fieldName]));
                    }
                    else
                    {
                        // 特定地區：直接取值
                        ws.Cells[startRow, dataCol++].Value = specificRow[fieldName] == DBNull.Value ? 0 : specificRow[fieldName];
                    }
                }
                startRow++;
            }

            // D. 留一個空行
            startRow++;
        }



        /// <summary>
        /// 將 DataSet 轉成 Excel 檔案，每個 DataTable 對應一個 Sheet
        /// </summary>
        /// <param name="dataSet">包含多個 DataTable 的 DataSet</param>
        /// <param name="sheetNames">ArrayList，紀錄每個 Sheet 的名稱</param>
        /// <returns>Excel 檔案的 byte[]，可用於下載</returns>
        public static byte[] ExportDailyReportToExcel(DataSet dataSet, ArrayList sheetNames,bool isSum=true)
        {
            if (dataSet == null || dataSet.Tables.Count == 0)
                throw new ArgumentException("DataSet 不可為空");

            if (sheetNames == null || sheetNames.Count < dataSet.Tables.Count)
                throw new ArgumentException("Sheet 名稱數量不足");

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // 必須設定 LicenseContext

            using (var package = new ExcelPackage())
            {
                for (int i = 0; i < dataSet.Tables.Count; i++)
                {
                    var table = dataSet.Tables[i];
                    string sheetName = sheetNames[i].ToString();
                    ArrayList arrFromRow = new ArrayList();
                    ArrayList arrToRow = new ArrayList();
                    var worksheet = package.Workbook.Worksheets.Add(sheetName);

                    // 輸出資料列
                    for (int row = 0; row < table.Rows.Count; row++)
                    {
                        if(isSum)
                        {
                            if (table.Rows[row]["U_PFT_name"].ToString() == "" && table.Rows[row]["plan_name"].ToString() != "" && table.Rows[row]["plan_name"].ToString() != "合計" && table.Rows[row]["plan_name"].ToString() != "總計")
                            {
                                arrFromRow.Add(row + 2);
                            }
                            if (table.Rows[row]["U_PFT_name"].ToString() == "" && table.Rows[row]["plan_name"].ToString() != "" && table.Rows[row]["plan_name"].ToString() == "合計")
                            {
                                arrToRow.Add(row + 1);
                            }
                        }
                       

                        for (int col = 0; col < table.Columns.Count; col++)
                        {
                            var cellValue = table.Rows[row][col];

                            var cell = worksheet.Cells[row + 1, col + 1];

                            if (cellValue != DBNull.Value)
                            {
                                // 嘗試轉成數字
                                if (double.TryParse(cellValue.ToString(), out double numericValue))
                                {
                                    cell.Value = numericValue;

                                    // 如果大於 1000 → 套用三位一撇格式
                                    if (numericValue >= 1000)
                                    {
                                        cell.Style.Numberformat.Format = "#,##0"; // 整數三位一撇
                                    }
                                    else
                                    {
                                        cell.Style.Numberformat.Format = "0"; // 一般數字格式
                                    }
                                }
                                else
                                {
                                    // 非數字 → 原樣輸出
                                    cell.Value = cellValue;
                                }
                            }
                            else
                            {
                                cell.Value = null;
                            }
                        }
                    }
                    if (isSum)
                    {
                        foreach (Int32 FromRow in arrFromRow)
                        {//FromRow , FromCol,ToRow , ToCol
                            var range = worksheet.Cells[FromRow, 1, FromRow, table.Columns.Count];
                            range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            range.Style.Border.Top.Style = ExcelBorderStyle.Medium;
                            range = worksheet.Cells[FromRow - 1, 1, FromRow, table.Columns.Count];
                            range.Style.Font.Bold = true;
                        }
                        foreach (Int32 ToRow in arrToRow)
                        {//FromRow , FromCol,ToRow , ToCol
                            var range = worksheet.Cells[ToRow, 1, ToRow, table.Columns.Count];
                            range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            range.Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                            range.Style.Font.Bold = true;
                        }

                        var rangeEnd = worksheet.Cells[table.Rows.Count, 1, table.Rows.Count, table.Columns.Count];
                        rangeEnd.Style.Border.Top.Style = ExcelBorderStyle.Medium;
                        rangeEnd.Style.Font.Bold = true;
                    }
                   

                    worksheet.Cells.AutoFitColumns();
                }

                return package.GetAsByteArray();
            }
        }

    }
}