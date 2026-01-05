using KF_WebAPI.BaseClass;
using KF_WebAPI.BaseClass.AE;
using KF_WebAPI.FunctionHandler;
using Microsoft.AspNetCore.Http;
using Microsoft.Data.SqlClient;
using OfficeOpenXml;
using System.Data;
using System.Drawing;
using System.Reflection.PortableExecutable;

namespace KF_WebAPI.DataLogic
{
    public class AE_Rpt
    {
        ADOData _adoData = new ADOData();

        /// <summary>
        /// 取得機車分期總表
        /// </summary>
        public List<MotocaseSummary> GetMotoSummaryList()
        {
            try
            {
                var T_SQL_SP = @"exec GetMotocaseSummary ";
                var result = _adoData.ExecuteSQuery(T_SQL_SP).AsEnumerable().Select(row => new MotocaseSummary
                {
                    YYYYMM = row.Field<string>("YYYYMM"),
                    SendCount = row.Field<int>("SendCount"),
                    PassCount = row.Field<int>("PassCount"),
                    GetCount = row.Field<int>("GetCount"),
                    PassAmount = row.Field<decimal>("PassAmount"),
                    GetAmount = row.Field<decimal>("GetAmount"),
                    PerAmount = row.Field<decimal>("PerAmount"),
                    RemAmount = row.Field<decimal>("RemAmount"),
                    SettCount = row.Field<int>("SettCount"),
                    SettAmount = row.Field<decimal>("SettAmount"),
                    BadCount = row.Field<int>("BadCount"),
                    BadAmount = row.Field<decimal>("BadAmount"),
                    PassRate = ((decimal)row.Field<int>("PassCount") / row.Field<int>("SendCount")).ToString("0.00%"),
                    GetRate = ((decimal)row.Field<int>("GetCount") / row.Field<int>("PassCount")).ToString("0.00%")
                }).ToList();


                return result;
            }
            catch
            {
                throw ;
            }
        }

        /// <summary>
        /// 取得各項機車專案分期總表
        /// </summary>
        public List<MotocaseSummary> GetProjectMoto(string project)
        {
            try
            {
                var T_SQL_SP = @"exec GetProjectMonthlyReport @project";
                var parameters = new List<SqlParameter>
                {
                    new SqlParameter("@project",project)
                };
                var result = _adoData.ExecuteQuery(T_SQL_SP, parameters).AsEnumerable().Select(row =>
                {
                    int sendCount = row.Field<int?>("SendCount") ?? 0;
                    int passCount = row.Field<int?>("PassCount") ?? 0;
                    int getCount = row.Field<int?>("GetCount") ?? 0;
                    decimal passAmount = row.Field<decimal?>("PassAmount") ?? 0;
                    decimal getAmount = row.Field<decimal?>("GetAmount") ?? 0;
                    return new MotocaseSummary
                    {
                        YYYYMM = row.Field<string>("YYYYMM"),
                
                        SendCount = sendCount,
                        PassCount = passCount,
                        GetCount = getCount,
                
                        PassAmount = passAmount,
                        GetAmount = getAmount,
                
                        PassRate = sendCount == 0
                            ? "0.00%"
                            : ((decimal)passCount / sendCount).ToString("0.00%"),
                
                        GetRate = passCount == 0
                            ? "0.00%"
                            : ((decimal)getCount / passCount).ToString("0.00%")
                    };
                }).ToList();

                return result;
            }
            catch 
            {
                throw;
            }
        }

        /// <summary>
        /// 匯出機車分期總表EXCEL
        /// </summary>
        public byte[] GetMotoSummaryExcel()
        {
            try
            {
                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("總表");

                    #region 機車貸A&B彙總
                    var mtoSummaryList = GetMotoSummaryList();

                    string[] headers = { "進件數", "核准數", "撥款數", "核准總額", "撥款總額", "呆帳準備金", "目前月付金總額", "本金餘額", "已清償"
                            , "已清償金額", "呆帳", "呆帳金額", "核准率", "動撥率" };

                    // 添加合併標題
                    worksheet.Cells[1, 1].Value = "機車分期付款專案彙整表";
                    worksheet.Cells[1, 1, 1, headers.Length + 1].Merge = true; // 合併儲存格
                    worksheet.Cells[1, 1, 1, headers.Length + 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center; // 標題居中
                    worksheet.Cells[1, 1, 1, headers.Length + 1].Style.Font.Bold = true; // 標題加粗

                    // 添加子標題
                    int colIndex = 2;
                    foreach (var header in headers)
                    {
                        var cell = worksheet.Cells[2, colIndex++];
                        cell.Value = header;
                        // 設置儲存格底色為淺藍色
                        cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        cell.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                    }

                    int totalSendCount = 0;
                    int totalPassCount = 0;
                    int totalGetCount = 0;
                    decimal totalPassAmount = 0;
                    decimal totalGetAmount = 0;
                    decimal totalBadReAmount = 0;
                    decimal totalPerAmount = 0;
                    decimal totalRemAmount = 0;
                    int totalSettCount = 0;
                    decimal totalSettAmount = 0;
                    int totalBadCount = 0;
                    decimal totalBadAmount = 0;

                    // 添加表身
                    int rowIndex = 2;
                    int col = 1;
                    foreach (var item in mtoSummaryList)
                    {
                        rowIndex++;
                        col = 1;
                        var parts = item.YYYYMM.Split('-');
                        string month = parts[1].TrimStart('0');
                        int year = int.Parse(parts[0]) - 1911;
                        worksheet.Cells[rowIndex, col++].Value = $"{year}年{month}月";
                        worksheet.Cells[rowIndex, col++].Value = item.SendCount;
                        worksheet.Cells[rowIndex, col++].Value = item.PassCount;
                        worksheet.Cells[rowIndex, col++].Value = item.GetCount;
                        worksheet.Cells[rowIndex, col++].Value = item.PassAmount + "萬";
                        worksheet.Cells[rowIndex, col++].Value = item.GetAmount + "萬";
                        worksheet.Cells[rowIndex, col++].Value = item.GetAmount * 0.03m + "萬";
                        worksheet.Cells[rowIndex, col].Value = item.PerAmount;
                        worksheet.Cells[rowIndex, col].Style.Numberformat.Format = "#,##0\"元\"";
                        col++;
                        worksheet.Cells[rowIndex, col].Value = item.RemAmount;
                        worksheet.Cells[rowIndex, col].Style.Numberformat.Format = "#,##0\"元\"";
                        col++;
                        worksheet.Cells[rowIndex, col++].Value = item.SettCount;
                        worksheet.Cells[rowIndex, col].Value = item.SettAmount;
                        worksheet.Cells[rowIndex, col].Style.Numberformat.Format = item.SettAmount > 0 ? "0\"萬\"" : "0";
                        col++;
                        worksheet.Cells[rowIndex, col++].Value = item.BadCount;
                        worksheet.Cells[rowIndex, col].Value = item.BadAmount;
                        worksheet.Cells[rowIndex, col].Style.Numberformat.Format = item.BadAmount > 0 ? "0\"萬\"" : "0";
                        col++;
                        worksheet.Cells[rowIndex, col++].Value = item.PassRate;
                        worksheet.Cells[rowIndex, col++].Value = item.GetRate;

                        totalSendCount += item.SendCount;
                        totalPassCount += item.PassCount;
                        totalGetCount += item.GetCount;
                        totalPassAmount += item.PassAmount;
                        totalGetAmount += item.GetAmount;
                        totalBadReAmount += item.GetAmount * 0.03m;
                        totalPerAmount += item.PerAmount;
                        totalRemAmount += item.RemAmount;
                        totalSettCount += item.SettCount;
                        totalSettAmount += item.SettAmount;
                        totalBadCount += item.BadCount;
                        totalBadAmount += item.BadAmount;
                    }

                    //合計列
                    rowIndex++;
                    col = 1;
                    worksheet.Cells[rowIndex, col++].Value = "合計";
                    worksheet.Cells[rowIndex, col++].Value = totalSendCount;
                    worksheet.Cells[rowIndex, col++].Value = totalPassCount;
                    worksheet.Cells[rowIndex, col++].Value = totalGetCount;
                    worksheet.Cells[rowIndex, col++].Value = totalPassAmount + "萬";
                    worksheet.Cells[rowIndex, col++].Value = totalGetAmount + "萬";
                    worksheet.Cells[rowIndex, col++].Value = totalBadReAmount + "萬";
                    worksheet.Cells[rowIndex, col].Value = totalPerAmount;
                    worksheet.Cells[rowIndex, col].Style.Numberformat.Format = "#,##0\"元\"";
                    col++;
                    worksheet.Cells[rowIndex, col].Value = totalRemAmount;
                    worksheet.Cells[rowIndex, col].Style.Numberformat.Format = "#,##0\"元\"";
                    col++;
                    worksheet.Cells[rowIndex, col++].Value = totalSettCount;
                    worksheet.Cells[rowIndex, col].Value = totalSettAmount;
                    worksheet.Cells[rowIndex, col].Style.Numberformat.Format = totalSettAmount > 0 ? "0\"萬\"" : "0";
                    col++;
                    worksheet.Cells[rowIndex, col++].Value = totalBadCount;
                    worksheet.Cells[rowIndex, col].Value = totalBadAmount;
                    worksheet.Cells[rowIndex, col].Style.Numberformat.Format = totalBadAmount > 0 ? "0\"萬\"" : "0";
                    col++;
                    worksheet.Cells[rowIndex, col++].Value = "-";
                    worksheet.Cells[rowIndex, col++].Value = "-";


                    // 合計列樣式
                    using (var range = worksheet.Cells[rowIndex, 1, rowIndex, headers.Length + 1])
                    {
                        range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 217, 102));
                        range.Style.Font.Bold = true;
                        range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    }

                    // 框線
                    using (var range = worksheet.Cells[1, 1, rowIndex, headers.Length + 1])
                    {
                        range.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        range.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        range.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    }
                    #endregion

                    #region A&B總計
                    rowIndex += 2;
                    col = 1;
                    worksheet.Cells[rowIndex, col++].Value = "總核准率";
                    worksheet.Cells[rowIndex, col++].Value = ((decimal)totalPassCount / totalSendCount).ToString("0.00%");
                    col++;
                    worksheet.Cells[rowIndex, col++].Value = "總動撥率";
                    worksheet.Cells[rowIndex, col++].Value = ((decimal)totalGetCount / totalPassCount).ToString("0.00%");
                    col++;
                    worksheet.Cells[rowIndex, col, rowIndex, col + 1].Merge = true;
                    worksheet.Cells[rowIndex, col].Value = "已撥款總額(扣除已清償+呆帳)";
                    col += 2; // 合併兩欄後移到下一欄
                    worksheet.Cells[rowIndex, col++].Value = (totalGetAmount - totalSettAmount - totalBadAmount) + "萬";
                    col++;
                    worksheet.Cells[rowIndex, col++].Value = "製表日";
                    worksheet.Cells[rowIndex, col++].Value = DateTime.Now.ToString("yyy.M.d", new System.Globalization.CultureInfo("zh-TW"));
                    #endregion

                    rowIndex += 2;
                    string[] headersAB = { "進件數", "核准數", "撥款數", "核准總額", "撥款總額", "核准率", "動撥率" };

                    #region 機車貸A
                    var mtoSummaryListA = GetProjectMoto("PJ00046");
                    int rowIndexA = rowIndex;
                    // 添加合併標題
                    worksheet.Cells[rowIndex, 1].Value = "機車A專案彙整表";
                    worksheet.Cells[rowIndex, 1, rowIndex, headersAB.Length + 1].Merge = true; // 合併儲存格
                    worksheet.Cells[rowIndex, 1, rowIndex, headersAB.Length + 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center; // 標題居中
                    worksheet.Cells[rowIndex, 1, rowIndex, headersAB.Length + 1].Style.Font.Bold = true; // 標題加粗

                    // 添加子標題
                    colIndex = 2;
                    rowIndex++;
                    foreach (var header in headersAB)
                    {
                        var cell = worksheet.Cells[rowIndex, colIndex++];
                        cell.Value = header;
                        // 設置儲存格底色為淺藍色
                        cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        cell.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                    }

                    totalSendCount = 0;
                    totalPassCount = 0;
                    totalGetCount = 0;
                    totalPassAmount = 0;
                    totalGetAmount = 0;

                    // 添加表身
                    foreach (var item in mtoSummaryListA)
                    {
                        rowIndex++;
                        col = 1;
                        var parts = item.YYYYMM.Split('-');
                        string month = parts[1].TrimStart('0');
                        int year = int.Parse(parts[0]) - 1911;
                        worksheet.Cells[rowIndex, col++].Value = $"{year}年{month}月";
                        worksheet.Cells[rowIndex, col++].Value = item.SendCount;
                        worksheet.Cells[rowIndex, col++].Value = item.PassCount;
                        worksheet.Cells[rowIndex, col++].Value = item.GetCount;
                        worksheet.Cells[rowIndex, col++].Value = item.PassAmount + "萬";
                        worksheet.Cells[rowIndex, col++].Value = item.GetAmount + "萬";
                        worksheet.Cells[rowIndex, col++].Value = item.PassRate;
                        worksheet.Cells[rowIndex, col++].Value = item.GetRate;

                        totalSendCount += item.SendCount;
                        totalPassCount += item.PassCount;
                        totalGetCount += item.GetCount;
                        totalPassAmount += item.PassAmount;
                        totalGetAmount += item.GetAmount;
                    }

                    //合計列
                    rowIndex++;
                    col = 1;
                    worksheet.Cells[rowIndex, col++].Value = "合計";
                    worksheet.Cells[rowIndex, col++].Value = totalSendCount;
                    worksheet.Cells[rowIndex, col++].Value = totalPassCount;
                    worksheet.Cells[rowIndex, col++].Value = totalGetCount;
                    worksheet.Cells[rowIndex, col++].Value = totalPassAmount + "萬";
                    worksheet.Cells[rowIndex, col++].Value = totalGetAmount + "萬";
                    worksheet.Cells[rowIndex, col++].Value = "-";
                    worksheet.Cells[rowIndex, col++].Value = "-";

                    // 合計列樣式
                    using (var range = worksheet.Cells[rowIndex, 1, rowIndex, headersAB.Length + 1])
                    {
                        range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 217, 102));
                        range.Style.Font.Bold = true;
                        range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    }

                    // 框線
                    using (var range = worksheet.Cells[rowIndexA, 1, rowIndex, headersAB.Length + 1])
                    {
                        range.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        range.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        range.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    }
                    #endregion

                    rowIndex += 2;

                    #region 機車貸B
                    var mtoSummaryListB = GetProjectMoto("PJ00047");
                    int rowIndexB = rowIndex;
                    // 添加合併標題
                    worksheet.Cells[rowIndex, 1].Value = "機車B專案彙整表";
                    worksheet.Cells[rowIndex, 1, rowIndex, headersAB.Length + 1].Merge = true; // 合併儲存格
                    worksheet.Cells[rowIndex, 1, rowIndex, headersAB.Length + 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center; // 標題居中
                    worksheet.Cells[rowIndex, 1, rowIndex, headersAB.Length + 1].Style.Font.Bold = true; // 標題加粗

                    // 添加子標題
                    colIndex = 2;
                    rowIndex++;
                    foreach (var header in headersAB)
                    {
                        var cell = worksheet.Cells[rowIndex, colIndex++];
                        cell.Value = header;
                        // 設置儲存格底色為淺藍色
                        cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        cell.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                    }

                    totalSendCount = 0;
                    totalPassCount = 0;
                    totalGetCount = 0;
                    totalPassAmount = 0;
                    totalGetAmount = 0;

                    // 添加表身
                    foreach (var item in mtoSummaryListB)
                    {
                        rowIndex++;
                        col = 1;
                        var parts = item.YYYYMM.Split('-');
                        string month = parts[1].TrimStart('0');
                        int year = int.Parse(parts[0]) - 1911;
                        worksheet.Cells[rowIndex, col++].Value = $"{year}年{month}月";
                        worksheet.Cells[rowIndex, col++].Value = item.SendCount;
                        worksheet.Cells[rowIndex, col++].Value = item.PassCount;
                        worksheet.Cells[rowIndex, col++].Value = item.GetCount;
                        worksheet.Cells[rowIndex, col++].Value = item.PassAmount + "萬";
                        worksheet.Cells[rowIndex, col++].Value = item.GetAmount + "萬";
                        worksheet.Cells[rowIndex, col++].Value = item.PassRate;
                        worksheet.Cells[rowIndex, col++].Value = item.GetRate;

                        totalSendCount += item.SendCount;
                        totalPassCount += item.PassCount;
                        totalGetCount += item.GetCount;
                        totalPassAmount += item.PassAmount;
                        totalGetAmount += item.GetAmount;
                    }

                    //合計列
                    rowIndex++;
                    col = 1;
                    worksheet.Cells[rowIndex, col++].Value = "合計";
                    worksheet.Cells[rowIndex, col++].Value = totalSendCount;
                    worksheet.Cells[rowIndex, col++].Value = totalPassCount;
                    worksheet.Cells[rowIndex, col++].Value = totalGetCount;
                    worksheet.Cells[rowIndex, col++].Value = totalPassAmount + "萬";
                    worksheet.Cells[rowIndex, col++].Value = totalGetAmount + "萬";
                    worksheet.Cells[rowIndex, col++].Value = "-";
                    worksheet.Cells[rowIndex, col++].Value = "-";

                    // 合計列樣式
                    using (var range = worksheet.Cells[rowIndex, 1, rowIndex, headersAB.Length + 1])
                    {
                        range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 217, 102));
                        range.Style.Font.Bold = true;
                        range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    }

                    // 框線
                    using (var range = worksheet.Cells[rowIndexB, 1, rowIndex, headersAB.Length + 1])
                    {
                        range.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        range.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        range.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    }
                    #endregion


                    // 自動調整列寬
                    worksheet.Cells.AutoFitColumns();

                    return package.GetAsByteArray();
                }
            }
            catch
            {
                throw;
            }
        }
    }
}
