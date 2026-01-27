using KF_WebAPI.BaseClass;
using KF_WebAPI.BaseClass.AE;
using KF_WebAPI.FunctionHandler;
using Microsoft.AspNetCore.Http;
using Microsoft.Data.SqlClient;
using OfficeOpenXml;
using System;
using System.Collections;
using System.Data;
using System.Drawing;
using System.Dynamic;
using System.Reflection;
using System.Reflection.PortableExecutable;

namespace KF_WebAPI.DataLogic
{
    public class AE_Rpt
    {
        ADOData _adoData = new ADOData();

        /// <summary>
        /// 取得機車分期總表
        /// </summary>
        public List<MotocaseSummary> GetMotoSummaryList(Motocase_req model)
        {
            try
            {
                var T_SQL_SP = @"exec GetMotocaseSummary @checkDateS,@checkDateE";
                var parameters = new List<SqlParameter>
                {
                    new SqlParameter("@checkDateS",model.checkDateS),
                    new SqlParameter("@checkDateE",model.checkDateE)
                };
                var result = _adoData.ExecuteQuery(T_SQL_SP,parameters).AsEnumerable().Select(row => new MotocaseSummary
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
        public List<MotocaseSummary> GetProjectMoto(Motocase_req model)
        {
            try
            {
                var T_SQL_SP = @"exec GetProjectMonthlyReport @project,@checkDateS,@checkDateE";
                var parameters = new List<SqlParameter>
                {
                    new SqlParameter("@project",model.project),
                    new SqlParameter("@checkDateS",model.checkDateS),
                    new SqlParameter("@checkDateE",model.checkDateE)
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
        public byte[] GetMotoSummaryExcel(Motocase_req model)
        {
            try
            {
                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("總表");

                    #region 機車貸A&B彙總
                    var mtoSummaryList = GetMotoSummaryList(model);

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
                    model.project = "PJ00046";
                    var mtoSummaryListA = GetProjectMoto(model);
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
                    model.project = "PJ00047";
                    var mtoSummaryListB = GetProjectMoto(model);
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

        /// <summary>
        /// 取得汽車分期總表
        /// </summary>
        public List<CarcaseSummary> GetCarSummaryList(Carcase_req model)
        {
            try
            {
                var T_SQL_SP = @"exec GetCarcaseSummary @checkDateS,@checkDateE";
                var parameters = new List<SqlParameter>
                {
                    new SqlParameter("@checkDateS",model.checkDateS),
                    new SqlParameter("@checkDateE",model.checkDateE)
                };
                var result = _adoData.ExecuteQuery(T_SQL_SP, parameters).AsEnumerable().Select(row => new CarcaseSummary
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
                    PassRate = row.Field<int>("SendCount") == 0 ? "0.00%": ((decimal)row.Field<int>("PassCount") / row.Field<int>("SendCount")).ToString("0.00%"),
                    GetRate = row.Field<int>("PassCount") == 0 ? "0.00%": ((decimal)row.Field<int>("GetCount") / row.Field<int>("PassCount")).ToString("0.00%")
                
                }).ToList();


                return result;
            }
            catch
            {
                throw;
            }
        }

        /// <summary>
        /// 匯出汽車分期總表EXCEL
        /// </summary>
        public byte[] GetCarSummaryExcel(Carcase_req model)
        {
            try
            {
                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("總表");

                    #region 汽車貸彙總
                    var mtoSummaryList = GetCarSummaryList(model);

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

                    #region 總計
                    rowIndex += 2;
                    col = 1;
                    worksheet.Cells[rowIndex, col++].Value = "總核准率";
                    worksheet.Cells[rowIndex, col++].Value = totalSendCount == 0 ? "0.00%" : ((decimal)totalPassCount / totalSendCount).ToString("0.00%");
                    col++;
                    worksheet.Cells[rowIndex, col++].Value = "總動撥率";
                    worksheet.Cells[rowIndex, col++].Value = totalPassCount == 0 ? "0.00%" : ((decimal)totalGetCount / totalPassCount).ToString("0.00%");
                    col++;
                    worksheet.Cells[rowIndex, col, rowIndex, col + 1].Merge = true;
                    worksheet.Cells[rowIndex, col].Value = "已撥款總額(扣除已清償+呆帳)";
                    col += 2; // 合併兩欄後移到下一欄
                    worksheet.Cells[rowIndex, col++].Value = (totalGetAmount - totalSettAmount - totalBadAmount) + "萬";
                    col++;
                    worksheet.Cells[rowIndex, col++].Value = "製表日";
                    worksheet.Cells[rowIndex, col++].Value = DateTime.Now.ToString("yyy.M.d", new System.Globalization.CultureInfo("zh-TW"));
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

        /// <summary>
        /// 業績報表_日報表
        /// </summary>
        /// <param name="Base_Date"></param>
        /// <param name="isPreDay"></param>
        /// <returns></returns>
        public DataSet GetDailyReportByDate(string Base_Date, string U_BC, Boolean isPreDay)
        {
            DataSet DsResult = new DataSet();

            string ThisMon = "";
            string PreMon = "";
            string PreMon1 = "";
            string PE_Date = "";
            string T_SQL = "";
            string AddDate = "";
            if (Base_Date == "")
            {
                if (isPreDay)
                {
                    AddDate = DateTime.Now.AddDays(-1).ToString("yyyy/MM/dd");
                    Base_Date = DateTime.Now.AddDays(-1).ToString("yyyyMMdd");
                    ThisMon = DateTime.Now.AddDays(-1).ToString("yyyyMM");
                    PE_Date = DateTime.Now.AddDays(-1).ToString("yyyy-MM");
                    PreMon = Convert.ToDateTime(DateTime.Now.AddDays(-1)).AddMonths(-1).ToString("yyyyMM");
                    PreMon1 = Convert.ToDateTime(DateTime.Now.AddDays(-1)).AddMonths(-2).ToString("yyyyMM");
                }
                else
                {
                    AddDate = DateTime.Now.ToString("yyyy/MM/dd");
                    Base_Date = DateTime.Now.ToString("yyyyMMdd");
                    ThisMon = DateTime.Now.ToString("yyyyMM");
                    PE_Date = DateTime.Now.ToString("yyyy-MM");
                    PreMon = Convert.ToDateTime(DateTime.Now).AddMonths(-1).ToString("yyyyMM");
                    PreMon1 = Convert.ToDateTime(DateTime.Now).AddMonths(-2).ToString("yyyyMM");
                }


            }
            else
            {
                if (isPreDay)
                {
                    AddDate = (Convert.ToInt16(Base_Date.Split("/")[0]) + 1911).ToString() + "/" + Base_Date.Split("/")[1].ToString().PadLeft(2, '0') + "/" + Base_Date.Split("/")[2].ToString().PadLeft(2, '0');
                    AddDate = Convert.ToDateTime(AddDate).AddDays(-1).ToString("yyyy/MM/dd");

                    PreMon = Convert.ToDateTime(AddDate).AddMonths(-1).ToString("yyyyMM");
                    PreMon1 = Convert.ToDateTime(AddDate).AddMonths(-2).ToString("yyyyMM");
                    Base_Date = Convert.ToDateTime(AddDate).ToString("yyyyMMdd");
                    ThisMon = Base_Date.Substring(0, 6);
                    PE_Date = Base_Date.Substring(0, 4) + "-" + Base_Date.Substring(4, 2);
                }
                else
                {
                    AddDate = (Convert.ToInt16(Base_Date.Split("/")[0]) + 1911).ToString() + "/" + Base_Date.Split("/")[1].ToString().PadLeft(2, '0') + "/" + Base_Date.Split("/")[2].ToString().PadLeft(2, '0');
                    PreMon = Convert.ToDateTime(AddDate).AddMonths(-1).ToString("yyyyMM");
                    PreMon1 = Convert.ToDateTime(AddDate).AddMonths(-2).ToString("yyyyMM");
                    Base_Date = (Convert.ToInt16(Base_Date.Split("/")[0]) + 1911).ToString() + Base_Date.Split("/")[1].ToString().PadLeft(2, '0') + Base_Date.Split("/")[2].ToString().PadLeft(2, '0');
                    ThisMon = Base_Date.Substring(0, 6);
                    PE_Date = Base_Date.Substring(0, 4) + "-" + Base_Date.Substring(4, 2);
                }

            }
            ADOData _adoData = new ADOData();
            #region SQL-各區人員
            var parameters = new List<SqlParameter>();
            if (U_BC == "")
            {
                T_SQL = @"select isnull(G.Spec_Group, U_BC)U_BC,BC_Name,bc_sort,count(*) PelCount
                         from (select U_num,U_BC from User_M where U_BC between 'BC0100' and 'BC0600' and U_PFT in('PFT050','PFT030','PFT060','PFT300')  and  U_leave_date is null or convert(varchar, U_arrive_date, 112) =@ThisMon ) U
                        Left join  User_Spec_Group G  on U.U_num=G.U_num 
                        left join (
                        select item_D_code,item_D_name BC_Name,0 bc_sort from Item_list where item_M_code = 'Spec_Group' and item_M_type='N'
                        union all
                        select  item_D_code,item_D_name BC_Name,item_sort bc_sort from Item_list  where item_M_code = 'branch_company' and item_M_type='N'
                        ) BC on isnull(G.Spec_Group, U_BC)=BC.item_D_code  where isnull(G.Spec_Group, U_BC) <> 'BC0100' and isnull(G.Spec_Group, U_BC)<>'BC_CAR'
                        group by isnull(G.Spec_Group, U_BC),BC_Name,bc_sort
                        union all 
                        select U_BC,BC_Name,bc_sort,count(*)PelCount  from USER_M M Left Join
                        (select  item_D_code,item_D_name BC_Name,item_sort bc_sort from Item_list  where item_M_code = 'branch_company' and item_M_type='N' )
                        D on M.U_BC=D.item_D_code
                        where U_BC='BC0900' and U_num <> 'K9999' and U_PFT in('PFT050','PFT030','PFT060','PFT300')  and U_susp_date is null 
                        and U_leave_date is null or convert(varchar, U_arrive_date, 112) =@ThisMon
                        group by  U_BC,BC_Name,bc_sort 
                        union all 
                        select 'BC0901','湧立','999',1
                        order by bc_sort,isnull(G.Spec_Group, U_BC) ";

                parameters.Add(new SqlParameter("@ThisMon", ThisMon));
               
            }
            else
            {
                T_SQL = @"select isnull(G.Spec_Group, U_BC)U_BC,BC_Name,bc_sort,count(*) PelCount
                         from (select U_num,U_BC from User_M where U_BC=@U_BC and U_leave_date is null or convert(varchar, U_arrive_date, 112) = @ThisMon ) U
                        Left join  User_Spec_Group G  on U.U_num=G.U_num 
                        left join (
                        select item_D_code,item_D_name BC_Name,0 bc_sort from Item_list where item_M_code = 'Spec_Group' and item_M_type='N'
                        union all
                        select  item_D_code,item_D_name BC_Name,item_sort bc_sort from Item_list  where item_M_code = 'branch_company' and item_M_type='N'
                        ) BC on isnull(G.Spec_Group, U_BC)=BC.item_D_code 
                        group by isnull(G.Spec_Group, U_BC),BC_Name,bc_sort
                        order by bc_sort,isnull(G.Spec_Group, U_BC) ";

                parameters.Add(new SqlParameter("@ThisMon", ThisMon));
              
                parameters.Add(new SqlParameter("@U_BC", U_BC));
            }
            
           
            DataTable dtBCResult = _adoData.ExecuteQuery(T_SQL, parameters, true);
            ArrayList arrTitle = new ArrayList();
            foreach (DataRow dr in dtBCResult.Rows)
            {
                arrTitle.Add("國峯租賃 (" + dr["BC_Name"].ToString() + ") 1 + " + dr["PelCount"].ToString() + "人");
            }
            #endregion


            #region SQL-汽車貸人員
             parameters = new List<SqlParameter>();
            T_SQL = @"select M.u_num,U_name,M.U_PFT,BC_Name+PFT_Name PFT_Name from 
                    Item_list I
                    left join USER_M M on I.item_D_code=M.U_num
                    left join (select item_D_code U_BC,item_D_name BC_Name,item_sort BC_Sort from Item_list 
                    where item_M_code='branch_company'  and item_D_type='Y') BC on M.U_BC=BC.U_BC
                    left join  (select item_D_code U_PFT,item_D_name PFT_Name from Item_list 
                    where item_M_code='professional_title'  and item_D_type='Y')PFT on M.U_PFT=PFT.U_PFT
                    where item_M_code=@Car_Sales and item_D_type='Y'
                    order by BC_Sort ";
            parameters.Add(new SqlParameter("@Car_Sales", "Car_Sales"));

            DataTable dtCarResult = _adoData.ExecuteQuery(T_SQL, parameters);
            string CarTitle= "汽車購車改裝分期 " + dtCarResult.Rows.Count.ToString() + "人";
            ArrayList arrCar = new ArrayList();
           
            foreach (DataRow dr in dtCarResult.Rows)
            {
                arrCar.Add(dr["u_num"].ToString());
            }
            #endregion


            #region SQL-業績相關
            parameters = new List<SqlParameter>();
            T_SQL = @"
WITH A (bc_sort, U_BC, U_susp_date, is_susp, leader_name, plan_name, plan_num, group_id, group_M_id, group_M_title, U_PFT_sort, U_PFT_name,
day_incase_num_PJ00048,month_incase_num_PJ00048,day_get_amount_num_Car,day_get_amount_Car,month_pass_num_Car,month_get_amount_num_Car,month_get_amount_PJ00048,month_pass_amount_PJ00048,
day_incase_num_PJ00046, day_incase_num_PJ00047, month_incase_num_PJ00046, month_incase_num_PJ00047, day_get_amount_num_engine, day_get_amount_engine, month_pass_num_engine, month_get_amount_num_engine, month_get_amount_PJ00046, month_get_amount_PJ00047, month_pass_amount_PJ00046, month_pass_amount_PJ00047, day_incase_num_FDCOM001, month_incase_num_FDCOM001, day_incase_num_FDCOM003, month_incase_num_FDCOM003, day_incase_num_FDCOM003_1, month_incase_num_FDCOM003_1, day_incase_num_FDCOM004, month_incase_num_FDCOM004, day_incase_num_FDCOM005, month_incase_num_FDCOM005, day_get_amount_num, day_get_amount, month_pass_num, month_get_amount_num, month_get_amount_FDCOM001, month_get_amount_FDCOM003, month_get_amount_FDCOM003_1, month_get_amount_FDCOM004, month_get_amount_FDCOM005, month_pass_amount_FDCOM001, month_pre_amount_FDCOM001, month_pass_amount_FDCOM003, month_pass_amount_FDCOM003_1, month_pass_amount_FDCOM004, month_pass_amount_FDCOM005, advance_payment_AE) AS
  (SELECT leader.bc_sort,leader.U_BC,sa.U_susp_date,sa.is_susp,
          isnull(ug.group_M_name, '未分組') leader_name,
          ug.group_D_name,ug.group_D_code,ug.group_id,ug.group_M_id,
          ug.group_M_title,sa.U_PFT_sort,sa.U_PFT_name
        /* 汽車-日進件數 */ 
			 ,sum(CASE
                  WHEN project_title = 'PJ00048'
                       AND convert(varchar, Send_amount_date, 112) = @Base_date THEN 1
                  ELSE 0
              END) AS day_incase_num_PJ00048 
			  
			/* 汽車-月進件數*/ 
			 ,sum(CASE
                  WHEN project_title = 'PJ00048'
                       AND left(convert(varchar, Send_amount_date, 112), 6) = @ThisMon
                       AND convert(varchar, Send_amount_date, 112) <= @Base_date THEN 1
                  ELSE 0
              END) AS month_incase_num_PJ00048 
			/* 汽車-日撥款數*/ 
			  ,sum(CASE
                  WHEN project_title IN('PJ00048')
                       AND convert(varchar,get_amount_date, 112) = @Base_date THEN 1
                  ELSE 0
              END) AS day_get_amount_num_Car
			/* 汽車-日撥款額*/ 
			  ,sum(CASE
                  WHEN project_title IN('PJ00048')
                       AND convert(varchar,get_amount_date, 112) = @Base_date THEN get_amount
                  ELSE 0
              END) AS day_get_amount_Car
			/* 汽車-月核准數*/ 
              ,sum(CASE
                  WHEN project_title IN ('PJ00048')
                       AND left(convert(varchar, Send_amount_date, 112), 6)=@ThisMon
                       AND convert(varchar,Send_result_date, 112) <= @Base_date
                       AND Send_result_type IN ('SRT002', 'SRT005') THEN 1
                  ELSE 0
              END) AS month_pass_num_Car
			/* 汽車-月撥款數*/ 
              ,sum(CASE
                  WHEN project_title IN ('PJ00048')
                       AND left(convert(varchar, get_amount_date, 112), 6) = @ThisMon
                       AND convert(varchar,get_amount_date, 112) <= @Base_date
                       AND get_amount_type IN ('GTAT002') THEN 1
                  ELSE 0
              END) AS month_get_amount_num_Car
			/* 汽車-月撥款額*/ 
              ,sum(CASE
                  WHEN 'FDCOM003'=fund_company
                       AND project_title IN ('PJ00048')
                       AND left(convert(varchar, get_amount_date, 112), 6) = @ThisMon
                       AND convert(varchar,get_amount_date, 112) <= @Base_date THEN get_amount
                  ELSE 0
              END) AS month_get_amount_PJ00048
			/* 汽車-已核未撥*/ 
			  ,sum(CASE
                  WHEN 'FDCOM003'=fund_company
                       AND project_title IN ('PJ00048')
                       AND left(convert(varchar, Send_amount_date, 112), 6) IN (@ThisMon, @PreMon, @PreMon1)
                       AND convert(varchar,Send_result_date, 112) <= @Base_date
                       AND Send_result_type = 'SRT002'
                       AND isnull(check_amount_type, '') NOT IN ('CKAT003')
                       AND isnull(get_amount_type, '') NOT IN ('GTAT002', 'GTAT003') THEN pass_amount
                  ELSE 0
              END) AS month_pass_amount_PJ00048	
        /* 機車貸款A 日進件數 */ ,
          sum(CASE
                  WHEN project_title = 'PJ00046'
                       AND convert(varchar, Send_amount_date, 112) = @Base_date THEN 1
                  ELSE 0
              END) AS day_incase_num_PJ00046 /* 機車貸款B 日進件數 */ ,
          sum(CASE
                  WHEN project_title = 'PJ00047'
                       AND convert(varchar, Send_amount_date, 112) = @Base_date THEN 1
                  ELSE 0
              END) AS day_incase_num_PJ00047 /* 機車貸款A 月進件數*/ ,
          sum(CASE
                  WHEN project_title = 'PJ00046'
                       AND left(convert(varchar, Send_amount_date, 112), 6) = @ThisMon
                       AND convert(varchar, Send_amount_date, 112) <= @Base_date THEN 1
                  ELSE 0
              END) AS month_incase_num_PJ00046 /* 機車貸款B 月進件數*/ ,
          sum(CASE
                  WHEN project_title = 'PJ00047'
                       AND left(convert(varchar, Send_amount_date, 112), 6) = @ThisMon
                       AND convert(varchar, Send_amount_date, 112) <= @Base_date THEN 1
                  ELSE 0
              END) AS month_incase_num_PJ00047 /* 機車貸款-日撥款數*/ ,
          sum(CASE
                  WHEN project_title IN('PJ00046', 'PJ00047')
                       AND convert(varchar,get_amount_date, 112) = @Base_date THEN 1
                  ELSE 0
              END) AS day_get_amount_num_engine /* 機車貸款-日撥款額*/ ,
          sum(CASE
                  WHEN project_title IN('PJ00046', 'PJ00047')
                       AND convert(varchar,get_amount_date, 112) = @Base_date THEN get_amount
                  ELSE 0
              END) AS day_get_amount_engine /* 機車貸款-月核准數*/ ,
          sum(CASE
                  WHEN project_title IN ('PJ00046', 'PJ00047')
                       AND left(convert(varchar, Send_result_date, 112), 6)=@ThisMon
                       AND convert(varchar,Send_result_date, 112) <= @Base_date
                       AND Send_result_type IN ('SRT002', 'SRT005') THEN 1
                  ELSE 0
              END) AS month_pass_num_engine /* 機車貸款-月撥款數*/ ,
          sum(CASE
                  WHEN project_title IN ('PJ00046', 'PJ00047')
                       AND left(convert(varchar, get_amount_date, 112), 6) = @ThisMon
                       AND convert(varchar,get_amount_date, 112) <= @Base_date
                       AND get_amount_type IN ('GTAT002') THEN 1
                  ELSE 0
              END) AS month_get_amount_num_engine /* 機車貸款A; 月撥款額*/ ,
          sum(CASE
                  WHEN 'FDCOM003'=fund_company
                       AND project_title IN ('PJ00046')
                       AND left(convert(varchar, get_amount_date, 112), 6) = @ThisMon
                       AND convert(varchar,get_amount_date, 112) <= @Base_date THEN get_amount
                  ELSE 0
              END) AS month_get_amount_PJ00046 /* 機車貸款B; 月撥款額*/ ,
          sum(CASE
                  WHEN 'FDCOM003'=fund_company
                       AND project_title IN ('PJ00047')
                       AND left(convert(varchar, get_amount_date, 112), 6) = @ThisMon
                       AND convert(varchar,get_amount_date, 112) <= @Base_date THEN get_amount
                  ELSE 0
              END) AS month_get_amount_PJ00047 /* 機車貸款A; 已核未撥*/ ,
          sum(CASE
                  WHEN 'FDCOM003'=fund_company
                       AND project_title IN ('PJ00046')
                       AND left(convert(varchar, Send_amount_date, 112), 6) IN (@ThisMon, @PreMon, @PreMon1)
                       AND convert(varchar,Send_result_date, 112) <= @Base_date
                       AND Send_result_type = 'SRT002'
                       AND isnull(check_amount_type, '') NOT IN ('CKAT003')
                       AND isnull(get_amount_type, '') NOT IN ('GTAT002', 'GTAT003') THEN pass_amount
                  ELSE 0
              END) AS month_pass_amount_PJ00046 /* 機車貸款B; 已核未撥*/ ,
          sum(CASE
                  WHEN 'FDCOM003'=fund_company
                       AND project_title IN ('PJ00047')
                       AND left(convert(varchar, Send_amount_date, 112), 6) IN (@ThisMon, @PreMon, @PreMon1)
                       AND convert(varchar,Send_result_date, 112) <= @Base_date
                       AND Send_result_type = 'SRT002'
                       AND isnull(check_amount_type, '') NOT IN ('CKAT003')
                       AND isnull(get_amount_type, '') NOT IN ('GTAT002', 'GTAT003') THEN pass_amount
                  ELSE 0
              END) AS month_pass_amount_PJ00047 /* 新鑫 日進件數*/ ,
          sum(CASE
                  WHEN 'FDCOM001'=fund_company
                       AND convert(varchar, Send_amount_date, 112) = @Base_date THEN 1
                  ELSE 0
              END) AS day_incase_num_FDCOM001 /* 新鑫 月進件數*/ ,
          sum(CASE
                  WHEN 'FDCOM001'=fund_company
                       AND left(convert(varchar, Send_amount_date, 112), 6) = @ThisMon
                       AND convert(varchar, Send_amount_date, 112) <= @Base_date THEN 1
                  ELSE 0
              END) AS month_incase_num_FDCOM001 /* 國&#23791; 日進件數*/ ,
          sum(CASE
                  WHEN 'FDCOM003'=fund_company
                       AND project_title NOT IN ('PJ00046', 'PJ00047', 'PJ00048')
                       AND convert(varchar, Send_amount_date, 112) = @Base_date THEN 1
                  ELSE 0
              END) AS day_incase_num_FDCOM003 /* 國&#23791; 月進件數*/ ,
          sum(CASE
                  WHEN 'FDCOM003'=fund_company
                       AND project_title NOT IN ('PJ00046', 'PJ00047', 'PJ00048')
                       AND left(convert(varchar, Send_amount_date, 112), 6) = @ThisMon
                       AND convert(varchar, Send_amount_date, 112) <= @Base_date THEN 1
                  ELSE 0
              END) AS month_incase_num_FDCOM003 /* 國&#23791; 日進件數*/ ,
          sum(CASE
                  WHEN 'FDCOM003_1'=fund_company
                       AND convert(varchar, Send_amount_date, 112) = @Base_date THEN 1
                  ELSE 0
              END) AS day_incase_num_FDCOM003_1 /* 國&#23791; 月進件數*/ ,
          sum(CASE
                  WHEN 'FDCOM003_1'=fund_company
                       AND left(convert(varchar, Send_amount_date, 112), 6) = @ThisMon
                       AND convert(varchar, Send_amount_date, 112) <= @Base_date THEN 1
                  ELSE 0
              END) AS month_incase_num_FDCOM003_1 /* 和潤 日進件數*/ ,
          sum(CASE
                  WHEN 'FDCOM004'=fund_company
                       AND convert(varchar, Send_amount_date, 112) = @Base_date THEN 1
                  ELSE 0
              END) AS day_incase_num_FDCOM004 /* 和潤 月進件數*/ ,
          sum(CASE
                  WHEN 'FDCOM004'=fund_company
                       AND left(convert(varchar, Send_amount_date, 112), 6) = @ThisMon
                       AND convert(varchar, Send_amount_date, 112) <= @Base_date THEN 1
                  ELSE 0
              END) AS month_incase_num_FDCOM004 /* 福斯 日進件數*/ ,
          sum(CASE
                  WHEN 'FDCOM005'=fund_company
                       AND convert(varchar, Send_amount_date, 112) = @Base_date THEN 1
                  ELSE 0
              END) AS day_incase_num_FDCOM005 /* 福斯 月進件數*/ ,
          sum(CASE
                  WHEN 'FDCOM005'=fund_company
                       AND left(convert(varchar, Send_amount_date, 112), 6) = @ThisMon
                       AND convert(varchar, Send_amount_date, 112) <= @Base_date THEN 1
                  ELSE 0
              END) AS month_incase_num_FDCOM005 /* 日撥款數*/ ,
          sum(CASE
                  WHEN project_title NOT IN ('PJ00046', 'PJ00047', 'PJ00048')
                       AND convert(varchar,get_amount_date, 112) = @Base_date THEN 1
                  ELSE 0
              END) AS day_get_amount_num /* 日撥款額*/ ,
          sum(CASE
                  WHEN project_title NOT IN ('PJ00046', 'PJ00047', 'PJ00048')
                       AND convert(varchar,get_amount_date, 112) = @Base_date THEN get_amount
                  ELSE 0
              END) AS day_get_amount /* 月核准數*/ ,
          sum(CASE
                  WHEN project_title NOT IN ('PJ00046', 'PJ00047', 'PJ00048')
                       AND left(convert(varchar, Send_result_date, 112), 6)=@ThisMon
                       AND convert(varchar, Send_result_date, 112) <= @Base_date
                       AND Send_result_type IN ('SRT002', 'SRT005') THEN 1
                  ELSE 0
              END) AS month_pass_num /* 月撥款數*/ ,
          sum(CASE
                  WHEN project_title NOT IN ('PJ00046', 'PJ00047', 'PJ00048')
                       AND left(convert(varchar, get_amount_date, 112), 6) = @ThisMon
                       AND convert(varchar,get_amount_date, 112) <= @Base_date
                       AND get_amount_type IN ('GTAT002') THEN 1
                  ELSE 0
              END) AS month_get_amount_num /* 新鑫 月撥款額*/ ,
          sum(CASE
                  WHEN 'FDCOM001'=fund_company
                       AND left(convert(varchar, get_amount_date, 112), 6) = @ThisMon
                       AND convert(varchar,get_amount_date, 112) <= @Base_date THEN get_amount
                  ELSE 0
              END) AS month_get_amount_FDCOM001 /* 國&#23791; 月撥款額*/ ,
          sum(CASE
                  WHEN 'FDCOM003'=fund_company
                       AND project_title NOT IN ('PJ00046', 'PJ00047', 'PJ00048')
                       AND left(convert(varchar, get_amount_date, 112), 6) = @ThisMon
                       AND convert(varchar,get_amount_date, 112) <= @Base_date THEN get_amount
                  ELSE 0
              END) AS month_get_amount_FDCOM003 /*專案PJ00098、PJ00099-薪分期、股好貸  月撥款額 */ ,
          sum(CASE
                  WHEN 'FDCOM003_1'=fund_company
                       AND left(convert(varchar, get_amount_date, 112), 6) = @ThisMon
                       AND convert(varchar,get_amount_date, 112) <= @Base_date THEN get_amount
                  ELSE 0
              END) AS month_get_amount_FDCOM003_1 /* 和潤 月撥款額*/ ,
          sum(CASE
                  WHEN 'FDCOM004'=fund_company
                       AND left(convert(varchar, get_amount_date, 112), 6) = @ThisMon
                       AND convert(varchar,get_amount_date, 112) <= @Base_date THEN get_amount
                  ELSE 0
              END) AS month_get_amount_FDCOM004 /* 福斯 月撥款額*/ ,
          sum(CASE
                  WHEN 'FDCOM005'=fund_company
                       AND left(convert(varchar, get_amount_date, 112), 6) = @ThisMon
                       AND convert(varchar,get_amount_date, 112) <= @Base_date THEN get_amount
                  ELSE 0
              END) AS month_get_amount_FDCOM005 /* 新鑫 已核未撥*/ ,
          sum(CASE
                  WHEN 'FDCOM001'=fund_company
                       AND left(convert(varchar, Send_amount_date, 112), 6) IN (@ThisMon, @PreMon, @PreMon1)
                       AND convert(varchar,Send_result_date, 112) <= @Base_date
                       AND Send_result_type = 'SRT002'
                       AND isnull(check_amount_type, '') NOT IN ('CKAT003')
                       AND isnull(get_amount_type, '') NOT IN ('GTAT002', 'GTAT003') THEN pass_amount
                  ELSE 0
              END) AS month_pass_amount_FDCOM001 /* 新鑫 預核額度*/ ,
          sum(CASE
                  WHEN 'FDCOM001'=fund_company
                       AND left(convert(varchar, Send_amount_date, 112), 6) IN (@ThisMon, @PreMon)
                       AND convert(varchar,Send_result_date, 112) <= @Base_date
                       AND Send_result_type = 'SRT005' THEN pass_amount
                  ELSE 0
              END) AS month_pre_amount_FDCOM001 /* 國&#23791; 已核未撥*/ ,
          sum(CASE
                  WHEN 'FDCOM003'=fund_company
                       AND project_title NOT IN ('PJ00046', 'PJ00047', 'PJ00048')
                       AND left(convert(varchar, Send_amount_date, 112), 6) IN (@ThisMon, @PreMon, @PreMon1)
                       AND convert(varchar,Send_result_date, 112) <= @Base_date
                       AND Send_result_type = 'SRT002'
                       AND isnull(check_amount_type, '') NOT IN ('CKAT003')
                       AND isnull(get_amount_type, '') NOT IN ('GTAT002', 'GTAT003') THEN pass_amount
                  ELSE 0
              END) AS month_pass_amount_FDCOM003 /*專案PJ00098、PJ00099-薪分期、股好貸  已核未撥*/ ,
          sum(CASE
                  WHEN 'FDCOM003_1'=fund_company
                       AND left(convert(varchar, Send_amount_date, 112), 6) IN (@ThisMon, @PreMon, @PreMon1)
                       AND convert(varchar,Send_result_date, 112) <= @Base_date
                       AND Send_result_type = 'SRT002'
                       AND isnull(check_amount_type, '') NOT IN ('CKAT003')
                       AND isnull(get_amount_type, '') NOT IN ('GTAT002', 'GTAT003') THEN pass_amount
                  ELSE 0
              END) AS month_pass_amount_FDCOM003_1 /* 和潤 已核未撥*/ ,
          sum(CASE
                  WHEN 'FDCOM004'=fund_company
                       AND left(convert(varchar, Send_amount_date, 112), 6) IN (@ThisMon, @PreMon)
                       AND convert(varchar,Send_result_date, 112) <= @Base_date
                       AND Send_result_type = 'SRT002'
                       AND isnull(check_amount_type, '') NOT IN ('CKAT003')
                       AND isnull(get_amount_type, '') NOT IN ('GTAT002', 'GTAT003') THEN pass_amount
                  ELSE 0
              END) AS month_pass_amount_FDCOM004 /* 福斯 已核未撥*/ ,
          sum(CASE
                  WHEN 'FDCOM005'=fund_company
                       AND left(convert(varchar, Send_amount_date, 112), 6) IN (@ThisMon, @PreMon)
                       AND convert(varchar,Send_result_date, 112) <= @Base_date
                       AND Send_result_type = 'SRT002'
                       AND isnull(check_amount_type, '') NOT IN ('CKAT003')
                       AND isnull(get_amount_type, '') NOT IN ('GTAT002', 'GTAT003') THEN pass_amount
                  ELSE 0
              END) AS month_pass_amount_FDCOM005 /* 國&#23791; 代墊款(萬)*/ ,
          sum(CASE
                  WHEN 'FDCOM003'=fund_company
                       AND project_title NOT IN ('PJ00046', 'PJ00047', 'PJ00048')
                       AND left(convert(varchar, get_amount_date, 112), 6) = @ThisMon
                       AND convert(varchar,get_amount_date, 112) <= @Base_date THEN advance_payment_AE
                  ELSE 0
              END) AS advance_payment_AE /* viewFeats 沒有的欄位 */
       FROM view_User_group ug
       JOIN view_User_sales leader ON leader.U_num = ug.group_M_code
       JOIN view_User_sales sa ON sa.U_num = ug.group_D_code
       LEFT JOIN viewFeats f ON ug.group_D_code = f.plan_num
       AND ( left(convert(varchar, f.Send_amount_date, 112), 6)  = @ThisMon
            OR left(convert(varchar, f.Send_amount_date, 112), 6) IN (@ThisMon,@PreMon,@PreMon1)
            OR left(convert(varchar, f.get_amount_date, 112), 6) = @ThisMon)
       WHERE @Base_date BETWEEN ug.group_M_start_day AND ug.group_M_end_day
         AND @Base_date BETWEEN ug.group_D_start_day AND ug.group_D_end_day
       GROUP BY leader.bc_sort,leader.U_BC,sa.U_susp_date,sa.is_susp,
                isnull(ug.group_M_name, '未分組'),ug.group_D_name,ug.group_D_code,
                ug.group_id,ug.group_M_id,ug.group_M_title,sa.U_PFT_sort,
                sa.U_PFT_name) /* ===== cte A ===== */
    SELECT isnull(SG.Spec_Group, U_BC) Dis_BC,BC_Name,PE.PE_target,a.*,isnull(ft.target_quota, 0) target_quota,
    case when 
     ROUND(dbo.DivToFloat(isnull(month_get_amount_FDCOM001,0)+isnull(month_get_amount_FDCOM003,0),target_quota),2)*100 is null then
     '--'
     else
     CAST(CAST( ROUND(dbo.DivToFloat(isnull(month_get_amount_FDCOM001,0)+isnull(month_get_amount_FDCOM003,0),target_quota),4)*100 as decimal(12,2))  as varchar)+'%'
     end target_perc
    FROM A LEFT JOIN (SELECT * FROM Person_target WHERE PE_Date=@PE_Date) PE ON A.plan_num=PE.PE_num
    LEFT JOIN User_Spec_Group SG ON A.plan_num=SG.U_num
    LEFT JOIN Feat_target ft ON ft.del_tag='0'
    AND ft.U_num=A.plan_num AND ft.group_id=A.group_id
    AND ft.target_ym=@ThisMon
    left join (
    select item_D_code,item_D_name BC_Name from Item_list where item_M_code = 'Spec_Group' and item_M_type='N'
    union all
    select  item_D_code,item_D_name BC_Name from Item_list  where item_M_code = 'branch_company' and item_M_type='N'
    ) BC on isnull(SG.Spec_Group, U_BC)=BC.item_D_code
    WHERE  (isnull(is_susp, 'N')='N'
           OR  convert(varchar, U_susp_date, 112) >=@Base_date)";
            if (U_BC != "")
            {
                T_SQL += @" and U_BC=@U_BC ";
            }
            


            T_SQL += @"ORDER BY bc_sort,isnull(SG.Spec_Group, A.U_BC),
             A.leader_name,
             CASE
                 WHEN plan_num='K9999'THEN 999
                 ELSE U_PFT_sort
             END,A.U_PFT_name,A.plan_num";
            parameters.Add(new SqlParameter("@Base_date", Base_Date));
            parameters.Add(new SqlParameter("@ThisMon", ThisMon));
            parameters.Add(new SqlParameter("@PreMon", PreMon));
            parameters.Add(new SqlParameter("@PreMon1", PreMon1));
            parameters.Add(new SqlParameter("@PE_Date", PE_Date));
            if (U_BC != "")
            {
                parameters.Add(new SqlParameter("@U_BC", U_BC));
            }
           #endregion
           DataTable dtResult = _adoData.ExecuteQuery(T_SQL, parameters,true);

            //總計
            DataTable dtTotle = new DataTable();
            dtTotle.Columns.Add("SEQ");//職位
            dtTotle.Columns.Add("U_PFT_name");//職位
            dtTotle.Columns.Add("plan_name");//業務
            dtTotle.Columns.Add("day_incase_num_FDCOM001");//新鑫日進件數
            dtTotle.Columns.Add("month_incase_num_FDCOM001");//新鑫累積進件
            dtTotle.Columns.Add("day_incase_num_FDCOM003");//國峯日進件數
            dtTotle.Columns.Add("month_incase_num_FDCOM003");//國峯累積進件
            dtTotle.Columns.Add("day_get_amount_num");//日撥件數
            dtTotle.Columns.Add("day_get_amount");//日撥金額    
            dtTotle.Columns.Add("month_pass_num");//核准件數    
            dtTotle.Columns.Add("month_get_amount_num");//累積撥款件數    
            dtTotle.Columns.Add("month_get_amount_FDCOM001");//新鑫-撥款金額(萬)    
            dtTotle.Columns.Add("month_get_amount_FDCOM003");//國峯-撥款金額(萬)    
            dtTotle.Columns.Add("month_pass_amount_FDCOM001");//新鑫已核未撥
            dtTotle.Columns.Add("month_pass_amount_FDCOM003");//國峯已核未撥
            dtTotle.Columns.Add("target_quota");//目標    
            dtTotle.Columns.Add("target_perc");//達成率
            dtTotle.Columns.Add("PE_target");//責任額 

            //總計_T
            DataTable dtTotle_T = new DataTable();
            //機車
            DataTable dtEngine = new DataTable();

            dtEngine.Columns.Add("SEQ");//職位
            dtEngine.Columns.Add("U_PFT_name");//職位
            dtEngine.Columns.Add("plan_name");//業務
            dtEngine.Columns.Add("day_incase_num_PJ00046");//機車貸A日進件數
            dtEngine.Columns.Add("month_incase_num_PJ00046");//機車貸A累積進件
            dtEngine.Columns.Add("day_incase_num_PJ00047");//機車貸B日進件數
            dtEngine.Columns.Add("month_incase_num_PJ00047");//機車貸B累積進件
            dtEngine.Columns.Add("day_get_amount_num_engine");//日撥件數
            dtEngine.Columns.Add("day_get_amount_engine");//日撥金額(萬) 
            dtEngine.Columns.Add("month_pass_num_engine");//核准件數
            dtEngine.Columns.Add("month_get_amount_num_engine");//累積撥款件數
            dtEngine.Columns.Add("month_get_amount_PJ00046");//機車貸A撥款金額(萬)
            dtEngine.Columns.Add("month_get_amount_PJ00047");//機車貸B撥款金額(萬)
            dtEngine.Columns.Add("month_pass_amount_PJ00046");//機車貸款A; 已核未撥 
            dtEngine.Columns.Add("month_pass_amount_PJ00047");//機車貸款B; 已核未撥 
            dtEngine.Columns.Add("sum_amount");//累計業績

            //機車_T
            DataTable dtEngine_T = new DataTable();
            //房貸前一天
            DataTable dtPreTotle_T = new DataTable();
            //機車前一天
            DataTable dtPreEngine_T = new DataTable();

            //汽車
            DataTable dtCar = new DataTable();
            dtCar.Columns.Add("SEQ");//職位
            dtCar.Columns.Add("U_PFT_name");//職位
            dtCar.Columns.Add("plan_name");//業務
            dtCar.Columns.Add("day_incase_num_PJ00048");//汽車-日進件數
            dtCar.Columns.Add("month_incase_num_PJ00048");// 汽車-月進件數
            dtCar.Columns.Add("day_get_amount_num_Car");//汽車-日撥款數
            dtCar.Columns.Add("day_get_amount_Car");//汽車-日撥款額
            dtCar.Columns.Add("month_pass_num_Car");//汽車-月核准數
            dtCar.Columns.Add("month_get_amount_num_Car");//汽車-月撥款數
            dtCar.Columns.Add("month_get_amount_PJ00048");//汽車-月撥款額
            dtCar.Columns.Add("month_pass_amount_PJ00048");//汽車-已核未撥
            //汽車_T
            DataTable dtCar_T = new DataTable();
            //汽車_T前一天
            DataTable dtPreCar_T = new DataTable();


            /*房貸欄位*/
            string U_PFT_name, plan_num, BC_Name, plan_name, day_incase_num_FDCOM001, month_incase_num_FDCOM001, day_incase_num_FDCOM003, month_incase_num_FDCOM003;
            string day_get_amount_num, day_get_amount, month_pass_num, month_get_amount_num, month_get_amount_FDCOM001, month_get_amount_FDCOM003, month_pass_amount_FDCOM001;
            string month_pass_amount_FDCOM003, PE_target, target_perc, target_quota;
            Int32 iday_incase_num_FDCOM001 = 0, imonth_incase_num_FDCOM001 = 0, iday_incase_num_FDCOM003 = 0, imonth_incase_num_FDCOM003 = 0, iday_get_amount_num = 0, iday_get_amount = 0, imonth_pass_num = 0;
            Int32 imonth_get_amount_num = 0, imonth_get_amount_FDCOM001 = 0, imonth_get_amount_FDCOM003 = 0, imonth_pass_amount_FDCOM001 = 0, imonth_pass_amount_FDCOM003 = 0, iPE_target = 0, itarget_quota = 0;
            Int32 iday_incase_num_FDCOM001_T = 0, imonth_incase_num_FDCOM001_T = 0, iday_incase_num_FDCOM003_T = 0, imonth_incase_num_FDCOM003_T = 0, iday_get_amount_num_T = 0, iday_get_amount_T = 0, imonth_pass_num_T = 0;
            Int32 imonth_get_amount_num_T = 0, imonth_get_amount_FDCOM001_T = 0, imonth_get_amount_FDCOM003_T = 0, imonth_pass_amount_FDCOM001_T = 0, imonth_pass_amount_FDCOM003_T = 0, iPE_target_T = 0, itarget_quota_T = 0;
            Int32 BC_Count = 0;
            /*機車貸欄位*/
            string day_incase_num_PJ00046, month_incase_num_PJ00046, day_incase_num_PJ00047, month_incase_num_PJ00047, day_get_amount_num_engine, day_get_amount_engine;
            string month_pass_num_engine, month_get_amount_num_engine, month_get_amount_PJ00046, month_get_amount_PJ00047, month_pass_amount_PJ00046, month_pass_amount_PJ00047, sum_amount;
            Int32 iday_incase_num_PJ00046 = 0, imonth_incase_num_PJ00046 = 0, iday_incase_num_PJ00047 = 0, imonth_incase_num_PJ00047 = 0, iday_get_amount_num_engine = 0, iday_get_amount_engine = 0;
            Int32 iday_incase_num_PJ00046_T = 0, imonth_incase_num_PJ00046_T = 0, iday_incase_num_PJ00047_T = 0, imonth_incase_num_PJ00047_T = 0, iday_get_amount_num_engine_T = 0, iday_get_amount_engine_T = 0;

            Int32 imonth_pass_num_engine = 0, imonth_get_amount_num_engine = 0, imonth_get_amount_PJ00046 = 0, imonth_get_amount_PJ00047 = 0, imonth_pass_amount_PJ00046 = 0, imonth_pass_amount_PJ00047 = 0, isum_amount = 0;
            Int32 imonth_pass_num_engine_T = 0, imonth_get_amount_num_engine_T = 0, imonth_get_amount_PJ00046_T = 0, imonth_get_amount_PJ00047_T = 0, imonth_pass_amount_PJ00046_T = 0, imonth_pass_amount_PJ00047_T = 0, isum_amount_T = 0;

            /*汽車貸欄位*/
            string day_incase_num_PJ00048, month_incase_num_PJ00048, day_get_amount_num_Car, day_get_amount_Car, month_pass_num_Car, month_get_amount_num_Car, month_get_amount_PJ00048, month_pass_amount_PJ00048;
            Int32 iday_incase_num_PJ00048 = 0, imonth_incase_num_PJ00048 = 0, iday_get_amount_num_Car = 0, iday_get_amount_Car = 0, imonth_pass_num_Car = 0, imonth_get_amount_num_Car = 0, imonth_get_amount_PJ00048 = 0, imonth_pass_amount_PJ00048 = 0;

            string U_PFT_sort = "";
            Int32 m_RowIdx = 0, BC0900Count = 0, iSEQ = 1, iCar_SEQ = 1; 
            foreach (DataRow dr in dtResult.Rows)
            {

                string Dis_BC = dr["Dis_BC"].ToString();
                U_PFT_sort = dr["U_PFT_sort"].ToString();
                U_PFT_name = dr["U_PFT_name"].ToString();
                plan_name = dr["plan_name"].ToString();
                if (plan_name.IndexOf("許彥") != -1)
                {
                    plan_name = "許彥珉";
                }
                plan_num = dr["plan_num"].ToString();
                BC_Name = dr["BC_Name"].ToString();


                /*房貸欄位*/
                day_incase_num_FDCOM001 = dr["day_incase_num_FDCOM001"].ToString();
                iday_incase_num_FDCOM001 += Convert.ToInt32(dr["day_incase_num_FDCOM001"].ToString());
                iday_incase_num_FDCOM001_T += Convert.ToInt32(dr["day_incase_num_FDCOM001"].ToString());

                month_incase_num_FDCOM001 = dr["month_incase_num_FDCOM001"].ToString();
                imonth_incase_num_FDCOM001 += Convert.ToInt32(dr["month_incase_num_FDCOM001"].ToString());
                imonth_incase_num_FDCOM001_T += Convert.ToInt32(dr["month_incase_num_FDCOM001"].ToString());

                day_incase_num_FDCOM003 = dr["day_incase_num_FDCOM003"].ToString();
                iday_incase_num_FDCOM003 += Convert.ToInt32(dr["day_incase_num_FDCOM003"].ToString());
                iday_incase_num_FDCOM003_T += Convert.ToInt32(dr["day_incase_num_FDCOM003"].ToString());

                month_incase_num_FDCOM003 = dr["month_incase_num_FDCOM003"].ToString();
                imonth_incase_num_FDCOM003 += Convert.ToInt32(dr["month_incase_num_FDCOM003"].ToString());
                imonth_incase_num_FDCOM003_T += Convert.ToInt32(dr["month_incase_num_FDCOM003"].ToString());

                day_get_amount_num = dr["day_get_amount_num"].ToString();
                iday_get_amount_num += Convert.ToInt32(dr["day_get_amount_num"].ToString());
                iday_get_amount_num_T += Convert.ToInt32(dr["day_get_amount_num"].ToString());

                day_get_amount = dr["day_get_amount"].ToString();
                iday_get_amount += Convert.ToInt32(dr["day_get_amount"].ToString());
                iday_get_amount_T += Convert.ToInt32(dr["day_get_amount"].ToString());

                month_pass_num = dr["month_pass_num"].ToString();
                imonth_pass_num += Convert.ToInt32(dr["month_pass_num"].ToString());
                imonth_pass_num_T += Convert.ToInt32(dr["month_pass_num"].ToString());

                month_get_amount_num = dr["month_get_amount_num"].ToString();
                imonth_get_amount_num += Convert.ToInt32(dr["month_get_amount_num"].ToString());
                imonth_get_amount_num_T += Convert.ToInt32(dr["month_get_amount_num"].ToString());

                month_get_amount_FDCOM001 = dr["month_get_amount_FDCOM001"].ToString();
                imonth_get_amount_FDCOM001 += Convert.ToInt32(dr["month_get_amount_FDCOM001"].ToString());
                imonth_get_amount_FDCOM001_T += Convert.ToInt32(dr["month_get_amount_FDCOM001"].ToString());

                month_get_amount_FDCOM003 = dr["month_get_amount_FDCOM003"].ToString();
                imonth_get_amount_FDCOM003 += Convert.ToInt32(dr["month_get_amount_FDCOM003"].ToString());
                imonth_get_amount_FDCOM003_T += Convert.ToInt32(dr["month_get_amount_FDCOM003"].ToString());

                month_pass_amount_FDCOM001 = dr["month_pass_amount_FDCOM001"].ToString();
                imonth_pass_amount_FDCOM001 += Convert.ToInt32(dr["month_pass_amount_FDCOM001"].ToString());
                imonth_pass_amount_FDCOM001_T += Convert.ToInt32(dr["month_pass_amount_FDCOM001"].ToString());

                month_pass_amount_FDCOM003 = dr["month_pass_amount_FDCOM003"].ToString();
                imonth_pass_amount_FDCOM003 += Convert.ToInt32(dr["month_pass_amount_FDCOM003"].ToString());
                imonth_pass_amount_FDCOM003_T += Convert.ToInt32(dr["month_pass_amount_FDCOM003"].ToString());
                target_perc = "";
                /*機車貸欄位*/
                day_incase_num_PJ00046 = dr["day_incase_num_PJ00046"].ToString();
                iday_incase_num_PJ00046 += Convert.ToInt32(dr["day_incase_num_PJ00046"].ToString());
                iday_incase_num_PJ00046_T += Convert.ToInt32(dr["day_incase_num_PJ00046"].ToString());

                month_incase_num_PJ00046 = dr["month_incase_num_PJ00046"].ToString();
                imonth_incase_num_PJ00046 += Convert.ToInt32(dr["month_incase_num_PJ00046"].ToString());
                imonth_incase_num_PJ00046_T += Convert.ToInt32(dr["month_incase_num_PJ00046"].ToString());

                day_incase_num_PJ00047 = dr["day_incase_num_PJ00047"].ToString();
                iday_incase_num_PJ00047 += Convert.ToInt32(dr["day_incase_num_PJ00047"].ToString());
                iday_incase_num_PJ00047_T += Convert.ToInt32(dr["day_incase_num_PJ00047"].ToString());

                month_incase_num_PJ00047 = dr["month_incase_num_PJ00047"].ToString();
                imonth_incase_num_PJ00047 += Convert.ToInt32(dr["month_incase_num_PJ00047"].ToString());
                imonth_incase_num_PJ00047_T += Convert.ToInt32(dr["month_incase_num_PJ00047"].ToString());

                day_get_amount_num_engine = dr["day_get_amount_num_engine"].ToString();
                iday_get_amount_num_engine += Convert.ToInt32(dr["day_get_amount_num_engine"].ToString());
                iday_get_amount_num_engine_T += Convert.ToInt32(dr["day_get_amount_num_engine"].ToString());

                day_get_amount_engine = dr["day_get_amount_engine"].ToString();
                iday_get_amount_engine += Convert.ToInt32(dr["day_get_amount_engine"].ToString());
                iday_get_amount_engine_T += Convert.ToInt32(dr["day_get_amount_engine"].ToString());

                month_pass_num_engine = dr["month_pass_num_engine"].ToString();
                imonth_pass_num_engine += Convert.ToInt32(dr["month_pass_num_engine"].ToString());
                imonth_pass_num_engine_T += Convert.ToInt32(dr["month_pass_num_engine"].ToString());

                month_get_amount_num_engine = dr["month_get_amount_num_engine"].ToString();
                imonth_get_amount_num_engine += Convert.ToInt32(dr["month_get_amount_num_engine"].ToString());
                imonth_get_amount_num_engine_T += Convert.ToInt32(dr["month_get_amount_num_engine"].ToString());

                month_get_amount_PJ00046 = dr["month_get_amount_PJ00046"].ToString();
                imonth_get_amount_PJ00046 += Convert.ToInt32(dr["month_get_amount_PJ00046"].ToString());
                imonth_get_amount_PJ00046_T += Convert.ToInt32(dr["month_get_amount_PJ00046"].ToString());

                month_get_amount_PJ00047 = dr["month_get_amount_PJ00047"].ToString();
                imonth_get_amount_PJ00047 += Convert.ToInt32(dr["month_get_amount_PJ00047"].ToString());
                imonth_get_amount_PJ00047_T += Convert.ToInt32(dr["month_get_amount_PJ00047"].ToString());

                month_pass_amount_PJ00046 = dr["month_pass_amount_PJ00046"].ToString();
                imonth_pass_amount_PJ00046 += Convert.ToInt32(dr["month_pass_amount_PJ00046"].ToString());
                imonth_pass_amount_PJ00046_T += Convert.ToInt32(dr["month_pass_amount_PJ00046"].ToString());

                month_pass_amount_PJ00047 = dr["month_pass_amount_PJ00047"].ToString();
                imonth_pass_amount_PJ00047 += Convert.ToInt32(dr["month_pass_amount_PJ00047"].ToString());
                imonth_pass_amount_PJ00047_T += Convert.ToInt32(dr["month_pass_amount_PJ00047"].ToString());

                sum_amount = (Convert.ToInt32(month_get_amount_PJ00046) + Convert.ToInt32(month_get_amount_PJ00047)).ToString();
                isum_amount += Convert.ToInt32(sum_amount);
                isum_amount_T += Convert.ToInt32(sum_amount);

                if (Dis_BC != "BC0900" && Dis_BC != "BC0901")
                {
                    PE_target = dr["PE_target"].ToString();
                    if (PE_target != "")
                    {
                        iPE_target += Convert.ToInt32(dr["PE_target"].ToString());
                        iPE_target_T += Convert.ToInt32(dr["PE_target"].ToString());
                    }
                    

                    target_perc = dr["target_perc"].ToString();
                    target_quota = dr["target_quota"].ToString();
                    if (target_quota != "")
                    {
                        itarget_quota += Convert.ToInt32(dr["target_quota"].ToString());
                        itarget_quota_T += Convert.ToInt32(dr["target_quota"].ToString());
                    }
                    
                }
                else
                {
                    PE_target = "--";
                    target_perc = "--";
                    target_quota = "--";
                }

                
                    if (Dis_BC != "BC0900" && Dis_BC != "BC0901")
                    {
                        if ((U_PFT_sort == "120" || U_PFT_sort == "130"))
                        {
                            DataRow TotRow = dtTotle.NewRow();
                            TotRow["SEQ"] = "";
                            TotRow["U_PFT_name"] = "";
                            TotRow["plan_name"] = arrTitle[BC_Count];
                            TotRow["month_get_amount_FDCOM001"] = AddDate;
                            dtTotle.Rows.Add(TotRow);
                            dtTotle.Rows.Add(FuncHandler.AddTitleByTable(dtTotle, "Totle"));

                            DataRow TotRow_E = dtEngine.NewRow();
                            TotRow_E["SEQ"] = "";
                            TotRow_E["U_PFT_name"] = "";
                            TotRow_E["plan_name"] = arrTitle[BC_Count];
                            TotRow_E["month_get_amount_PJ00046"] = AddDate;
                            dtEngine.Rows.Add(TotRow_E);
                            dtEngine.Rows.Add(FuncHandler.AddTitleByTable(dtEngine, "Engine"));
                            BC_Count++;
                        }
                    }
                    else
                    {
                        if (Dis_BC == "BC0900")
                        {
                            if (BC0900Count == 0)
                            {
                                DataRow TotRow = dtTotle.NewRow();
                                TotRow["SEQ"] = "";
                                TotRow["U_PFT_name"] = "";
                                TotRow["plan_name"] = arrTitle[BC_Count];
                                TotRow["month_get_amount_FDCOM001"] = AddDate;
                                dtTotle.Rows.Add(TotRow);
                                dtTotle.Rows.Add(FuncHandler.AddTitleByTable(dtTotle, "Totle"));

                                DataRow TotRow_E = dtEngine.NewRow();
                                TotRow_E["SEQ"] = "";
                                TotRow_E["U_PFT_name"] = "";
                                TotRow_E["plan_name"] = arrTitle[BC_Count];
                                TotRow_E["month_get_amount_PJ00046"] = AddDate;
                                dtEngine.Rows.Add(TotRow_E);
                                dtEngine.Rows.Add(FuncHandler.AddTitleByTable(dtEngine, "Engine"));
                                BC_Count++;
                            }
                            BC0900Count++;
                        }
                        else
                        {
                            DataRow TotRow = dtTotle.NewRow();
                            TotRow["SEQ"] = "";
                            TotRow["U_PFT_name"] = "";
                            TotRow["plan_name"] = arrTitle[BC_Count];
                            TotRow["month_get_amount_FDCOM001"] = AddDate;
                            dtTotle.Rows.Add(TotRow);
                            dtTotle.Rows.Add(FuncHandler.AddTitleByTable(dtTotle, "Totle"));

                            DataRow TotRow_E = dtEngine.NewRow();
                            TotRow_E["SEQ"] = "";
                            TotRow_E["U_PFT_name"] = "";
                            TotRow_E["plan_name"] = arrTitle[BC_Count];
                            TotRow_E["month_get_amount_PJ00046"] = AddDate;
                            dtEngine.Rows.Add(TotRow_E);
                            dtEngine.Rows.Add(FuncHandler.AddTitleByTable(dtEngine, "Engine"));
                            BC_Count++;
                        }
                    }
               
                DataRow DaRow = dtTotle.NewRow();
                DaRow = dtTotle.NewRow();
                DaRow["U_PFT_name"] = U_PFT_name;
                DaRow["plan_name"] = plan_name;
                DaRow["day_incase_num_FDCOM001"] = day_incase_num_FDCOM001;
                DaRow["month_incase_num_FDCOM001"] = month_incase_num_FDCOM001;
                DaRow["day_incase_num_FDCOM003"] = day_incase_num_FDCOM003;
                DaRow["month_incase_num_FDCOM003"] = month_incase_num_FDCOM003;
                DaRow["day_get_amount_num"] = day_get_amount_num;
                DaRow["day_get_amount"] = day_get_amount;
                DaRow["month_pass_num"] = month_pass_num;
                DaRow["month_get_amount_num"] = month_get_amount_num;
                DaRow["month_get_amount_FDCOM001"] = month_get_amount_FDCOM001;
                DaRow["month_get_amount_FDCOM003"] = month_get_amount_FDCOM003;
                DaRow["month_pass_amount_FDCOM001"] = month_pass_amount_FDCOM001;
                DaRow["month_pass_amount_FDCOM003"] = month_pass_amount_FDCOM003;
                DaRow["SEQ"] = iSEQ.ToString();
                if (Dis_BC != "BC0900" && Dis_BC != "BC0901")
                {
                    DaRow["PE_target"] = PE_target;
                    DaRow["target_perc"] = target_perc;
                    DaRow["target_quota"] = target_quota;
                }
                dtTotle.Rows.Add(DaRow);


                DataRow DaRow_E = dtEngine.NewRow();
                DaRow_E = dtEngine.NewRow();
                DaRow_E["U_PFT_name"] = U_PFT_name;
                DaRow_E["plan_name"] = plan_name;
                DaRow_E["day_incase_num_PJ00046"] = day_incase_num_PJ00046;
                DaRow_E["month_incase_num_PJ00046"] = month_incase_num_PJ00046;
                DaRow_E["day_incase_num_PJ00047"] = day_incase_num_PJ00047;
                DaRow_E["month_incase_num_PJ00047"] = month_incase_num_PJ00047;
                DaRow_E["day_get_amount_num_engine"] = day_get_amount_num_engine;
                DaRow_E["day_get_amount_engine"] = day_get_amount_engine;
                DaRow_E["month_pass_num_engine"] = month_pass_num_engine;
                DaRow_E["month_get_amount_num_engine"] = month_get_amount_num_engine;
                DaRow_E["month_get_amount_PJ00046"] = month_get_amount_PJ00046;
                DaRow_E["month_get_amount_PJ00047"] = month_get_amount_PJ00047;
                DaRow_E["month_pass_amount_PJ00046"] = month_pass_amount_PJ00046;
                DaRow_E["month_pass_amount_PJ00047"] = month_pass_amount_PJ00047;
                DaRow_E["sum_amount"] = sum_amount;
                dtEngine.Rows.Add(DaRow_E);

                iSEQ++;



                if (dtResult.Rows.Count != m_RowIdx + 1)
                {
                    if (Dis_BC != dtResult.Rows[m_RowIdx + 1]["Dis_BC"].ToString())
                    {
                        DataRow TotRow = dtTotle.NewRow();
                        TotRow["SEQ"] = "";
                        TotRow["U_PFT_name"] = "";
                        TotRow["plan_name"] = "合計";
                        TotRow["day_incase_num_FDCOM001"] = iday_incase_num_FDCOM001;
                        TotRow["month_incase_num_FDCOM001"] = imonth_incase_num_FDCOM001;
                        TotRow["day_incase_num_FDCOM003"] = iday_incase_num_FDCOM003;
                        TotRow["month_incase_num_FDCOM003"] = imonth_incase_num_FDCOM003;
                        TotRow["day_get_amount_num"] = iday_get_amount_num;
                        TotRow["day_get_amount"] = iday_get_amount;
                        TotRow["month_pass_num"] = imonth_pass_num;
                        TotRow["month_get_amount_num"] = imonth_get_amount_num;
                        TotRow["month_get_amount_FDCOM001"] = imonth_get_amount_FDCOM001;
                        TotRow["month_get_amount_FDCOM003"] = imonth_get_amount_FDCOM003;
                        TotRow["month_pass_amount_FDCOM001"] = imonth_pass_amount_FDCOM001;
                        TotRow["month_pass_amount_FDCOM003"] = imonth_pass_amount_FDCOM003;


                        DataRow TotRow_E = dtEngine.NewRow();
                        TotRow_E["SEQ"] = "";
                        TotRow_E["U_PFT_name"] = "";
                        TotRow_E["plan_name"] = "合計";
                        TotRow_E["day_incase_num_PJ00046"] = iday_incase_num_PJ00046;
                        TotRow_E["month_incase_num_PJ00046"] = imonth_incase_num_PJ00046;
                        TotRow_E["day_incase_num_PJ00047"] = iday_incase_num_PJ00047;
                        TotRow_E["month_incase_num_PJ00047"] = imonth_incase_num_PJ00047;
                        TotRow_E["day_get_amount_num_engine"] = iday_get_amount_num_engine;
                        TotRow_E["day_get_amount_engine"] = iday_get_amount_engine;
                        TotRow_E["month_pass_num_engine"] = imonth_pass_num_engine;
                        TotRow_E["month_get_amount_num_engine"] = imonth_get_amount_num_engine;
                        TotRow_E["month_get_amount_PJ00046"] = imonth_get_amount_PJ00046;
                        TotRow_E["month_get_amount_PJ00047"] = imonth_get_amount_PJ00047;
                        TotRow_E["month_pass_amount_PJ00046"] = imonth_pass_amount_PJ00046;
                        TotRow_E["month_pass_amount_PJ00047"] = imonth_pass_amount_PJ00047;
                        TotRow_E["sum_amount"] = isum_amount;

                        if (Dis_BC != "BC0900" && Dis_BC != "BC0901")
                        {
                            TotRow["PE_target"] = iPE_target;
                            // 先轉成 double 做除法
                            double result = (double)(imonth_get_amount_FDCOM001 + imonth_get_amount_FDCOM003) / itarget_quota;
                            // 四捨五入到小數點後兩位
                            double rounded = Math.Round(result * 100, 2);
                            // 加上百分比字串
                            string percentage = rounded.ToString("F2") + "%";

                            TotRow["target_perc"] = percentage;
                            TotRow["target_quota"] = itarget_quota;
                        }
                        else
                        {
                            TotRow["PE_target"] = "--";
                            TotRow["target_perc"] = "--";
                            TotRow["target_quota"] = "--";
                        }
                        dtTotle.Rows.Add(TotRow);
                        DataRow EmpyRow = dtTotle.NewRow();
                        EmpyRow["SEQ"] = "";
                        EmpyRow["U_PFT_name"] = "";
                        EmpyRow["plan_name"] = "";
                        dtTotle.Rows.Add(EmpyRow);

                        dtEngine.Rows.Add(TotRow_E);
                        DataRow EmpyRow_E = dtEngine.NewRow();
                        EmpyRow_E["SEQ"] = "";
                        EmpyRow_E["U_PFT_name"] = "";
                        EmpyRow_E["plan_name"] = "";
                        dtEngine.Rows.Add(EmpyRow_E);

                        iday_incase_num_FDCOM001 = 0;
                        imonth_incase_num_FDCOM001 = 0;
                        iday_incase_num_FDCOM003 = 0;
                        imonth_incase_num_FDCOM003 = 0;
                        iday_get_amount_num = 0;
                        iday_get_amount = 0;
                        imonth_pass_num = 0;
                        imonth_get_amount_num = 0;
                        imonth_get_amount_FDCOM001 = 0;
                        imonth_get_amount_FDCOM003 = 0;
                        imonth_pass_amount_FDCOM001 = 0;
                        imonth_pass_amount_FDCOM003 = 0;
                        iPE_target = 0;
                        itarget_quota = 0;

                        iday_incase_num_PJ00046 = 0; imonth_incase_num_PJ00046 = 0; iday_incase_num_PJ00047 = 0; imonth_incase_num_PJ00047 = 0; iday_get_amount_num_engine = 0; iday_get_amount_engine = 0;
                        imonth_pass_num_engine = 0; imonth_get_amount_num_engine = 0; imonth_get_amount_PJ00046 = 0; imonth_get_amount_PJ00047 = 0; imonth_pass_amount_PJ00046 = 0; imonth_pass_amount_PJ00047 = 0; isum_amount = 0;

                    }
                }
                else
                {
                    DataRow TotRow = dtTotle.NewRow();
                    TotRow["SEQ"] = "";
                    TotRow["U_PFT_name"] = "";
                    TotRow["plan_name"] = "合計";
                    TotRow["day_incase_num_FDCOM001"] = iday_incase_num_FDCOM001;
                    TotRow["month_incase_num_FDCOM001"] = imonth_incase_num_FDCOM001;
                    TotRow["day_incase_num_FDCOM003"] = day_incase_num_FDCOM003;
                    TotRow["month_incase_num_FDCOM003"] = imonth_incase_num_FDCOM003;
                    TotRow["day_get_amount_num"] = iday_get_amount_num;
                    TotRow["day_get_amount"] = iday_get_amount;
                    TotRow["month_pass_num"] = imonth_pass_num;
                    TotRow["month_get_amount_num"] = imonth_get_amount_num;
                    TotRow["month_get_amount_FDCOM001"] = imonth_get_amount_FDCOM001;
                    TotRow["month_get_amount_FDCOM003"] = imonth_get_amount_FDCOM003;
                    TotRow["month_pass_amount_FDCOM001"] = imonth_pass_amount_FDCOM001;
                    TotRow["month_pass_amount_FDCOM003"] = imonth_pass_amount_FDCOM003;
                    if (Dis_BC != "BC0900" && Dis_BC != "BC0901")
                    {
                        TotRow["PE_target"] = iPE_target;
                        // 先轉成 double 做除法
                        double result0 = (double)(imonth_get_amount_FDCOM001 + imonth_get_amount_FDCOM003) / itarget_quota;
                        // 四捨五入到小數點後兩位
                        double rounded0 = Math.Round(result0 * 100, 2);
                        // 加上百分比字串
                        string percentage0 = rounded0.ToString("F2") + "%";
                        TotRow["target_perc"] = percentage0;
                        TotRow["target_quota"] = itarget_quota;
                    }
                    else
                    {
                        TotRow["PE_target"] = "--";
                        TotRow["target_perc"] = "--";
                        TotRow["target_quota"] = "--";
                    }
                    dtTotle.Rows.Add(TotRow);
                    DataRow EmpyRow = dtTotle.NewRow();
                    EmpyRow["SEQ"] = "";
                    EmpyRow["U_PFT_name"] = "";
                    EmpyRow["plan_name"] = "";
                    dtTotle.Rows.Add(EmpyRow);

                    iday_incase_num_FDCOM001 = 0;
                    imonth_incase_num_FDCOM001 = 0;
                    iday_incase_num_FDCOM003 = 0;
                    imonth_incase_num_FDCOM003 = 0;
                    iday_get_amount_num = 0;
                    iday_get_amount = 0;
                    imonth_pass_num = 0;
                    imonth_get_amount_num = 0;
                    imonth_get_amount_FDCOM001 = 0;
                    imonth_get_amount_FDCOM003 = 0;
                    imonth_pass_amount_FDCOM001 = 0;
                    imonth_pass_amount_FDCOM003 = 0;
                    iPE_target = 0;
                    itarget_quota = 0;


                    DataRow TotRow_T = dtTotle.NewRow();
                    TotRow_T["SEQ"] = "";
                    TotRow_T["U_PFT_name"] = "";
                    TotRow_T["plan_name"] = "總計";
                    TotRow_T["day_incase_num_FDCOM001"] = iday_incase_num_FDCOM001_T;
                    TotRow_T["month_incase_num_FDCOM001"] = imonth_incase_num_FDCOM001_T;
                    TotRow_T["day_incase_num_FDCOM003"] = iday_incase_num_FDCOM003_T;
                    TotRow_T["month_incase_num_FDCOM003"] = imonth_incase_num_FDCOM003_T;
                    TotRow_T["day_get_amount_num"] = iday_get_amount_num_T;
                    TotRow_T["day_get_amount"] = iday_get_amount_T;
                    TotRow_T["month_pass_num"] = imonth_pass_num_T;
                    TotRow_T["month_get_amount_num"] = imonth_get_amount_num_T;
                    TotRow_T["month_get_amount_FDCOM001"] = imonth_get_amount_FDCOM001_T;
                    TotRow_T["month_get_amount_FDCOM003"] = imonth_get_amount_FDCOM003_T;
                    TotRow_T["month_pass_amount_FDCOM001"] = imonth_pass_amount_FDCOM001_T;
                    TotRow_T["month_pass_amount_FDCOM003"] = imonth_pass_amount_FDCOM003_T;

                    // 先轉成 double 做除法
                    double result = (double)(imonth_get_amount_FDCOM001_T + imonth_get_amount_FDCOM003_T) / itarget_quota_T;
                    // 四捨五入到小數點後兩位
                    double rounded = Math.Round(result * 100, 2);
                    // 加上百分比字串
                    string percentage = rounded.ToString("F2") + "%";

                    TotRow_T["PE_target"] = iPE_target_T;
                    TotRow_T["target_perc"] = percentage;
                    TotRow_T["target_quota"] = itarget_quota_T;
                    dtTotle.Rows.Add(TotRow_T);


                    DataRow TotRow_E = dtEngine.NewRow();
                    TotRow_E["SEQ"] = "";
                    TotRow_E["U_PFT_name"] = "";
                    TotRow_E["plan_name"] = "合計";
                    TotRow_E["day_incase_num_PJ00046"] = iday_incase_num_PJ00046;
                    TotRow_E["month_incase_num_PJ00046"] = imonth_incase_num_PJ00046;
                    TotRow_E["day_incase_num_PJ00047"] = iday_incase_num_PJ00047;
                    TotRow_E["month_incase_num_PJ00047"] = imonth_incase_num_PJ00047;
                    TotRow_E["day_get_amount_num_engine"] = iday_get_amount_num_engine;
                    TotRow_E["day_get_amount_engine"] = iday_get_amount_engine;
                    TotRow_E["month_pass_num_engine"] = imonth_pass_num_engine;
                    TotRow_E["month_get_amount_num_engine"] = imonth_get_amount_num_engine;
                    TotRow_E["month_get_amount_PJ00046"] = imonth_get_amount_PJ00046;
                    TotRow_E["month_get_amount_PJ00047"] = imonth_get_amount_PJ00047;
                    TotRow_E["month_pass_amount_PJ00046"] = imonth_pass_amount_PJ00046;
                    TotRow_E["month_pass_amount_PJ00047"] = imonth_pass_amount_PJ00047;
                    TotRow_E["sum_amount"] = isum_amount;

                    dtEngine.Rows.Add(TotRow_E);
                    DataRow EmpyRow_E = dtEngine.NewRow();
                    EmpyRow_E["SEQ"] = "";
                    EmpyRow_E["U_PFT_name"] = "";
                    EmpyRow_E["plan_name"] = "";
                    dtEngine.Rows.Add(EmpyRow_E);

                    DataRow TotRow_E_T = dtEngine.NewRow();
                    TotRow_E_T["SEQ"] = "";
                    TotRow_E_T["U_PFT_name"] = "";
                    TotRow_E_T["plan_name"] = "總計";
                    TotRow_E_T["day_incase_num_PJ00046"] = iday_incase_num_PJ00046_T;
                    TotRow_E_T["month_incase_num_PJ00046"] = imonth_incase_num_PJ00046_T;
                    TotRow_E_T["day_incase_num_PJ00047"] = iday_incase_num_PJ00047_T;
                    TotRow_E_T["month_incase_num_PJ00047"] = imonth_incase_num_PJ00047_T;
                    TotRow_E_T["day_get_amount_num_engine"] = iday_get_amount_num_engine_T;
                    TotRow_E_T["day_get_amount_engine"] = iday_get_amount_engine_T;
                    TotRow_E_T["month_pass_num_engine"] = imonth_pass_num_engine_T;
                    TotRow_E_T["month_get_amount_num_engine"] = imonth_get_amount_num_engine_T;
                    TotRow_E_T["month_get_amount_PJ00046"] = imonth_get_amount_PJ00046_T;
                    TotRow_E_T["month_get_amount_PJ00047"] = imonth_get_amount_PJ00047_T;
                    TotRow_E_T["month_pass_amount_PJ00046"] = imonth_pass_amount_PJ00046_T;
                    TotRow_E_T["month_pass_amount_PJ00047"] = imonth_pass_amount_PJ00047_T;
                    TotRow_E_T["sum_amount"] = isum_amount_T;
                    dtEngine.Rows.Add(TotRow_E_T);
                }

                //符合汽車再新增
                if (arrCar.Contains(plan_num))
                {
                    /*汽車貸欄位*/
                    day_incase_num_PJ00048 = dr["day_incase_num_PJ00048"].ToString();
                    iday_incase_num_PJ00048 += Convert.ToInt32(dr["day_incase_num_PJ00048"].ToString());

                    month_incase_num_PJ00048 = dr["month_incase_num_PJ00048"].ToString();
                    imonth_incase_num_PJ00048 += Convert.ToInt32(dr["month_incase_num_PJ00048"].ToString());

                    day_get_amount_num_Car = dr["day_get_amount_num_Car"].ToString();
                    iday_get_amount_num_Car += Convert.ToInt32(dr["day_get_amount_num_Car"].ToString());

                    day_get_amount_Car = dr["day_get_amount_Car"].ToString();
                    iday_get_amount_Car += Convert.ToInt32(dr["day_get_amount_Car"].ToString());

                    month_pass_num_Car = dr["month_pass_num_Car"].ToString();
                    imonth_pass_num_Car += Convert.ToInt32(dr["month_pass_num_Car"].ToString());

                    month_get_amount_num_Car = dr["month_get_amount_num_Car"].ToString();
                    imonth_get_amount_num_Car += Convert.ToInt32(dr["month_get_amount_num_Car"].ToString());

                    month_get_amount_PJ00048 = dr["month_get_amount_PJ00048"].ToString();
                    imonth_get_amount_PJ00048 += Convert.ToInt32(dr["month_get_amount_PJ00048"].ToString());

                    month_pass_amount_PJ00048 = dr["month_pass_amount_PJ00048"].ToString();
                    imonth_pass_amount_PJ00048 += Convert.ToInt32(dr["month_pass_amount_PJ00048"].ToString());
                    
                    if (iCar_SEQ == 1)
                    {
                        DataRow drCar_tit = dtCar.NewRow();
                        drCar_tit["SEQ"] = "";
                        drCar_tit["U_PFT_name"] = "";
                        drCar_tit["plan_name"] = CarTitle;
                        drCar_tit["day_get_amount_Car"] = AddDate;
                        dtCar.Rows.Add(drCar_tit);
                        dtCar.Rows.Add(FuncHandler.AddTitleByTable(dtCar, "Car"));

                        DataRow drCar = dtCar.NewRow();
                        drCar["SEQ"] = iCar_SEQ.ToString();
                        drCar["U_PFT_name"] = BC_Name+U_PFT_name;
                        drCar["plan_name"] = plan_name;
                        drCar["day_incase_num_PJ00048"] = day_incase_num_PJ00048;
                        drCar["month_incase_num_PJ00048"] = month_incase_num_PJ00048;
                        drCar["day_get_amount_num_Car"] = day_get_amount_num_Car;
                        drCar["day_get_amount_Car"] = day_get_amount_Car;
                        drCar["month_pass_num_Car"] = month_pass_num_Car;
                        drCar["month_get_amount_num_Car"] = month_get_amount_num_Car;
                        drCar["month_get_amount_PJ00048"] = month_get_amount_PJ00048;
                        drCar["month_pass_amount_PJ00048"] = month_pass_amount_PJ00048;
                        dtCar.Rows.Add(drCar);
                    }
                    else
                    {
                        DataRow drCar1 = dtCar.NewRow();
                        drCar1["SEQ"] = iCar_SEQ.ToString();
                        drCar1["U_PFT_name"] = BC_Name + U_PFT_name;
                        drCar1["plan_name"] =  plan_name;
                        drCar1["day_incase_num_PJ00048"] = day_incase_num_PJ00048;
                        drCar1["month_incase_num_PJ00048"] = month_incase_num_PJ00048;
                        drCar1["day_get_amount_num_Car"] = day_get_amount_num_Car;
                        drCar1["day_get_amount_Car"] = day_get_amount_Car;
                        drCar1["month_pass_num_Car"] = month_pass_num_Car;
                        drCar1["month_get_amount_num_Car"] = month_get_amount_num_Car;
                        drCar1["month_get_amount_PJ00048"] = month_get_amount_PJ00048;
                        drCar1["month_pass_amount_PJ00048"] = month_pass_amount_PJ00048;
                        dtCar.Rows.Add(drCar1);
                    }
                    iCar_SEQ++;
                }
                m_RowIdx++;
            }
            //汽車貸合計
            DataRow drCar_T = dtCar.NewRow();
            drCar_T["SEQ"] = "";
            drCar_T["U_PFT_name"] = "";
            drCar_T["plan_name"] = "合計";
            drCar_T["day_incase_num_PJ00048"] = iday_incase_num_PJ00048;
            drCar_T["month_incase_num_PJ00048"] = imonth_incase_num_PJ00048;
            drCar_T["day_get_amount_num_Car"] = iday_get_amount_num_Car;
            drCar_T["day_get_amount_Car"] = iday_get_amount_Car;
            drCar_T["month_pass_num_Car"] = imonth_pass_num_Car;
            drCar_T["month_get_amount_num_Car"] = imonth_get_amount_num_Car;
            drCar_T["month_get_amount_PJ00048"] = imonth_get_amount_PJ00048;
            drCar_T["month_pass_amount_PJ00048"] = imonth_pass_amount_PJ00048;
            dtCar.Rows.Add(drCar_T);
            //只有合計的sheet-房貸
            DataRow[] filteredRows = dtTotle.Select("SEQ = ''");
            dtTotle_T = filteredRows.CopyToDataTable();
            //只有合計的sheet-機車貸
            DataRow[] filteredRows_E = dtEngine.Select("SEQ = ''");
            dtEngine_T = filteredRows_E.CopyToDataTable();
            //只有合計的sheet-汽車貸
            DataRow[] filteredRows_C = dtCar.Select("SEQ = ''");
            dtCar_T = filteredRows_C.CopyToDataTable();
            //前一天只要總和
            if (isPreDay)
            {
                DsResult.Tables.Add(dtTotle_T);
                DsResult.Tables.Add(dtEngine_T);
                DsResult.Tables.Add(dtCar_T);
            }
            else
            {
                if (U_BC == "")
                {
                    DsResult.Tables.Add(dtTotle);
                    DsResult.Tables.Add(dtTotle_T);
                    DsResult.Tables.Add(dtEngine);
                    DsResult.Tables.Add(dtEngine_T);
                    DsResult.Tables.Add(dtCar);
                    DsResult.Tables.Add(dtCar_T);
                }
                else
                {
                    DsResult.Tables.Add(dtTotle);
                    DsResult.Tables.Add(dtTotle_T);
                    DsResult.Tables.Add(dtEngine);
                    DsResult.Tables.Add(dtEngine_T);
                }
            }

            return DsResult;
        }

        /// <summary>
        /// 取得客戶來電資料
        /// </summary>
        public IEnumerable<dynamic> GetIncoming(Incoming_req model)
        {
            
            #region SQL
            var T_SQL = @"SELECT InTime,COUNT(*) as _count FROM TelemarketingCSList tc
                          LEFT JOIN Telemarketing_M tm ON tm.HA_id = tc.TC_id AND tm.TM_type = 2
                          OUTER APPLY ( SELECT TOP 1 * FROM Telemarketing_D d WHERE d.TM_id = tm.TM_id AND tm.TM_type = 2 ORDER BY d.add_date DESC ) td
                          where ISNULL(InTime,'') <> '' and ISNULL(InDate,'') <> '' 
                          AND CONVERT(date, CAST(CAST(LEFT(InDate, CHARINDEX('/', InDate)-1) AS int) + 1911 AS varchar(4)) + '/' + 
                          SUBSTRING(InDate, CHARINDEX('/', InDate)+1, CHARINDEX('/', InDate, CHARINDEX('/', InDate)+1) - CHARINDEX('/', InDate) - 1) + '/' + 
                          RIGHT(InDate, LEN(InDate) - CHARINDEX('/', InDate, CHARINDEX('/', InDate)+1)), 111 ) 
                          BETWEEN @checkDateS AND @checkDateE ";
            var parameters = new List<SqlParameter> 
            {
                new SqlParameter("@checkDateS", model.checkDateS),
                new SqlParameter("@checkDateE", model.checkDateE)
            };
            if (!string.IsNullOrEmpty(model.TelAsk))
            {
                T_SQL += " and TelAsk = @TelAsk";
                parameters.Add(new SqlParameter("@TelAsk", model.TelAsk));
            }
            if (!string.IsNullOrEmpty(model.TelSour))
            {
                T_SQL += " and TelSour = @TelSour";
                parameters.Add(new SqlParameter("@TelSour", model.TelSour));
            }
            if (!string.IsNullOrEmpty(model.Fin_type))
            {
                T_SQL += " and td.Fin_type = @Fin_type";
                parameters.Add(new SqlParameter("@Fin_type", model.Fin_type));
            }
            T_SQL += " group by InTime order by InTime";
            #endregion
            var result = _adoData.ExecuteQuery(T_SQL, parameters).AsEnumerable().Select(row => new
            {
                InTime = row.Field<string>("InTime"),
                Count = row.Field<int>("_count")
            });
            return result;  
        }



       

        /// <summary>
        /// 取得客戶來電資料
        /// </summary>
        public DataTable GetSalesDataByDate(string BaseDate)
        {
           
            // 1. 解析傳入的字串 (建議使用 ParseExact 確保格式正確)
            // 假設輸入格式為 "yyyy/MM/dd" 或 "yyyy-MM-dd"
            if (!DateTime.TryParse(BaseDate, out DateTime parsedDate))
            {
                throw new ArgumentException("日期格式錯誤，請輸入 yyyy/MM/dd");
            }

            // 2. 計算 SQL 所需的四個日期標記
            // BaseDate: 設為該日的最後一秒 23:59:59
            DateTime baseDate = parsedDate.Date.AddDays(1).AddSeconds(-1);
            // StartDate: 該月 1 號 00:00:00
            DateTime startDate = new DateTime(baseDate.Year, baseDate.Month, 1);
            // PreBaseDate: 上個月的同一天
            DateTime preBaseDate = baseDate.AddMonths(-1);
            // PreStartDate: 上個月 1 號
            DateTime preStartDate = startDate.AddMonths(-1);

            #region SQL
            var T_SQL = @"
                    /*declare @BaseDate as Datetime,@StartDate as Datetime,@Pre_BaseDate as Datetime,@Pre_StartDate as Datetime
                    set @BaseDate=CAST('2026/1/20 23:59:59'as datetime)
                    set @StartDate=CAST(FORMAT(@BaseDate,'yyyy/MM')+'/01 00:00:00'as datetime)
                    set @Pre_BaseDate=DATEADD(month,-1,@BaseDate)
                    set @Pre_StartDate=DATEADD(month,-1,@StartDate)
                    select @BaseDate,@StartDate,@Pre_BaseDate,@Pre_StartDate*/
                    select u_bc,UC_Na,
                    sum(case when project_title='House' and add_date >=  @StartDate AND add_date <=  @BaseDate
                    then 1 else 0 end) Rate,
                    sum(case when project_title='House' and add_date >=  @Pre_StartDate AND add_date <=  @Pre_BaseDate
                    then 1 else 0 end) Pre_Rate,
                    sum(case when project_title='PJ00048' and add_date >=  @StartDate AND add_date <=  @BaseDate then 1 else 0 end) Car_Rate,
                    sum(case when project_title='PJ00048' and add_date >=  @Pre_StartDate AND add_date <=  @Pre_BaseDate then 1 else 0 end) Car_Pre_Rate,
                    sum(case when project_title in('PJ00046','PJ00047')  and add_date >=  @StartDate AND add_date <=  @BaseDate then 1 else 0 end) Engine_Rate, 
                    sum(case when project_title in('PJ00046','PJ00047')  and add_date >=  @Pre_StartDate AND add_date <=  @Pre_BaseDate then 1 else 0 end) Engine_Pre_Rate 
                    from(
                    select 'House' CaseKind,'House' project_title , ha.HA_id, case when ha.plan_num<>'N0001' then um.u_bc else 'BC0901' end u_bc,  ha.plan_num, um.U_name,  hp.pre_address,ha.add_date
	                    from User_M um
	                    left join House_apply ha on um.U_num = ha.plan_num
	                    left join house_pre hp on ha.HA_id = hp.HA_id
	                    Left Join (select H.HA_id,project_title from 
	                    House_sendcase H LEFT JOIN House_pre_project PP ON PP.HP_project_id = H.HP_project_id
	                    where H.del_tag='0' AND PP.del_tag='0'
	                    ) H on hp.HA_id=H.HA_id
	                    WHERE um.del_tag = '0'and hp.del_tag = '0' and ha.del_tag = '0' 
		                      AND hp.pre_process_type IN ( 'PRCT0002', 'PRCT0003', 'PRCT0005' )
		                      AND (ha.add_date >=  @Pre_StartDate AND ha.add_date <=  @BaseDate )
		                      AND (project_title not in ('PJ00048','PJ00047','PJ00046')or project_title is null)
	                    group by ha.HA_id, um.u_bc,  ha.plan_num, um.U_name,  hp.pre_address ,ha.add_date
                    union all
                    select 'Other' CaseKind,project_title, ha.HA_id, case when ha.plan_num<>'N0001' then um.u_bc else 'BC0901' end u_bc,  ha.plan_num, um.U_name,  hp.pre_address,ha.add_date
	                    from User_M um
	                    left join House_apply ha on um.U_num = ha.plan_num
	                    left join house_pre hp on ha.HA_id = hp.HA_id
	                    Left Join (select H.HA_id,project_title from 
	                    House_sendcase H LEFT JOIN House_pre_project PP ON PP.HP_project_id = H.HP_project_id
	                    where H.del_tag='0' AND PP.del_tag='0'
	                    ) H on hp.HA_id=H.HA_id
	                    WHERE um.del_tag = '0'and hp.del_tag = '0' and ha.del_tag = '0' 
		                      AND hp.pre_process_type IN ('PRCT0002','PRCT0003','PRCT0005')
		                      AND (ha.add_date >=  @Pre_StartDate AND ha.add_date <=  @BaseDate )
		                      AND (project_title in ('PJ00048','PJ00047','PJ00046'))
	                    group by project_title,ha.HA_id, um.u_bc,  ha.plan_num, um.U_name,  hp.pre_address ,ha.add_date) A
	                    Left join 
                     (
                    select item_D_code,item_D_name UC_Na,item_sort from Item_list  where item_M_code = 'branch_company' AND item_D_type='Y' AND show_tag='0' AND del_tag='0'
                    union all
	                select 'BC0901','湧立',1104
                    )U on A.u_bc=U.item_D_code
	                    Group by u_bc,UC_Na,item_sort order by item_sort
                     ";
            #endregion

            var parameters = new List<SqlParameter>
            {
                new SqlParameter("@BaseDate", baseDate),
                new SqlParameter("@StartDate", startDate),
                new SqlParameter("@Pre_BaseDate", preBaseDate),
                new SqlParameter("@Pre_StartDate", preStartDate)
               
            };
            DataTable dt = _adoData.ExecuteQuery(T_SQL, parameters);

            #region SQL
            T_SQL = @"
                        /*declare @BaseDate as Datetime,@Pre_BaseDate as Datetime
                    set @BaseDate=CAST('2026/1/26 23:59:59'as datetime)
                    set @Pre_BaseDate=DATEADD(month,-1,@BaseDate)*/

                    select u_bc,UC_Na
                    ,/*===========機車==============*/
                    /*月進件數*/
                     sum(CASE WHEN project_title  IN ('PJ00046','PJ00047')
		                      AND left(convert(varchar, Send_amount_date, 112), 6) = left(convert(varchar, @BaseDate, 112), 6)
                              AND convert(varchar, Send_amount_date, 112) <= convert(varchar, @BaseDate, 112) THEN 1
                              ELSE 0
                         END) AS Engine_month_incase,
                     /*前一個月進件數*/
                    sum(CASE WHEN  project_title   IN ('PJ00046','PJ00047')
                             AND left(convert(varchar, Send_amount_date, 112), 6) = left(convert(varchar, @Pre_BaseDate, 112), 6)
                             AND convert(varchar, Send_amount_date, 112) <= convert(varchar, @Pre_BaseDate, 112) THEN 1
                             ELSE 0
                        END) AS Engine_Pre_month_incase,
                     /*月撥款數*/
                     sum(CASE WHEN project_title  IN ('PJ00046','PJ00047')
                              AND left(convert(varchar, get_amount_date, 112), 6) = left(convert(varchar, @BaseDate, 112), 6)
                              AND convert(varchar, get_amount_date, 112) <=convert(varchar, @BaseDate, 112)
                              AND get_amount_type IN ('GTAT002') THEN 1
                              ELSE 0
                          END) AS Engine_month_get_amount_num,
                    /*前一個月撥款數*/
                    sum(CASE WHEN project_title IN ('PJ00046','PJ00047')
                             AND left(convert(varchar, get_amount_date, 112), 6) = left(convert(varchar, @Pre_BaseDate, 112), 6)
                             AND convert(varchar, get_amount_date, 112) <=convert(varchar, @Pre_BaseDate, 112)
                             AND get_amount_type IN ('GTAT002') THEN 1
	                         ELSE 0
                         END) AS Engine_Pre_month_get_amount_num,
                    /*月撥款額*/
                    sum(CASE WHEN project_title  IN ('PJ00046','PJ00047')
                               AND left(convert(varchar, get_amount_date, 112), 6) = left(convert(varchar, @BaseDate, 112), 6)
                               AND convert(varchar, get_amount_date, 112) <= convert(varchar, @BaseDate, 112) THEN get_amount
                               ELSE 0
                           END) AS Engine_month_get_amount,
                    /* 前一個月撥款額*/
                    sum(CASE WHEN project_title  IN ('PJ00046','PJ00047')
                             AND left(convert(varchar, get_amount_date, 112), 6) = left(convert(varchar, @Pre_BaseDate, 112), 6)
                             AND convert(varchar, get_amount_date, 112) <= convert(varchar, @Pre_BaseDate, 112) THEN get_amount
                             ELSE 0
                         END) AS Engine_Pre_month_get_amount ,
                    /* 已核未撥*/
                     sum(CASE WHEN  project_title IN ('PJ00046','PJ00047')
                              AND left(convert(varchar, Send_result_date, 112), 6) IN (left(convert(varchar, @BaseDate, 112), 6), left(convert(varchar,DATEADD(month,-1, @BaseDate), 112), 6),    left(convert(varchar,DATEADD(month,-2, @BaseDate), 112), 6))
                              AND convert(varchar, Send_result_date, 112) <=  convert(varchar, @BaseDate, 112)
                              AND Send_result_type = 'SRT002' AND isnull(check_amount_type, '') NOT IN ('CKAT003')
                              AND isnull(get_amount_type, '') NOT IN ('GTAT002', 'GTAT003') THEN pass_amount
                              ELSE 0
                          END) AS Engine_month_pass_amount,
                    /*前一個月 已核未撥*/
                    sum(CASE WHEN  project_title IN ('PJ00046','PJ00047')
                             AND left(convert(varchar, Send_result_date, 112), 6) IN (left(convert(varchar, @Pre_BaseDate, 112), 6), left(convert(varchar,DATEADD(month,-1, @Pre_BaseDate), 112), 6),    left(convert(varchar,DATEADD(month,-2, @Pre_BaseDate), 112), 6))
                             AND convert(varchar, Send_result_date, 112) <=  convert(varchar, @Pre_BaseDate, 112)
                             AND Send_result_type = 'SRT002'AND isnull(check_amount_type, '') NOT IN ('CKAT003')
                             AND isnull(get_amount_type, '') NOT IN ('GTAT002', 'GTAT003') THEN pass_amount
                             ELSE 0
                         END) AS Engine_Pre_month_pass_amount,
                    /*===========汽車==============*/
                     /*月進件數*/
                     sum(CASE WHEN project_title  IN ('PJ00048')
		                      AND left(convert(varchar, Send_amount_date, 112), 6) = left(convert(varchar, @BaseDate, 112), 6)
                              AND convert(varchar, Send_amount_date, 112) <= convert(varchar, @BaseDate, 112) THEN 1
                              ELSE 0
                         END) AS Car_month_incase,
                     /*前一個月進件數*/
                    sum(CASE WHEN  project_title   IN ('PJ00048')
                             AND left(convert(varchar, Send_amount_date, 112), 6) = left(convert(varchar, @Pre_BaseDate, 112), 6)
                             AND convert(varchar, Send_amount_date, 112) <= convert(varchar, @Pre_BaseDate, 112) THEN 1
                             ELSE 0
                        END) AS Car_Pre_month_incase,
                     /*月撥款數*/
                     sum(CASE WHEN project_title  IN ('PJ00048')
                              AND left(convert(varchar, get_amount_date, 112), 6) = left(convert(varchar, @BaseDate, 112), 6)
                              AND convert(varchar, get_amount_date, 112) <=convert(varchar, @BaseDate, 112)
                              AND get_amount_type IN ('GTAT002') THEN 1
                              ELSE 0
                          END) AS Car_month_get_amount_num,
                    /*前一個月撥款數*/
                    sum(CASE WHEN project_title IN ('PJ00048')
                             AND left(convert(varchar, get_amount_date, 112), 6) = left(convert(varchar, @Pre_BaseDate, 112), 6)
                             AND convert(varchar, get_amount_date, 112) <=convert(varchar, @Pre_BaseDate, 112)
                             AND get_amount_type IN ('GTAT002') THEN 1
	                         ELSE 0
                         END) AS Car_Pre_month_get_amount_num,
                    /*月撥款額*/
                    sum(CASE WHEN project_title  IN ('PJ00048')
                               AND left(convert(varchar, get_amount_date, 112), 6) = left(convert(varchar, @BaseDate, 112), 6)
                               AND convert(varchar, get_amount_date, 112) <= convert(varchar, @BaseDate, 112) THEN get_amount
                               ELSE 0
                           END) AS Car_month_get_amount,
                    /* 前一個月撥款額*/
                    sum(CASE WHEN project_title  IN ('PJ00048')
                             AND left(convert(varchar, get_amount_date, 112), 6) = left(convert(varchar, @Pre_BaseDate, 112), 6)
                             AND convert(varchar, get_amount_date, 112) <= convert(varchar, @Pre_BaseDate, 112) THEN get_amount
                             ELSE 0
                         END) AS Car_Pre_month_get_amount ,
                    /* 已核未撥*/
                     sum(CASE WHEN  project_title IN ('PJ00048')
                              AND left(convert(varchar, Send_result_date, 112), 6) IN (left(convert(varchar, @BaseDate, 112), 6), left(convert(varchar,DATEADD(month,-1, @BaseDate), 112), 6),    left(convert(varchar,DATEADD(month,-2, @BaseDate), 112), 6))
                              AND convert(varchar, Send_result_date, 112) <=  convert(varchar, @BaseDate, 112)
                              AND Send_result_type = 'SRT002' AND isnull(check_amount_type, '') NOT IN ('CKAT003')
                              AND isnull(get_amount_type, '') NOT IN ('GTAT002', 'GTAT003') THEN pass_amount
                              ELSE 0
                          END) AS Car_month_pass_amount,
                    /*前一個月 已核未撥*/
                    sum(CASE WHEN  project_title IN ('PJ00048')
                             AND left(convert(varchar, Send_result_date, 112), 6) IN (left(convert(varchar, @Pre_BaseDate, 112), 6), left(convert(varchar,DATEADD(month,-1, @Pre_BaseDate), 112), 6),    left(convert(varchar,DATEADD(month,-2, @Pre_BaseDate), 112), 6))
                             AND convert(varchar, Send_result_date, 112) <=  convert(varchar, @Pre_BaseDate, 112)
                             AND Send_result_type = 'SRT002'AND isnull(check_amount_type, '') NOT IN ('CKAT003')
                             AND isnull(get_amount_type, '') NOT IN ('GTAT002', 'GTAT003') THEN pass_amount
                             ELSE 0
                         END) AS Car_Pre_month_pass_amount,
                    /*===========房貸==============*/
                     /*月進件數*/
                     sum(CASE WHEN project_title not IN ('PJ00046','PJ00047','PJ00048')
		                      AND left(convert(varchar, Send_amount_date, 112), 6) = left(convert(varchar, @BaseDate, 112), 6)
                              AND convert(varchar, Send_amount_date, 112) <= convert(varchar, @BaseDate, 112) THEN 1
                              ELSE 0
                         END) AS month_incase,
                     /*前一個月進件數*/
                    sum(CASE WHEN  project_title not IN ('PJ00046','PJ00047','PJ00048')
                             AND left(convert(varchar, Send_amount_date, 112), 6) = left(convert(varchar, @Pre_BaseDate, 112), 6)
                             AND convert(varchar, Send_amount_date, 112) <= convert(varchar, @Pre_BaseDate, 112) THEN 1
                             ELSE 0
                        END) AS Pre_month_incase,
                     /*月撥款數*/
                     sum(CASE WHEN project_title not IN ('PJ00046','PJ00047','PJ00048')
                              AND left(convert(varchar, get_amount_date, 112), 6) = left(convert(varchar, @BaseDate, 112), 6)
                              AND convert(varchar, get_amount_date, 112) <=convert(varchar, @BaseDate, 112)
                              AND get_amount_type IN ('GTAT002') THEN 1
                              ELSE 0
                          END) AS month_get_amount_num,
                    /*前一個月撥款數*/
                    sum(CASE WHEN project_title not IN ('PJ00046','PJ00047','PJ00048')
                             AND left(convert(varchar, get_amount_date, 112), 6) = left(convert(varchar, @Pre_BaseDate, 112), 6)
                             AND convert(varchar, get_amount_date, 112) <=convert(varchar, @Pre_BaseDate, 112)
                             AND get_amount_type IN ('GTAT002') THEN 1
	                         ELSE 0
                         END) AS Pre_month_get_amount_num,
                    /*月撥款額*/
                    sum(CASE WHEN project_title not IN ('PJ00046','PJ00047','PJ00048')
                               AND left(convert(varchar, get_amount_date, 112), 6) = left(convert(varchar, @BaseDate, 112), 6)
                               AND convert(varchar, get_amount_date, 112) <= convert(varchar, @BaseDate, 112) THEN get_amount
                               ELSE 0
                           END) AS month_get_amount,
                    /* 前一個月撥款額*/
                    sum(CASE WHEN project_title not IN ('PJ00046','PJ00047','PJ00048')
                             AND left(convert(varchar, get_amount_date, 112), 6) = left(convert(varchar, @Pre_BaseDate, 112), 6)
                             AND convert(varchar, get_amount_date, 112) <= convert(varchar, @Pre_BaseDate, 112) THEN get_amount
                             ELSE 0
                         END) AS Pre_month_get_amount ,
                    /* 已核未撥*/
                     sum(CASE WHEN  project_title not IN ('PJ00046','PJ00047','PJ00048')
                              AND left(convert(varchar, Send_result_date, 112), 6) IN (left(convert(varchar, @BaseDate, 112), 6), left(convert(varchar,DATEADD(month,-1, @BaseDate), 112), 6),    left(convert(varchar,DATEADD(month,-2, @BaseDate), 112), 6))
                              AND convert(varchar, Send_result_date, 112) <=  convert(varchar, @BaseDate, 112)
                              AND Send_result_type = 'SRT002' AND isnull(check_amount_type, '') NOT IN ('CKAT003')
                              AND isnull(get_amount_type, '') NOT IN ('GTAT002', 'GTAT003') THEN pass_amount
                              ELSE 0
                          END) AS month_pass_amount,
                    /*前一個月 已核未撥*/
                    sum(CASE WHEN  project_title not IN ('PJ00046','PJ00047','PJ00048')
                             AND left(convert(varchar, Send_result_date, 112), 6) IN (left(convert(varchar, @Pre_BaseDate, 112), 6), left(convert(varchar,DATEADD(month,-1, @Pre_BaseDate), 112), 6),    left(convert(varchar,DATEADD(month,-2, @Pre_BaseDate), 112), 6))
                             AND convert(varchar, Send_result_date, 112) <=  convert(varchar, @Pre_BaseDate, 112)
                             AND Send_result_type = 'SRT002'AND isnull(check_amount_type, '') NOT IN ('CKAT003')
                             AND isnull(get_amount_type, '') NOT IN ('GTAT002', 'GTAT003') THEN pass_amount
                             ELSE 0
                         END) AS Pre_month_pass_amount
                    from(
                    select  project_title , ha.HA_id, case when ha.plan_num<>'N0001' then um.u_bc else 'BC0901' end u_bc,  ha.plan_num, um.U_name,ha.add_date
                    ,Send_result_date,get_amount_date,Send_result_type,get_amount_type,check_amount_type,pass_amount,get_amount,Send_amount_date
	                    from User_M um
	                    left join House_apply ha on um.U_num = ha.plan_num
	                    left join House_sendcase hs on ha.HA_id = hs.HA_id
	                    LEFT JOIN House_pre_project hpp ON hpp.HP_project_id = hs.HP_project_id
	                    WHERE  ha.del_tag = '0' and hs.del_tag = '0' 
		                      AND (left(convert(varchar, Send_amount_date, 112), 6) IN( left(convert(varchar, @BaseDate, 112), 6),left(convert(varchar, @Pre_BaseDate, 112), 6))
                            OR left(convert(varchar, Send_result_date, 112), 6) between left(convert(varchar,DATEADD(month,-3, @BaseDate), 112), 6) and left(convert(varchar, @BaseDate, 112), 6)
                            OR left(convert(varchar, get_amount_date, 112), 6) IN( left(convert(varchar, @BaseDate, 112), 6),left(convert(varchar, @Pre_BaseDate, 112), 6)) )
	                    ) A
	                    Left join (
                    select item_D_code,item_D_name UC_Na,item_sort from Item_list  where item_M_code = 'branch_company' AND item_D_type='Y' AND show_tag='0' AND del_tag='0'
                    union all
	                select 'BC0901','湧立',1104
                    )U on A.u_bc=U.item_D_code
                    Group by u_bc,UC_Na,item_sort order by item_sort


                     ";
            #endregion
            var parameters1 = new List<SqlParameter>
            {
                new SqlParameter("@BaseDate", baseDate),
                new SqlParameter("@Pre_BaseDate", preBaseDate)
              
            };
            DataTable dt1 = _adoData.ExecuteQuery(T_SQL, parameters1);
            FuncHandler func =new FuncHandler();
            
            return func.MergeSalesDataToDataTable(dt, dt1); 
        }


    }
}
