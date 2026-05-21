using KF_WebAPI.BaseClass;
using KF_WebAPI.BaseClass.AE;
using KF_WebAPI.BaseClass.Winton;
using KF_WebAPI.Controllers;
using KF_WebAPI.FunctionHandler;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections;
using System.Data;
using System.Drawing;
using System.Dynamic;
using System.Net.NetworkInformation;
using System.Reflection;
using System.Reflection.PortableExecutable;
using System.Xml.Linq;

namespace KF_WebAPI.DataLogic
{
    public class AE_ACC
    {
        ADOData _adoData = new ADOData();
        FuncHandler _Fun = new FuncHandler();

        public List<RC_ACH_Res> GetRcACH_LQuery(int yyyyMM, string type, string pjtype)
        {
            try
            {
                var parameters = new List<SqlParameter>();
                var T_SQL = @"select rm.RCM_id,b.HS_id,b.Send_amount_date,ha.CS_name,um.U_name,li_1.item_D_name as U_BC,b.get_amount
                              ,b.get_amount_date,b.interest_rate_pass,li_2.item_D_name as pjName,rm.month_total,b.Loan_rate,li_3.item_D_name as str_Ach_State,Ach_Note
                              ,rm.RCM_cknum,(select COUNT(*) from AE_Files where KeyID = rm.RCM_cknum) as FileCount                              
                              from view_HS_Base b
                              INNER JOIN Receivable_M rm ON rm.HS_id = b.HS_id AND rm.del_tag = 0
                              INNER JOIN House_apply ha ON ha.HA_id = b.HA_id AND ha.del_tag = 0
                              LEFT JOIN User_M um ON um.U_num = ha.plan_num
                              LEFT JOIN Item_list li_1 ON li_1.item_M_code = 'branch_company' and li_1.item_D_code = um.U_BC
                              LEFT JOIN Item_list li_2 ON li_2.item_M_code = 'project_title' and li_2.item_D_code = b.project_title
                              LEFT JOIN Item_list li_3 ON li_3.item_M_code = 'Ach_State' and li_3.item_D_code = rm.Ach_State
                              WHERE b.Send_result_type = 'SRT002' AND b.get_amount_type = 'GTAT002'";

                switch (pjtype)
                {
                    case "Moto":
                        T_SQL += @" AND b.project_title IN ('PJ00046','PJ00047')";
                        break;
                    case "Car":
                        T_SQL += @" AND b.project_title IN ('PJ00048')";
                        break;
                    default:
                        T_SQL += @" AND b.project_title NOT IN ('PJ00046','PJ00047','PJ00048')";
                        break;
                }

                switch (type)
                {
                    case "M":
                        T_SQL += @" AND (YEAR(b.get_amount_date)*100+MONTH(b.get_amount_date)) = @TargetMonth 
                               ORDER BY b.get_amount_date ,b.HS_id";
                        parameters.Add(new SqlParameter("@TargetMonth", yyyyMM));
                        break;
                    case "Y":
                        T_SQL += @" AND (YEAR(b.get_amount_date)*100) = @TargetYear 
                               ORDER BY b.get_amount_date ,b.HS_id";
                        int yyyy = (yyyyMM / 100) * 100;
                        parameters.Add(new SqlParameter("@TargetYear", yyyy));
                        break;
                    default:
                        T_SQL += @" ORDER BY b.get_amount_date ,b.HS_id";
                        break;
                }

                var result = _adoData.ExecuteQuery(T_SQL, parameters).AsEnumerable().Select(row=> new RC_ACH_Res
                {
                    RCM_id = row.Field<decimal>("RCM_id"),
                    HS_id = row.Field<decimal>("HS_id"),
                    str_Send_amount_date = FuncHandler.ConvertGregorianToROC(row.Field<DateTime>("Send_amount_date").ToString("yyyy/MM/dd")),
                    CS_name = _Fun.DeCodeBNWords(row.Field<string>("CS_name")),
                    U_name = _Fun.DeCodeBNWords(row.Field<string>("U_name")),
                    U_BC = row.Field<string>("U_BC"),
                    get_amount = row.Field<string>("get_amount"),
                    str_get_amount_date = FuncHandler.ConvertGregorianToROC(row.Field<DateTime>("get_amount_date").ToString("yyyy/MM/dd")),
                    interest_rate_pass = row.Field<string>("interest_rate_pass"),
                    pjName = _Fun.DeCodeBNWords(row.Field<string>("pjName")),
                    month_total = row.Field<int>("month_total"),
                    Loan_rate = row.Field<string>("Loan_rate"),
                    str_Ach_State = row.Field<string>("str_Ach_State"),
                    Ach_Note = row.Field<string>("Ach_Note"),
                    RCM_cknum = row.Field<string>("RCM_cknum"),
                    FileCount = row.Field<int>("FileCount")
                }).ToList();

                return result;
            }
            catch (Exception)
            {

                throw;
            }
            
        }

        public ResultClass<string> RC_Ach_SQuery(string Rcm_id)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            try
            {
                var parameters = new List<SqlParameter>();
                var T_SQL = @"SELECT ha.CS_name,rm.*,(select COUNT(*) from AE_Files where KeyID = rm.RCM_cknum) as FileCount
                              FROM Receivable_M rm
                              INNER JOIN House_apply ha ON ha.HA_id = rm.HA_id AND ha.del_tag = 0
                              WHERE RCM_id = @Rcm_id";
                parameters.Add(new SqlParameter("@Rcm_id", Rcm_id));
                var result = _adoData.ExecuteQuery(T_SQL, parameters).AsEnumerable().Select(row => new {
                    RCM_id = row.Field<decimal>("RCM_id"),
                    CS_name = _Fun.DeCodeBNWords(row.Field<string>("CS_name")),
                    Ach_State = row.Field<string>("Ach_State"),
                    Ach_Note = row.Field<string>("Ach_Note"),
                    RCM_cknum = row.Field<string>("RCM_cknum"),
                    FileCount = row.Field<int>("FileCount")
                }).ToList();

                resultClass.ResultCode = "000";
                resultClass.objResult = JsonConvert.SerializeObject(result);

                return resultClass;
            }
            catch (Exception)
            {

                throw;
            }
        }

        public ResultClass<string> RC_Ach_Upd(RC_Ach_Ins model)
        {
            try
            {
                ResultClass<string> resultClass = new ResultClass<string>();
                var T_SQL = @"Update Receivable_M Set Ach_State=@Ach_State,Ach_Note=@Ach_Note,BankNo=@BankNo,AccountNo=@AccountNo WHERE RCM_id = @Rcm_id";
                var parameters = new List<SqlParameter>()
                {
                    new SqlParameter("@Ach_State",model.Ach_State),
                    new SqlParameter("@Ach_Note",model.Ach_Note),
                    new SqlParameter("@BankNo",model.BankNo),
                    new SqlParameter("@AccountNo",model.AccountNo),
                    new SqlParameter("@Rcm_id",model.RCM_id)
                };
                int result = _adoData.ExecuteNonQuery(T_SQL, parameters);

                if (result == 0)
                {
                    resultClass.ResultCode = "400";
                    resultClass.ResultMsg = "異動失敗";
                }
                else
                {
                    resultClass.ResultCode = "000";
                    resultClass.ResultMsg = "異動成功";
                }
                return resultClass;
            }
            catch (Exception)
            {

                throw;
            }
        }

        public byte[] GetRcAchExcel(int yyyyMM, string type, string pjtype)
        {
            try
            {
                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("撥款清冊");
                    #region 撥款資料
                    var rcList = GetRcACH_LQuery(yyyyMM, type, pjtype);

                    string[] headers = { "件數", "案件編號", "進件日期", "申請人", "經辦人", "區域", "撥款金額", "撥款日期", "利率", "專案", "期數", "成數", "ACH", "ACH備註" };

                    int rowIndex = 1;
                    int colIndex = 1;
                    foreach (var header in headers)
                    {
                        var cell = worksheet.Cells[rowIndex, colIndex++];
                        cell.Value = header;
                        // 設置儲存格底色為淺藍色
                        cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        cell.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                    }

                    

                    // 添加表身
                    colIndex = 1;
                    int index = 1;
                    foreach  (var item in rcList)
                    {
                        rowIndex++;
                        worksheet.Cells[rowIndex, colIndex++].Value = index++;
                        worksheet.Cells[rowIndex, colIndex++].Value = item.HS_id;
                        worksheet.Cells[rowIndex, colIndex++].Value = item.str_Send_amount_date;
                        worksheet.Cells[rowIndex, colIndex++].Value = item.CS_name;
                        worksheet.Cells[rowIndex, colIndex++].Value = item.U_name;
                        worksheet.Cells[rowIndex, colIndex++].Value = item.U_BC;
                        worksheet.Cells[rowIndex, colIndex++].Value = item.get_amount;
                        worksheet.Cells[rowIndex, colIndex++].Value = item.str_get_amount_date;
                        worksheet.Cells[rowIndex, colIndex++].Value = item.interest_rate_pass;
                        worksheet.Cells[rowIndex, colIndex++].Value = item.pjName;
                        worksheet.Cells[rowIndex, colIndex++].Value = item.month_total;
                        worksheet.Cells[rowIndex, colIndex++].Value = item.Loan_rate;
                        worksheet.Cells[rowIndex, colIndex++].Value = item.str_Ach_State;
                        worksheet.Cells[rowIndex, colIndex++].Value = item.Ach_Note;
                        colIndex = 1;
                    }
                    // 框線
                    using (var range = worksheet.Cells[1, 1, rowIndex, headers.Length])
                    {
                        range.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        range.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        range.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    }
                    #endregion

                    // 調整列寬
                    worksheet.Cells[1, 1, rowIndex, headers.Length].AutoFitColumns();

                    return package.GetAsByteArray();
                }
            }
            catch (Exception)
            {

                throw;
            }
        }
    }
}
