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

        public List<RC_ACH_Res> GetRcACH_LQuery(int yyyyMM, string type, string pjtype, string CS_Name, string Ach_State)
        {
            try
            {
                var parameters = new List<SqlParameter>();
                var T_SQL = @"select ha.ha_id,rm.RCM_id,b.HS_id,b.Send_amount_date,ha.CS_name,um.U_name,li_1.item_D_name as U_BC,b.get_amount
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
                if (CS_Name != "")
                {
                    T_SQL += @" AND CS_Name like @CS_Name+'%'";
                    parameters.Add(new SqlParameter("@CS_Name", CS_Name));
                }

                if (Ach_State != "")
                {
                    T_SQL += @"  and isnull( rm.Ach_State,'NP')=@Ach_State ";
                    parameters.Add(new SqlParameter("@Ach_State", Ach_State));
                }


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
                    HA_id = row.Field<decimal>("HA_id"),
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
                var T_SQL = @"SELECT  case when  rd.RCM_id is null then 'N' else 'Y' end isPayOff , ha.CS_name,ha.CS_company_TaxNum,rm.*,(select COUNT(*) from AE_Files where KeyID = rm.RCM_cknum) as FileCount
                              FROM Receivable_M rm Left join (select distinct RCM_ID from  Receivable_D where check_pay_type='S') rd on rm.RCM_id=rd.RCM_id
                              INNER JOIN House_apply ha ON ha.HA_id = rm.HA_id AND ha.del_tag = 0
                              WHERE rm.RCM_id = @Rcm_id";
                parameters.Add(new SqlParameter("@Rcm_id", Rcm_id));
                var result = _adoData.ExecuteQuery(T_SQL, parameters).AsEnumerable().Select(row => new {
                    RCM_id = row.Field<decimal>("RCM_id"),
                    CS_name = _Fun.DeCodeBNWords(row.Field<string>("CS_name")),
                    Ach_State = row.Field<string>("Ach_State"),
                    Ach_Note = row.Field<string>("Ach_Note"),
                    BankNo = row.Field<string>("BankNo"),
                    AccountNo = row.Field<string>("AccountNo"),
                    RCM_cknum = row.Field<string>("RCM_cknum"),
                    FileCount = row.Field<int>("FileCount"),
                    isPayOff = row.Field<string>("isPayOff"),
                    CS_company_TaxNum = row.Field<string>("CS_company_TaxNum") 
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


                if (model.CS_company_TaxNum != null && model.CS_company_TaxNum != "")
                {
                    var T_SQL1 = @"Update House_apply Set CS_company_TaxNum=@CS_company_TaxNum  WHERE HA_id = @HA_id";
                    var parameters1 = new List<SqlParameter>()
                    {
                        new SqlParameter("@HA_id",model.HA_id),
                        new SqlParameter("@CS_company_TaxNum",model.CS_company_TaxNum)
                    };
                    int result1 = _adoData.ExecuteNonQuery(T_SQL1, parameters1);
                }

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
                    var rcList = GetRcACH_LQuery(yyyyMM, type, pjtype,"","");

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

        public List<Receivable_Debt_res> RC_Debt_LQuery(string? csName,string dateS,string dateE,string mType,string pjType)
        {
            try
            {
                var parameters = new List<SqlParameter>();
                var T_SQL = @"SELECT I.item_D_name AS_Name,REPLACE(REPLACE(Expe_note, CHAR(13), ''), CHAR(10), '<br>')disExpe_note,Expe_note,RP.Ex_RemainingPrincipal,RP.interest,
                              RP.Rmoney,Receivable_D.*,DATEDIFF(DAY, Receivable_D.RC_date , SYSDATETIME()) DiffDay,House_apply.CS_name,Receivable_M.RCM_cknum,
                              House_sendcase.interest_rate_pass,Receivable_M.amount_total,Receivable_M.month_total,Receivable_M.amount_per_month,Receivable_M.date_begin,
                              Receivable_M.RCM_note,Receivable_M.loan_grace_num,Receivable_M.auction_status,
                              (SELECT item_D_name FROM Item_list WHERE item_M_code = 'project_title' AND item_D_code = House_pre_project.project_title) AS project_name,
                              (SELECT U_name FROM User_M WHERE U_num = Receivable_D.add_num AND del_tag='0') AS add_name,
                              isnull( (SELECT U_name FROM User_M WHERE U_num = Receivable_D.check_pay_num AND del_tag='0'),'') AS check_pay_name,
                              isnull( (SELECT U_name FROM User_M WHERE U_num = Receivable_D.bad_debt_num AND del_tag='0'),'') AS bad_debt_name,
                              isnull( (SELECT U_name FROM User_M WHERE U_num = Receivable_D.cancel_num AND del_tag='0'),'') AS cancel_name,Item_list.item_D_name AS U_BC_name,
                              (SELECT ISNULL(Item_list.item_D_name, U_name) FROM User_M LEFT JOIN Item_list ON item_M_code = 'SpecName' AND item_D_type = 'Y' AND item_D_txt_A = U_num WHERE U_num = House_apply.plan_num) AS plan_name
                              FROM (SELECT bad_debt_type,check_pay_type,cancel_type,RC_amount,RCM_id,cancel_num,bad_debt_num,check_pay_num,add_num,min (RC_count) RC_count,
                              min (RC_date) RC_date FROM Receivable_D WHERE del_tag = '0' AND check_pay_type='N' AND bad_debt_type='N' AND cancel_type='N'
                              GROUP BY bad_debt_type,check_pay_type,cancel_type,RCM_id,cancel_num,bad_debt_num,check_pay_num,add_num,RC_amount) Receivable_D
                              LEFT JOIN Receivable_M ON Receivable_M.RCM_id = Receivable_D.RCM_id
                              LEFT JOIN Receivable_D RP ON Receivable_D.RCM_id = RP.RCM_id AND Receivable_D.RC_count = RP.RC_count
                              LEFT JOIN House_apply ON House_apply.HA_id = Receivable_M.HA_id
                              LEFT JOIN (SELECT U_num,U_BC FROM User_M) User_M ON User_M.U_num = House_apply.plan_num
                              LEFT JOIN Item_list ON item_M_code = 'branch_company' AND item_D_code = User_M.U_BC
                              LEFT JOIN House_sendcase ON House_sendcase.HS_id = Receivable_M.HS_id
                              LEFT JOIN House_pre_project ON House_pre_project.HP_project_id = House_sendcase.HP_project_id AND House_pre_project.del_tag='0'
                              LEFT JOIN Item_list I ON I.item_M_code = 'auction_status' AND I.item_D_code = auction_status
                              WHERE House_sendcase.del_tag='0' AND House_apply.del_tag='0' AND Receivable_M.del_tag='0' 
                              AND (Receivable_D.RC_date >= @dateS AND Receivable_D.RC_date <= @dateE)
                              AND User_M.U_BC IN ('zz','BC0100','BC0200','BC0600','BC0900','BC0700','BC0800','BC0300','BC0500','BC0400','BC0800','BC0701')";
                if (!string.IsNullOrEmpty(csName))
                {
                    T_SQL += @" AND (House_apply.CS_name=@CS_name)";
                    parameters.Add(new SqlParameter("@CS_name", csName));
                }
                switch (mType)
                {
                    case "M1":
                        T_SQL += @" AND((DATEDIFF(DAY, Receivable_D.RC_date , SYSDATETIME()) BETWEEN 31 AND 60))";
                        break;
                    case "M2":
                        T_SQL += @" AND((DATEDIFF(DAY, Receivable_D.RC_date , SYSDATETIME()) BETWEEN 61 AND 90))";
                        break;
                    case "M3":
                        T_SQL += @" AND((DATEDIFF(DAY, Receivable_D.RC_date , SYSDATETIME()) >= 91))";
                        break;
                    default:
                        T_SQL += @" AND((DATEDIFF(DAY, Receivable_D.RC_date , SYSDATETIME()) BETWEEN 1 AND 30))";
                        break;
                }
                switch (pjType)
                {
                    case "House":
                        T_SQL += @" AND project_title NOT IN ('PJ00046','PJ00047','PJ00048','PJ00998')";
                        break;
                    case "Moto":
                        T_SQL += @" AND project_title IN ('PJ00046','PJ00047')";
                        break;
                    case "Car":
                        T_SQL += @" AND project_title IN ('PJ00048','PJ00998')";
                        break;
                }
                T_SQL += @" ORDER BY Receivable_D.RC_date";
                parameters.Add(new SqlParameter("@dateS", dateS));
                parameters.Add(new SqlParameter("@dateE", dateE));

                var result = _adoData.ExecuteQuery(T_SQL, parameters).AsEnumerable().Select(row => new Receivable_Debt_res
                {
                    AS_Name = _Fun.DeCodeBNWords(row.Field<string>("AS_Name")),
                    disExpe_note = row.Field<string>("disExpe_note"),
                    Expe_note = row.Field<string>("Expe_note"),
                    Ex_RemainingPrincipal = row.Field<decimal>("Ex_RemainingPrincipal"),
                    interest = row.Field<decimal>("interest"),
                    Rmoney = row.Field<decimal>("Rmoney"),
                    bad_debt_type = row.Field<string>("bad_debt_type"),
                    check_pay_type = row.Field<string>("check_pay_type"),
                    cancel_type = row.Field<string>("cancel_type"),
                    RC_amount = row.Field<decimal>("RC_amount"),
                    RCM_id = row.Field<decimal>("RCM_id"),
                    cancel_num = row.Field<string>("cancel_num"),
                    bad_debt_num = row.Field<string>("bad_debt_num"),
                    check_pay_num = row.Field<string>("check_pay_num"),
                    add_num = row.Field<string>("add_num"),
                    RC_count = row.Field<int>("RC_count"),
                    RC_date = row.Field<DateTime>("RC_date"),
                    DiffDay = row.Field<int>("DiffDay"),
                    CS_name = _Fun.DeCodeBNWords(row.Field<string>("CS_name")),
                    RCM_cknum = row.Field<string>("RCM_cknum"),
                    interest_rate_pass = row.Field<string>("interest_rate_pass"),
                    amount_total = row.Field<decimal>("amount_total"),
                    month_total = row.Field<int>("month_total"),
                    amount_per_month = row.Field<decimal>("amount_per_month"),
                    date_begin = row.Field<DateTime>("date_begin"),
                    RCM_note = row.Field<string>("RCM_note"),
                    loan_grace_num = row.Field<int?>("loan_grace_num"),
                    auction_status = row.Field<string>("auction_status"),
                    project_name = _Fun.DeCodeBNWords(row.Field<string>("project_name")),
                    add_name = _Fun.DeCodeBNWords(row.Field<string>("add_name")),
                    check_pay_name = _Fun.DeCodeBNWords(row.Field<string>("check_pay_name")),
                    bad_debt_name = _Fun.DeCodeBNWords(row.Field<string>("bad_debt_name")),
                    cancel_name = _Fun.DeCodeBNWords(row.Field<string>("cancel_name")),
                    U_BC_name = _Fun.DeCodeBNWords(row.Field<string>("U_BC_name")),
                    plan_name = _Fun.DeCodeBNWords(row.Field<string>("plan_name"))
                }).ToList();

                return result;
            }
            catch (Exception)
            {

                throw;
            }
        }
    }
}
