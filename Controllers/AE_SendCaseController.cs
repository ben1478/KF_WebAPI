using KF_WebAPI.BaseClass;
using KF_WebAPI.BaseClass.AE;
using KF_WebAPI.FunctionHandler;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Newtonsoft.Json;
using System.Data;
using System.Text;

namespace KF_WebAPI.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class AE_SendCaseController : ControllerBase
    {
        /// <summary>
        /// 撥款及費用確認書列表查詢
        /// </summary>
        [HttpPost("House_SendCase_LQuery")]
        public ActionResult<ResultClass<string>> House_SendCase_LQuery(House_sendcase_Req model)
        {
            House_sendcase_Req cc = new KF_WebAPI.BaseClass.AE.House_sendcase_Req();
            ResultClass<string> resultClass = new ResultClass<string>();
            var sqlBuilder = new StringBuilder(
                @"SELECT 
                    'FEE-' + cast(HS_ID as varchar)[File_ID],
                    M.U_BC, M.U_name,
                    case when CancelDate is null then 'N' else 'Y' end isCancel,
                    isnull(Comparison, '')Comparison, project_title,interest_rate_pass refRateI, Loan_rate refRateL, 
                    isnull(I.Introducer_PID, '') I_PID,      
                    case when H.act_perf_amt is null  then 'N' else 'Y' end IsConfirm,
                    isnull(H.Introducer_PID, '')Introducer_PID,              
                    isnull(I_Count, 0)I_Count,H.HS_id, M.U_BC_name,A.CS_name,      
                    isnull(convert(varchar(4), (convert(varchar(4), misaligned_date, 126) - 1911)) + '-' + convert(varchar(2), dbo.PadWithZero(month(misaligned_date))) + '-' + convert(varchar(2), dbo.PadWithZero(day(misaligned_date))), '') misaligned_date,     
                    convert(varchar(4), (convert(varchar(4), Send_amount_date, 126) - 1911)) + '-' + convert(varchar(2), dbo.PadWithZero(month(Send_amount_date))) + '-' + convert(varchar(2), dbo.PadWithZero(day(Send_amount_date)))Send_amount_date,  
                    convert(varchar(4), (convert(varchar(4), get_amount_date, 126) - 1911)) + '-' + convert(varchar(2), dbo.PadWithZero(month(get_amount_date))) + '-' + convert(varchar(2), dbo.PadWithZero(day(get_amount_date)))get_amount_date,  
                    convert(varchar(4), (convert(varchar(4), H.Send_result_date, 126) - 1911)) + '-' + convert(varchar(2), dbo.PadWithZero(month(H.Send_result_date))) + '-' + convert(varchar(2), dbo.PadWithZero(day(H.Send_result_date)))Send_result_date,   
                    H.CS_introducer,M.U_name plan_name, H.pass_amount, get_amount,        
                    (SELECT item_D_name FROM Item_list WHERE item_M_code = 'fund_company' AND item_D_type = 'Y' AND item_D_code = H.fund_company AND show_tag = '0' AND del_tag = '0') AS show_fund_company,
                    (SELECT item_D_name FROM Item_list WHERE item_M_code = 'project_title' AND item_D_type = 'Y' AND item_D_code = House_pre_project.project_title  AND show_tag = '0' AND del_tag = '0' ) AS show_project_title,
                    Loan_rate+'%' Loan_rate,interest_rate_original + '%' interest_rate_original,interest_rate_pass + '%' interest_rate_pass,     
                    isnull(charge_flow, 0)charge_flow,isnull(charge_agent, 0)charge_agent, 
                    isnull(charge_check, 0)charge_check,isnull(get_amount_final, 0)get_amount_final, 
                    isnull(H.subsidized_interest, 0)subsidized_interest, R.Comm_Remark , I.Bank_name, I.Bank_account Bank_account,
                    isnull((select count(*) FROM ASP_UpLoad where cknum='FEE-'+cast(H.HS_ID as varchar) AND del_tag='0'),'0') as upLoad_Count
                    FROM House_sendcase H
                        LEFT JOIN House_apply A ON A.HA_id = H.HA_id AND A.del_tag = '0'
                        LEFT JOIN House_pre_project ON House_pre_project.HP_project_id = H.HP_project_id  AND House_pre_project.del_tag = '0'
                        LEFT JOIN(select u.*,item_D_name U_BC_name from User_M u LEFT JOIN Item_list ub on ub.item_M_code = 'branch_company' AND ub.item_D_type = 'Y' AND ub.item_D_code = u.U_BC)  M ON M.U_num = A.plan_num
                        LEFT JOIN User_M Users ON house_pre_project.fin_user = Users.U_num
                        LEFT JOIN(select* from Introducer_Comm where del_tag= '0') I ON replace(H.CS_introducer, ';', '') = I.Introducer_name
                            and case when H.Introducer_PID is null then I.Introducer_PID else H.Introducer_PID end = I.Introducer_PID
                        LEFT JOIN(select Introducer_name, count(U_ID) I_Count FROM Introducer_Comm group by Introducer_name ) I_Cou ON H.CS_introducer = I_Cou.Introducer_name
                        Left join(select item_M_code, item_D_code, item_D_name Comm_Remark from Item_list  where item_M_code = 'Return' AND item_D_type = 'Y') R on H.Comm_Remark = R.item_D_code
                        LEFT JOIN(select KeyVal, Max(LogDate) LogDate from[LogTable] group by KeyVal)L on convert(varchar, H.HS_id) = KeyVal
                        WHERE H.del_tag = '0' AND H.sendcase_handle_type = 'Y' AND isnull(H.Send_amount, '')<> ''
                            AND get_amount_type = 'GTAT002'"
                    );
            try
            {
                ADOData _adoData = new ADOData(); // 測試:"Test" / 正式:""

                var parameters = new List<SqlParameter>();
                if (!string.IsNullOrEmpty(model.CS_name))
                {
                    sqlBuilder.Append(" AND A.CS_name like  @CS_name ");
                    parameters.Add(new SqlParameter("@CS_name","%" + model.CS_name + "%"));
                }

                if (!string.IsNullOrEmpty(model.Date_S) && !string.IsNullOrEmpty(model.Date_E))
                {
                    sqlBuilder.Append(" AND get_amount_date between @Date_S and @Date_E ");
                    parameters.Add(new SqlParameter("@Date_S", FuncHandler.ConvertROCToGregorian(model.Date_S)));
                    parameters.Add(new SqlParameter("@Date_E", FuncHandler.ConvertROCToGregorian(model.Date_E)));
                }

                if(!string.IsNullOrEmpty(model.OrderByStr))
                {
                    sqlBuilder.Append($" order by {model.OrderByStr}");
                }


                DataTable dtResult = _adoData.ExecuteQuery(sqlBuilder.ToString(), parameters);

                resultClass.ResultCode = "000";
                resultClass.objResult = JsonConvert.SerializeObject(dtResult);
                return Ok(resultClass);
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                resultClass.ResultMsg = $"Response error: {ex.Message}";
                return StatusCode(500, resultClass);
            }
        }

    }
}
