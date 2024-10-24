using KF_WebAPI.BaseClass;
using KF_WebAPI.BaseClass.AE;
using KF_WebAPI.FunctionHandler;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Newtonsoft.Json;
using System;
using System.Data;

namespace KF_WebAPI.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class AE_ACCController : ControllerBase
    {
        #region 應收帳款分期管理 RC_D_list.asp
        /// <summary>
        /// 呆帳清單 GetAccRcDeatilList/RC_D_daily_debt.asp
        /// </summary>
        /// <param name="Date_E">113/10/24</param>
        [HttpGet("GetAccRcDeatilList")]
        public ActionResult<ResultClass<string>> GetAccRcDeatilList(string Date_E)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            Date_E = FuncHandler.ConvertROCToGregorian(Date_E);
            var Date_S = DateTime.Parse(Date_E).AddMonths(-2).ToString("yyyy/MM/dd");

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"        
                    select Receivable_D.*,House_apply.CS_name,Receivable_M.RCM_cknum,Receivable_M.amount_total,Receivable_M.month_total
                    ,Receivable_M.amount_per_month,Receivable_M.date_begin,Receivable_M.RCM_note
                    ,(select U_name FROM User_M where U_num = Receivable_D.add_num AND del_tag='0') as add_name   
                    ,isnull((select U_name FROM User_M where U_num = Receivable_D.check_pay_num AND del_tag='0'),'') as check_pay_name
                    ,isnull((select U_name FROM User_M where U_num = Receivable_D.bad_debt_num AND del_tag='0'),'') as bad_debt_name   
                    ,isnull((select U_name FROM User_M where U_num = Receivable_D.cancel_num AND del_tag='0'),'') as cancel_name  
                    from Receivable_D  
                    LEFT JOIN Receivable_M on Receivable_M.RCM_id = Receivable_D.RCM_id AND Receivable_M.del_tag='0'  
                    LEFT JOIN House_apply on House_apply.HA_id = Receivable_M.HA_id AND House_apply.del_tag='0'  
                    where Receivable_D.del_tag = '0'  
                    AND (Receivable_D.RC_date >= @Date_S + ' 00:00:00' AND Receivable_D.RC_date <= @Date_E + ' 23:59:59')   
                    AND (Receivable_D.check_pay_type='N') AND (Receivable_D.bad_debt_type='Y') AND (Receivable_D.cancel_type='N')   
                    order by Receivable_D.RC_date";
                parameters.Add(new SqlParameter("@Date_S", Date_S));
                parameters.Add(new SqlParameter("@Date_E", Date_E));
                #endregion
                DataTable dtResult = _adoData.ExecuteQuery(T_SQL,parameters);
                if (dtResult.Rows.Count > 0)
                {
                    resultClass.ResultCode = "000";
                    resultClass.objResult = JsonConvert.SerializeObject(dtResult);
                    return Ok(resultClass);
                }
                else
                {
                    resultClass.ResultCode = "400";
                    resultClass.ResultMsg = "查無資料";
                    return BadRequest(resultClass);
                }
            }
            catch (Exception)
            {
                resultClass.ResultCode = "500";
                return StatusCode(500, resultClass);
            }
        }

        //應收帳款 - 沖銷
        //應收帳款 - 轉呆帳
        //客戶申辦資訊
        //每日未沖銷清單
        #endregion

        #region 應收帳款-催收

        #endregion

        #region 應收帳款-逾期繳款

        #endregion

        #region 應收帳款-逾放比

        #endregion

        #region 應收帳款-本金餘額比

        #endregion

        #region 資產數據分析

        #endregion

        #region 債權憑證

        #endregion
    }
}
