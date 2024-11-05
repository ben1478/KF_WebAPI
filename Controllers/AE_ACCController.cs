using KF_WebAPI.BaseClass;
using KF_WebAPI.BaseClass.AE;
using KF_WebAPI.FunctionHandler;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Microsoft.SqlServer.Server;
using Newtonsoft.Json;
using System;
using System.Data;
using System.Data.SqlTypes;
using System.Reflection;

namespace KF_WebAPI.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class AE_ACCController : ControllerBase
    {
        #region 通用
        /// <summary>
        /// 每日未沖銷清單 RC_D_daily_LQuery/RC_D_daily.asp
        /// </summary>
        /// <param name="Date_E">113/10/24</param>
        [HttpGet("RC_D_daily_LQuery")]
        public ActionResult<ResultClass<string>> RC_D_daily_LQuery(string Date_E)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            Date_E = FuncHandler.ConvertROCToGregorian(Date_E);
            var Date_S = DateTime.Now.AddMonths(-2).ToString("yyyy/MM/dd");

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"        
                    select Rd.RCM_id,Rd.RCD_id,Ha.CS_name,isnull(Rm.amount_total, 0) amount_total ,isnull(Rm.month_total, 0) month_total,
                    Rd.RC_count,FORMAT(Rd.RC_date, 'yyyy/MM/dd') as RC_date,Rd.RC_amount
                    ,Rd.check_pay_type,FORMAT(Rd.check_pay_date, 'yyyy/MM/dd') as check_pay_date
                    ,isnull((select U_name FROM User_M where U_num = Rd.check_pay_num AND del_tag = '0'),'') as check_pay_name,Rm.RCM_note   
                    ,Rd.bad_debt_type,FORMAT(Rd.bad_debt_date, 'yyyy/MM/dd') as bad_debt_date
                    ,isnull((select U_name FROM User_M where U_num = Rd.bad_debt_num AND del_tag = '0'),'') as bad_debt_name
                    from Receivable_D Rd
                    LEFT JOIN Receivable_M Rm on Rm.RCM_id = Rd.RCM_id AND Rm.del_tag= '0'
                    LEFT JOIN House_apply Ha on Ha.HA_id = Rm.HA_id AND Ha.del_tag= '0'
                    where Rd.del_tag = '0'
                    AND (Rd.RC_date >= @Date_S +' 00:00:000' AND Rd.RC_date <= @Date_E +' 23:59:59' )
                    AND(Rd.check_pay_type= 'N')
                    AND(Rd.bad_debt_type= 'N')
                    AND(Rd.cancel_type= 'N')
                    order by Rd.RC_date";
                parameters.Add(new SqlParameter("@Date_S", Date_S));
                parameters.Add(new SqlParameter("@Date_E", Date_E));
                #endregion
                DataTable dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
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
        /// <summary>
        /// 呆帳清單查詢 Rcd_Bad_LQuery/RC_D_daily_debt.asp
        /// </summary>
        /// <param name="Date_E">113/10/24</param>
        [HttpGet("Rcd_Bad_LQuery")]
        public ActionResult<ResultClass<string>> Rcd_Bad_LQuery(string Date_E)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            Date_E = FuncHandler.ConvertROCToGregorian(Date_E);
            var Date_S = DateTime.Now.AddMonths(-2).ToString("yyyy/MM/dd");

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"        
                    select Rd.RCM_id,Rd.RCD_id,Ha.CS_name,isnull(Rm.amount_total, 0) amount_total ,isnull(Rm.month_total, 0) month_total,
                    Rd.RC_count,FORMAT(Rd.RC_date, 'yyyy/MM/dd') as RC_date,Rd.RC_amount
                    ,Rd.check_pay_type,FORMAT(Rd.check_pay_date, 'yyyy/MM/dd') as check_pay_date
                    ,isnull((select U_name FROM User_M where U_num = Rd.check_pay_num AND del_tag = '0'),'') as check_pay_name  
                    ,Rm.RCM_note   
                    ,Rd.bad_debt_type,FORMAT(Rd.bad_debt_date, 'yyyy/MM/dd') as bad_debt_date
                    ,isnull((select U_name FROM User_M where U_num = Rd.bad_debt_num AND del_tag = '0'),'') as bad_debt_name 
                    from Receivable_D Rd 
                    LEFT JOIN Receivable_M Rm on Rm.RCM_id = Rd.RCM_id AND Rm.del_tag='0'  
                    LEFT JOIN House_apply Ha on Ha.HA_id = Rm.HA_id AND Ha.del_tag='0'  
                    where Rd.del_tag = '0'  
                    AND (Rd.RC_date >= @Date_S + ' 00:00:00' AND Rd.RC_date <= @Date_E + ' 23:59:59')   
                    AND (Rd.check_pay_type='N') AND (Rd.bad_debt_type='Y') AND (Rd.cancel_type='N')   
                    order by Rd.RC_date";
                parameters.Add(new SqlParameter("@Date_S", Date_S));
                parameters.Add(new SqlParameter("@Date_E", Date_E));
                #endregion
                DataTable dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
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
        /// <summary>
        /// 應收帳款分期管理查詢 - 沖銷 && 呆帳 Rcd_Check_Query/RC_D_check.asp&RC_D_debt.asp
        /// </summary>
        /// <param name="RCD_id">10021822</param> 
        [HttpGet("Rcd_Check_Query")]
        public ActionResult<ResultClass<string>> Rcd_Check_Query(string RCD_id)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"
                    select Ha.CS_name,RD.RC_count,RC_date,RC_amount,RD.check_pay_type,RD.PartiallySettled,RD.check_pay_type,RD.check_pay_date
                    ,isnull((select U_name FROM User_M where U_num = RD.check_pay_num AND del_tag = '0'),'') as check_pay_name
                    ,RD.invoice_no,RD.invoice_date,RD.RC_note,RD.bad_debt_type,RD.bad_debt_date
                    ,isnull((select U_name FROM User_M where U_num = RD.bad_debt_num AND del_tag='0'),'') as bad_debt_name 
                    from Receivable_D RD
                    LEFT JOIN Receivable_M Rm on Rm.RCM_id = RD.RCM_id AND Rm.del_tag= '0'
                    LEFT JOIN House_apply Ha on Ha.HA_id = Rm.HA_id AND Ha.del_tag= '0'
                    where RD.del_tag = '0' AND RD.RCD_id = @RCD_id order by RD.RC_date";
                parameters.Add(new SqlParameter("@RCD_id", RCD_id));
                #endregion
                DataTable dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
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
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                return StatusCode(500, resultClass);
            }
        }
        /// <summary>
        /// 應收帳款分期管理異動 - 沖銷(Type:1) && 呆帳(Type:2) Rcd_Check_Query/RC_D_check.asp&RC_D_debt.asp
        /// </summary>
        [HttpPost("Rcd_Check_SUpd")]
        public ActionResult<ResultClass<string>> Rcd_Check_SUpd(Receivable_D_check_req model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            var User_Num = HttpContext.Session.GetString("UserID");

            try
            {
                var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

                if (model.Type == "1")
                {
                    if (model.check_pay_type != "Y") //要勾選已沖銷才會改資料
                    {
                        resultClass.ResultCode = "000";
                        resultClass.ResultMsg = "沖銷異動成功";
                        return Ok(resultClass);
                    }
                    else
                    {
                        ADOData _adoData = new ADOData();
                        #region SQL
                        var parameters = new List<SqlParameter>();
                        var T_SQL = @"update Receivable_D set check_pay_type='Y',PartiallySettled=@PartiallySettled,
                        check_pay_date=@check_pay_date,check_pay_num=@check_pay_num,invoice_no=@invoice_no,
                        invoice_date=@invoice_date,RC_note=@RC_note,edit_date=GETDATE(),edit_num=@edit_num,edit_ip=@IP 
                        where del_tag = '0' AND RCD_id =@RCD_id";
                        if (model.PartiallySettled > 0)
                        {
                            parameters.Add(new SqlParameter("@PartiallySettled", model.PartiallySettled));
                        }
                        else
                        {
                            parameters.Add(new SqlParameter("@PartiallySettled", 0));
                        }
                        if (model.str_check_pay_date != null)
                        {
                            var Date = DateTime.Parse(FuncHandler.ConvertROCToGregorian(model.str_check_pay_date));
                            parameters.Add(new SqlParameter("@check_pay_date", Date));
                        }
                        else
                        {
                            parameters.Add(new SqlParameter("@check_pay_date", DBNull.Value));
                        }
                        parameters.Add(new SqlParameter("@check_pay_num", User_Num));
                        if (model.invoice_no != null)
                        {
                            parameters.Add(new SqlParameter("@invoice_no", model.invoice_no));
                        }
                        else
                        {
                            parameters.Add(new SqlParameter("@invoice_no", DBNull.Value));
                        }
                        if (model.str_invoice_date != null)
                        {
                            var Date = DateTime.Parse(FuncHandler.ConvertROCToGregorian(model.str_invoice_date));
                            parameters.Add(new SqlParameter("@invoice_date", Date));
                        }
                        else
                        {
                            parameters.Add(new SqlParameter("@invoice_date", DBNull.Value));
                        }
                        parameters.Add(new SqlParameter("@edit_num", User_Num));
                        parameters.Add(new SqlParameter("@IP", clientIp));
                        parameters.Add(new SqlParameter("@RCD_id", model.RCD_id));
                        parameters.Add(new SqlParameter("@RC_note", model.RC_note));
                        #endregion

                        int result = _adoData.ExecuteNonQuery(T_SQL, parameters);
                        if (result == 0)
                        {
                            resultClass.ResultCode = "400";
                            resultClass.ResultMsg = "異動失敗";
                            return BadRequest(resultClass);
                        }
                        else
                        {
                            resultClass.ResultCode = "000";
                            resultClass.ResultMsg = "異動成功";
                            return Ok(resultClass);
                        }
                    }
                }
                else
                {
                    ADOData _adoData = new ADOData();
                    #region SQL
                    var parameters = new List<SqlParameter>();
                    var T_SQL = @"
                        update Receivable_D set bad_debt_type=@bad_debt_type,bad_debt_date=@bad_debt_date,bad_debt_num=@bad_debt_num
                        ,RC_note=@RC_note,edit_date=GETDATE(),edit_num=@edit_num,edit_ip=@IP 
                        where del_tag = '0' AND RCD_id =@RCD_id";
                    if (model.bad_debt_type == "Y")
                    {
                        var Date = DateTime.Parse(FuncHandler.ConvertROCToGregorian(model.str_bad_debt_date));
                        parameters.Add(new SqlParameter("@bad_debt_date", Date));
                        parameters.Add(new SqlParameter("@bad_debt_num", User_Num));
                    }
                    else
                    {
                        parameters.Add(new SqlParameter("@bad_debt_date", DBNull.Value));
                        parameters.Add(new SqlParameter("@bad_debt_num", ""));
                    }
                    parameters.Add(new SqlParameter("@bad_debt_type", model.bad_debt_type));
                    parameters.Add(new SqlParameter("@RC_note", model.RC_note));
                    parameters.Add(new SqlParameter("@edit_num", User_Num));
                    parameters.Add(new SqlParameter("@IP", clientIp));
                    parameters.Add(new SqlParameter("@RCD_id", model.RCD_id));
                    #endregion
                    int result = _adoData.ExecuteNonQuery(T_SQL, parameters);
                    if (result == 0)
                    {
                        resultClass.ResultCode = "400";
                        resultClass.ResultMsg = "異動失敗";
                        return BadRequest(resultClass);
                    }
                    else
                    {
                        resultClass.ResultCode = "000";
                        resultClass.ResultMsg = "異動成功";
                        return Ok(resultClass);
                    }
                }

            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                return StatusCode(500, resultClass); // 返回 500 錯誤碼
            }
        }
        /// <summary>
        /// 提供分期+案件資料 Rc_Deatil_M_SQuery/RC_M_detail.asp
        /// </summary>
        /// <param name="RCM_id">10020598</param>
        [HttpGet("Rc_Deatil_M_SQuery")]
        public ActionResult<ResultClass<string>> Rc_Deatil_M_SQuery(string RCM_id)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"
                    select Ha.CS_name,Ha.CS_birthday,Ha.CS_MTEL1,Ha.CS_register_address,Ha.CS_company_name,Ha.CS_company_address,Ha.CS_company_tel,
                    Li.item_D_name as CS_job_kind_name,Ha.CS_job_title,Ha.CS_job_years,Li_1.item_D_name as CS_income_way_name,Ha.CS_rental,
                    Ha.CS_income_everymonth,Rm.HS_id,Hs.get_amount,Hs.interest_rate_pass,Rm.amount_total,Rm.loan_grace_num,Rm.month_total,
                    Rm.amount_per_month,Hs.get_amount_date,Rm.date_begin,Rm.RCM_note,RM.court_sale 
                    from Receivable_M Rm
                    left join House_apply Ha on Ha.HA_id = Rm.HA_id and ha.del_tag='0'
                    left join House_sendcase Hs on Hs.HS_id = Rm.HS_id AND Hs.del_tag= '0'
                    left join Item_list Li on Li.item_M_code='job_kind' and Li.item_D_type='Y' and Li.del_tag='0' and Li.item_D_code=Ha.CS_job_kind
                    left join Item_list Li_1 on Li_1.item_M_code='income_way' and Li_1.item_D_type='Y' and Li_1.del_tag='0' and Li_1.item_D_code=Ha.CS_income_way
                    where Rm.del_tag = '0' AND Rm.RCM_id = @RCM_id ";
                parameters.Add(new SqlParameter("@RCM_id", RCM_id));
                #endregion
                DataTable dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
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
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                return StatusCode(500, resultClass);
            }
        }
        /// <summary>
        /// 提供分期明細資料 Rc_Deatil_D_LQuery/RC_M_detail.asp
        /// </summary>
        /// <param name="RCM_id">10020598</param>
        [HttpGet("Rc_Deatil_D_LQuery")]
        public ActionResult<ResultClass<string>> Rc_Deatil_D_LQuery(string RCM_id)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"
                    select Rd.RC_count,Rd.RC_date,Rd.RC_amount,Rd.check_pay_type,Rd.check_pay_date,
                    isnull((select U_name FROM User_M where U_num = Rd.check_pay_num AND del_tag = '0'),'') as check_pay_name,
                    Rd.bad_debt_type,Rd.bad_debt_date,
                    isnull((select U_name FROM User_M where U_num = Rd.bad_debt_num AND del_tag = '0'),'') as bad_debt_name,
                    Rd.cancel_type,Rd.cancel_date,
                    isnull((select U_name FROM User_M where U_num = Rd.cancel_num AND del_tag = '0'),'') as cancel_name
                    from Receivable_D Rd
                    LEFT JOIN Receivable_M Rm on Rm.RCM_id = Rd.RCM_id AND Rm.del_tag= '0'  
                    LEFT JOIN House_apply Ha on Ha.HA_id = Rm.HA_id AND Ha.del_tag= '0'  
                    where Rd.del_tag = '0' AND Rd.RCM_id = @RCM_id order by Rd.RC_date";
                parameters.Add(new SqlParameter("@RCM_id", RCM_id));
                #endregion
                DataTable dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
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
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                return StatusCode(500, resultClass);
            }
        }
        /// <summary>
        /// 應收帳款分期管理內容修改 Rc_Deatil_MD_Upd/RC_M_detail.asp
        /// </summary>
        [HttpPost("Rc_Deatil_MD_Upd")]
        public ActionResult<ResultClass<string>> Rc_Deatil_MD_Upd(Receivable_MD_req model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            var User_Num = HttpContext.Session.GetString("UserID");
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

            try
            {
                var Date_Begin = DateTime.Parse(FuncHandler.ConvertROCToGregorian(model.str_date_begin)).AddHours(6);
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters_M = new List<SqlParameter>();
                var T_SQL_M = @"Update Receivable_M set amount_per_month=@amount_per_month,date_begin=@date_begin,RCM_note=@RCM_note, 
                    edit_num=@edit_num,edit_date=GETDATE(),edit_ip=@edit_ip where RCM_id=@RCM_id";
                parameters_M.Add(new SqlParameter("@amount_per_month", model.amount_per_month));
                parameters_M.Add(new SqlParameter("@RCM_id", model.RCM_id));
                parameters_M.Add(new SqlParameter("date_begin", Date_Begin));
                parameters_M.Add(new SqlParameter("@RCM_note", model.RCM_note));
                parameters_M.Add(new SqlParameter("@edit_num", User_Num));
                parameters_M.Add(new SqlParameter("@edit_ip", clientIp));
                #endregion
                int result = _adoData.ExecuteNonQuery(T_SQL_M, parameters_M);
                if (result == 0)
                {
                    resultClass.ResultCode = "400";
                    resultClass.ResultMsg = "主檔異動失敗";
                    return BadRequest(resultClass);
                }
                else
                {
                    #region SQL 查詢 Receivable_D
                    var parameters = new List<SqlParameter>();
                    var T_SQL = @"select * from Receivable_D where RCM_id=@RCM_id order by RC_count";
                    parameters.Add(new SqlParameter("@RCM_id", model.RCM_id));
                    #endregion
                    var dtResult = _adoData.ExecuteQuery(T_SQL, parameters).AsEnumerable();
                    int count = 0;
                    int result_d = 0;
                    foreach (DataRow dr in dtResult)
                    {
                        var parameters_d = new List<SqlParameter>();
                        var T_SQL_D = @"Update Receivable_D set RC_amount=@RC_amount,RC_date=@RC_date where RCM_id=@RCM_id and RC_count=@RC_count";
                        parameters_d.Add(new SqlParameter("@RC_amount", model.amount_per_month));
                        parameters_d.Add(new SqlParameter("@RC_date", Date_Begin.AddMonths(count)));
                        parameters_d.Add(new SqlParameter("@RCM_id", model.RCM_id));
                        parameters_d.Add(new SqlParameter("@RC_count", dr["RC_count"]));
                        result_d = _adoData.ExecuteNonQuery(T_SQL_D, parameters_d);
                        count++;
                        if (result_d == 0)
                            break;
                    }
                    if (result_d == 0)
                    {
                        resultClass.ResultCode = "400";
                        resultClass.ResultMsg = "異動失敗";
                        return BadRequest(resultClass);
                    }
                    else
                    {
                        resultClass.ResultCode = "000";
                        resultClass.ResultMsg = "異動成功";
                        return Ok(resultClass);
                    }
                }
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                return StatusCode(500, resultClass);
            }
        }
        /// <summary>
        /// 提前還款 Rc_Deatil_Settle/RC_M_settleDB.asp
        /// </summary>
        [HttpPost("Rc_Deatil_Settle")]
        public ActionResult<ResultClass<string>> Rc_Deatil_Settle(Receivable_Settle_req model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            var User_Num = HttpContext.Session.GetString("UserID");
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

            try
            {
                var date_Begin = DateTime.Parse(FuncHandler.ConvertROCToGregorian(model.str_begin_settle)).AddHours(6);

                ADOData _adoData = new ADOData();
                #region SQL_Receivable_M
                var parameters = new List<SqlParameter>();
                var T_SQL = @"select * from Receivable_M where del_tag = '0' AND RCM_id = @RCM_id";
                parameters.Add(new SqlParameter("@RCM_id", model.RCM_id));
                #endregion
                DataTable dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
                if (dtResult.Rows.Count > 0)
                {
                    //異動Receivable_M
                    #region SQL_Receivable_M
                    var parameters_m = new List<SqlParameter>();
                    var T_SQL_M = @"Update Receivable_M set RCM_note=@RCM_note,court_sale=@court_sale,edit_date=GETDATE(),edit_num=@edit_num
                        ,edit_ip=@edit_ip where del_tag = '0' AND RCM_id = @RCM_id";
                    var str_note = "[於" + date_Begin + " " + model.RCM_note + "******" + dtResult.Rows[0]["RCM_note"].ToString();
                    parameters_m.Add(new SqlParameter("@RCM_note", str_note));
                    if (!string.IsNullOrEmpty(model.court_sale))
                    {
                        parameters_m.Add(new SqlParameter("@court_sale", model.court_sale));
                    }
                    else
                    {
                        parameters_m.Add(new SqlParameter("@court_sale", DBNull.Value));
                    }
                    parameters_m.Add(new SqlParameter("@edit_num", User_Num));
                    parameters_m.Add(new SqlParameter("@edit_ip", clientIp));
                    parameters_m.Add(new SqlParameter("@RCM_id", model.RCM_id));
                    #endregion
                    int result_m = _adoData.ExecuteNonQuery(T_SQL_M, parameters_m);
                    if (result_m == 0)
                    {
                        resultClass.ResultCode = "400";
                        resultClass.ResultMsg = "主檔異動失敗";
                        return BadRequest(resultClass);
                    }
                    else
                    {
                        //確認是否有小於結清日的未沖銷資料
                        #region SQL_Receivable_D
                        var parameters_check = new List<SqlParameter>();
                        var T_SQL_CHECK = @"select * from Receivable_D where del_tag = '0' AND RCM_id =@RCM_id 
                            and CONVERT(varchar(100),Cast(RC_date as datetime), 23) <= CONVERT(varchar(100),Cast(@date_Begin as datetime), 23)
                            and RemainingPrincipal is null order by RC_count";
                        parameters_check.Add(new SqlParameter("@RCM_id", model.RCM_id));
                        parameters_check.Add(new SqlParameter("@date_Begin", date_Begin));
                        #endregion
                        DataTable dtResultCheck = _adoData.ExecuteQuery(T_SQL_CHECK, parameters_check);
                        decimal RemainingPrincipal = 0;
                        //處理小於等於結清日的未沖銷資料
                        if (dtResultCheck.Rows.Count > 0)
                        {
                            //抓已沖銷的最大本金餘額
                            #region SQL
                            var parameters_re = new List<SqlParameter>();
                            var T_SQL_RE = @"select RemainingPrincipal from Receivable_D where RCM_id = @RCM_id AND
                                RC_count =  (select MAX(RC_count) RC_count from Receivable_D D 
                                where del_tag = '0' AND RCM_id = @RCM_id AND RemainingPrincipal is not null )";
                            parameters_re.Add(new SqlParameter("@RCM_id", model.RCM_id));
                            #endregion
                            DataTable dtResult_re = _adoData.ExecuteQuery(T_SQL_RE, parameters_re);
                            RemainingPrincipal = Convert.ToDecimal(dtResult_re.Rows[0]["RemainingPrincipal"]);
                            int result_d = 0;
                            foreach (DataRow dr in dtResultCheck.Rows)
                            {
                                var parameters_d = new List<SqlParameter>();
                                var T_SQL_D = @"Update Receivable_D set check_pay_type='Y',RC_note=@RC_note,RemainingPrincipal=@RemainingPrincipal 
                                    ,check_pay_num=@check_pay_num,check_pay_date=GETDATE(),edit_date=GETDATE(),edit_num=@edit_num,edit_ip=@edit_ip
                                    where RCM_id=@RCM_id and RCD_id=@RCD_id";
                                parameters_d.Add(new SqlParameter("@RC_note", model.RCM_note));
                                RemainingPrincipal = RemainingPrincipal - Convert.ToDecimal(dr["Rmoney"]);
                                parameters_d.Add(new SqlParameter("@RemainingPrincipal", RemainingPrincipal));
                                parameters_d.Add(new SqlParameter("@RCM_id", model.RCM_id));
                                parameters_d.Add(new SqlParameter("@RCD_id", Convert.ToDecimal(dr["RCD_id"])));
                                parameters_d.Add(new SqlParameter("@check_pay_num", User_Num));
                                parameters_d.Add(new SqlParameter("@edit_num", User_Num));
                                parameters_d.Add(new SqlParameter("@edit_ip", clientIp));
                                result_d = _adoData.ExecuteNonQuery(T_SQL_D, parameters_d);
                                if (result_d == 0)
                                    break;
                            }
                            if (result_d == 0)
                            {
                                resultClass.ResultCode = "400";
                                resultClass.ResultMsg = "提前結清明細失敗";
                                return BadRequest(resultClass);
                            }
                        }
                        //處理大於結清日的未沖銷資料
                        #region SQL
                        var parameters_non = new List<SqlParameter>();
                        var T_SQL_NON = @"Update Receivable_D set check_pay_type='S',RC_note=@RC_note,edit_date=GETDATE(),edit_num=@edit_num,edit_ip=@edit_ip 
                            where del_tag = '0' AND RCM_id =@RCM_id 
                            AND  CONVERT(varchar(100),Cast(RC_date as datetime), 23) > CONVERT(varchar(100),Cast(@date_Begin as datetime), 23) 
                            and RemainingPrincipal is null";
                        parameters_non.Add(new SqlParameter("@RC_note", model.RCM_note));
                        parameters_non.Add(new SqlParameter("@edit_num", User_Num));
                        parameters_non.Add(new SqlParameter("@edit_ip", clientIp));
                        parameters_non.Add(new SqlParameter("@RCM_id", model.RCM_id));
                        parameters_non.Add(new SqlParameter("@date_Begin", date_Begin));
                        #endregion
                        int result_non = _adoData.ExecuteNonQuery(T_SQL_NON, parameters_non);
                        if (result_non == 0)
                        {
                            resultClass.ResultCode = "400";
                            resultClass.ResultMsg = "提前結清明細失敗";
                            return BadRequest(resultClass);
                        }
                        else
                        {
                            resultClass.ResultCode = "000";
                            resultClass.ResultMsg = "結清成功";
                            return Ok(resultClass);
                        }
                    }
                }
                else
                {
                    resultClass.ResultCode = "400";
                    resultClass.ResultMsg = "查無資料";
                    return BadRequest(resultClass);
                }

            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                return StatusCode(500, resultClass);
            }
        }
        /// <summary>
        /// 取得是否有重新設定的權限
        /// </summary>
        [HttpGet("GetShowResetUser")]
        public ActionResult<ResultClass<string>> GetShowResetUser()
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var User_Num = HttpContext.Session.GetString("UserID");

            try
            {
                string[] str = new string[] { "7005" };
                SpecialClass specialClass = FuncHandler.CheckSpecial(str, User_Num);
                resultClass.objResult = JsonConvert.SerializeObject(specialClass);
                return Ok(resultClass);
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                return StatusCode(500, resultClass);
            }
        }
        /// <summary>
        /// 應收帳款分期管理內容重新設定 Rc_Deatil_Reset/RC_M_set.asp
        /// </summary>
        /// <param name="RCM_id">10020815</param>
        [HttpGet("Rc_Deatil_Reset")]
        public ActionResult<ResultClass<string>> Rc_Deatil_Reset(string RCM_id)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            var User_Num = HttpContext.Session.GetString("UserID");
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

            try
            {
                ADOData _adoData = new ADOData();
                //檢測是否有沖銷或轉呆
                var parameters_c = new List<SqlParameter>();
                var T_SQL_C = @"select * from Receivable_D where RCM_id=@RCM_id and (check_pay_type='Y' OR bad_debt_type='Y')";
                parameters_c.Add(new SqlParameter("@RCM_id", RCM_id));
                DataTable dtResult_c = _adoData.ExecuteQuery(T_SQL_C, parameters_c);
                if (dtResult_c.Rows.Count > 0)
                {
                    resultClass.ResultCode = "400";
                    resultClass.ResultMsg = "已有沖銷或轉呆,無法重新設定";
                    return BadRequest(resultClass);
                }
                else
                {
                    var parameters = new List<SqlParameter>();
                    var T_SQL = @"Update Receivable_M set del_tag='1',del_num=@del_num,del_date=GETDATE(),del_ip=@del_ip where RCM_id=@RCM_id;
                        Update Receivable_D set del_tag='1',del_num=@del_num,del_date=GETDATE(),del_ip=@del_ip where RCM_id=@RCM_id";
                    parameters.Add(new SqlParameter("@RCM_id", RCM_id));
                    parameters.Add(new SqlParameter("@del_num", User_Num));
                    parameters.Add(new SqlParameter("@del_ip", clientIp));
                    int result = _adoData.ExecuteNonQuery(T_SQL, parameters);
                    if (result == 0)
                    {
                        resultClass.ResultCode = "400";
                        resultClass.ResultMsg = "重新設定失敗";
                        return BadRequest(resultClass);
                    }
                    else
                    {
                        resultClass.ResultCode = "000";
                        resultClass.ResultMsg = "重新設定成功";
                        return Ok(resultClass);
                    }
                }
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                return StatusCode(500, resultClass);
            }
        }
        /// <summary>
        /// 應收帳款分期管理內容新增設定 Rc_Deatil_MD_Ins/RC_M_set.asp
        /// </summary>
        [HttpPost("Rc_Deatil_MD_Ins")]
        public ActionResult<ResultClass<string>> Rc_Deatil_MD_Ins(Receivable_M model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            var User_Num = HttpContext.Session.GetString("UserID");
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();
            var Rem_Fee = 20; //匯費20塊

            try
            {
                var Fun = new FuncHandler();
                var cknum = Fun.GetCheckNum();
                var date_Begin = DateTime.Parse(FuncHandler.ConvertROCToGregorian(model.str_date_begin)).AddHours(6);
                ADOData _adoData = new ADOData();
                #region SQL_Rc_Deatil_M
                var parameters_m = new List<SqlParameter>();
                var T_SQL_M = @"Insert into Receivable_M (RCM_cknum,HS_id,HA_id,amount_total,month_total,amount_per_month,date_begin,RCM_note
                    ,loan_grace_num,cust_rate,cust_amount_per_month,add_date,add_num,add_ip) 
                    Values (@RCM_cknum,@HS_id,@HA_id,@amount_total,@month_total,@amount_per_month,@date_begin,@RCM_note,@loan_grace_num,@cust_rate
                    ,@cust_amount_per_month,GETDATE(),@add_num,@add_ip)";
                parameters_m.Add(new SqlParameter("@RCM_cknum", cknum));
                parameters_m.Add(new SqlParameter("@HS_id", model.HS_id));
                parameters_m.Add(new SqlParameter("@HA_id", model.HA_id));
                parameters_m.Add(new SqlParameter("@amount_total", model.amount_total));
                parameters_m.Add(new SqlParameter("@month_total", model.month_total));
                parameters_m.Add(new SqlParameter("@amount_per_month", model.amount_per_month));
                parameters_m.Add(new SqlParameter("@date_begin", date_Begin));
                parameters_m.Add(new SqlParameter("@RCM_note", model.RCM_note));
                parameters_m.Add(new SqlParameter("@loan_grace_num", model.loan_grace_num));
                if (!string.IsNullOrEmpty(model.cust_rate))
                {
                    parameters_m.Add(new SqlParameter("@cust_rate", model.cust_rate));
                }
                else
                {
                    parameters_m.Add(new SqlParameter("@cust_rate", DBNull.Value));
                }
                if (model.cust_amount_per_month != null)
                {
                    parameters_m.Add(new SqlParameter("@cust_amount_per_month", model.cust_amount_per_month));
                }
                else
                {
                    parameters_m.Add(new SqlParameter("@cust_amount_per_month", DBNull.Value));
                }
                parameters_m.Add(new SqlParameter("@add_num", User_Num));
                parameters_m.Add(new SqlParameter("@add_ip", clientIp));
                #endregion
                int result_m = _adoData.ExecuteNonQuery(T_SQL_M, parameters_m);
                if (result_m == 0)
                {
                    resultClass.ResultCode = "400";
                    resultClass.ResultMsg = "主檔新增異動失敗";
                    return BadRequest(resultClass);
                }
                else
                {
                    var parameters = new List<SqlParameter>();
                    var T_SQL = @"select top 1 * from Receivable_M where del_tag = '0' AND HA_id =@HA_id and HS_id=@HS_id and RCM_cknum=@RCM_cknum";
                    parameters.Add(new SqlParameter("@HA_id", model.HA_id));
                    parameters.Add(new SqlParameter("@HS_id", model.HS_id));
                    parameters.Add(new SqlParameter("@RCM_cknum", cknum));
                    DataTable dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
                    if (dtResult.Rows.Count > 0)
                    {
                        int result = 0;
                        var maxCount = Convert.ToInt32(dtResult.Rows[0]["month_total"]);
                        var RCM_id = Convert.ToInt32(dtResult.Rows[0]["RCM_id"]);
                        decimal pri_bal = (decimal)model.amount_total;
                        for (int i = 1; i <= maxCount; i++)
                        {
                            var cknum_d = Fun.GetCheckNum();
                            var parameters_d = new List<SqlParameter>();
                            var T_SQL_D = @"Insert into Receivable_D (RCD_cknum,RCM_id,RC_count,RC_amount,RC_date,RC_type,RC_note,interest
                                ,Rmoney,add_date,add_num,add_ip) 
                                Values (@RCD_cknum,@RCM_id,@RC_count,@RC_amount,@RC_date,'','',@interest,@Rmoney,GETDATE(),@add_num,@add_ip)";
                            parameters_d.Add(new SqlParameter("@RCD_cknum", cknum_d));
                            parameters_d.Add(new SqlParameter("@RCM_id", RCM_id));
                            parameters_d.Add(new SqlParameter("@RC_count", i));
                            parameters_d.Add(new SqlParameter("@RC_amount", model.amount_per_month));
                            parameters_d.Add(new SqlParameter("@RC_date", date_Begin.AddMonths(i - 1)));
                            if (model.loan_grace_num > 0)
                            {
                                if (model.cust_amount_per_month != null)
                                {
                                    parameters_d.Add(new SqlParameter("@interest", model.cust_amount_per_month));
                                }
                                else
                                {
                                    parameters_d.Add(new SqlParameter("@interest", DBNull.Value));
                                }
                                parameters_d.Add(new SqlParameter("@Rmoney", 0));
                            }
                            else
                            {
                                var interest = Math.Round(pri_bal * ((Convert.ToDecimal(model.interest_rate_pass) / 100) / 12), 0);
                                decimal rmoney = Convert.ToDecimal((model.amount_per_month - Rem_Fee) - interest);
                                parameters_d.Add(new SqlParameter("@interest", interest));
                                parameters_d.Add(new SqlParameter("@Rmoney", rmoney));
                                pri_bal = pri_bal - rmoney;
                            }
                            parameters_d.Add(new SqlParameter("@add_num", User_Num));
                            parameters_d.Add(new SqlParameter("@add_ip", clientIp));

                            result = _adoData.ExecuteNonQuery(T_SQL_D, parameters_d);
                            if (result == 0)
                                break;
                        }

                        if (result == 0)
                        {
                            resultClass.ResultCode = "400";
                            resultClass.ResultMsg = "設定失敗";
                            return BadRequest(resultClass);
                        }
                        else
                        {
                            resultClass.ResultCode = "000";
                            resultClass.ResultMsg = "設定成功";
                            return Ok(resultClass);
                        }
                    }
                    else
                    {
                        resultClass.ResultCode = "400";
                        resultClass.ResultMsg = "主檔查詢失敗";
                        return BadRequest(resultClass);
                    }
                }
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                return StatusCode(500, resultClass);
            }
        }
        #endregion

        #region 應收帳款分期管理
        /// <summary>
        /// 應收帳款分期管理清單查詢 RC_D_LQuery/RC_D_list.asp
        /// </summary>
        [HttpPost("RC_D_LQuery")]
        public ActionResult<ResultClass<string>> RC_D_LQuery(Receivable_req model) 
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                model.str_Date_S = FuncHandler.ConvertROCToGregorian(model.str_Date_S);
                model.str_Date_E = FuncHandler.ConvertROCToGregorian(model.str_Date_E);
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"SELECT Rd.RCM_id,Rd.RCD_id,Ha.CS_name,FORMAT(Rm.amount_total, 'N0') AS str_amount_total,Rm.month_total,Rd.RC_count,
                    FORMAT(Rd.RC_date,'yyyy/MM/dd') AS RC_date,FORMAT(Rd.RC_amount, 'N0') AS str_RC_amount,FORMAT(Rd.interest, 'N0') AS str_interest,
                    FORMAT(Rd.Rmoney, 'N0') AS str_Rmoney,
                    FORMAT((Rm.amount_total - COALESCE((SELECT SUM(Rmoney) FROM Receivable_D WHERE RCM_ID = Rm.RCM_ID AND RC_count <= Rd.RC_count AND del_tag = '0'), 0)),'N0') AS str_RemainingAmount,
                    FORMAT(Rd.PartiallySettled, 'N0') AS str_PartiallySettled,
                    case when RC_date<SYSDATETIME() and check_pay_date is null then isnull(DATEDIFF(DAY, RC_date, SYSDATETIME()),0) else isnull(DATEDIFF(DAY, RC_date, check_pay_date),0) end DelayDay,
                    Rd.check_pay_type,FORMAT(Rd.check_pay_date, 'yyyy/MM/dd') as check_pay_date,
                    isnull((select U_name FROM User_M where U_num = Rd.check_pay_num AND del_tag = '0'),'') as check_pay_name,
                    Rd.RC_note,Rd.bad_debt_type,FORMAT(Rd.bad_debt_date, 'yyyy/MM/dd') as bad_debt_date,
                    isnull((select U_name FROM User_M where U_num = Rd.bad_debt_num AND del_tag = '0'),'') as bad_debt_name,Rd.invoice_no,
                    FORMAT(Rd.invoice_date, 'yyyy/MM/dd') as invoice_date,
                    interest_rate_pass,case when RC_date<SYSDATETIME() and check_pay_type = 'Y' then 'Y' else 'N' end isOver_RC,
                    case when Hs.sendcase_handle_date >='2023-04-24 00:00:00.000' then 'Y' else 'N' end isNewFun, Rd.RemainingPrincipal,
                    case when DATEDIFF(day, RC_date, check_pay_date) > 0 then CEILING(RC_amount*0.2/365*DATEDIFF(day, RC_date, check_pay_date)) else 0 end Fee,
                    case when DATEDIFF(day, RC_date, check_pay_date) > 0 then convert(decimal(12,6),convert(decimal(12,6),convert(decimal(4,2),convert(decimal(4,2),interest_rate_pass)/100)*5/10000*DATEDIFF(day, RC_date, check_pay_date))) else 1 end EXrate
                    FROM Receivable_D Rd
                    LEFT JOIN Receivable_M Rm ON Rm.RCM_id = Rd.RCM_id AND Rm.del_tag = '0'  
                    LEFT JOIN House_apply Ha ON Ha.HA_id = Rm.HA_id AND Ha.del_tag = '0'
                    LEFT JOIN (SELECT U_num, U_BC FROM User_M) Um ON Um.U_num = Ha.plan_num
                    LEFT JOIN House_sendcase Hs ON Hs.HS_id = Rm.HS_id AND Hs.del_tag = '0'
                    WHERE Rd.del_tag = '0' and Um.U_BC IN ('zz', 'BC0100', 'BC0200', 'BC0600', 'BC0900', 'BC0700', 'BC0800', 'BC0300', 'BC0500', 'BC0400', 'BC0800')
                    and Rd.cancel_type = 'N'                    
                    and Rd.RC_date >= @Date_S + ' 00:00:00' AND Rd.RC_date <= @Date_E + ' 23:59:59'";

                parameters.Add(new SqlParameter("@Date_S", model.str_Date_S));
                parameters.Add(new SqlParameter("@Date_E", model.str_Date_E));
                if (!string.IsNullOrEmpty(model.name))
                {
                    T_SQL = T_SQL + " and Ha.CS_name = @CS_name";
                    parameters.Add(new SqlParameter("@CS_name",model.name));
                }
                if(model.check_type != "A")
                {
                    T_SQL = T_SQL + " and Rd.check_pay_type = @check_pay_type";
                    parameters.Add(new SqlParameter("@check_pay_type", model.check_type));
                }
                if (!string.IsNullOrEmpty(model.NS_type))
                {
                    T_SQL = T_SQL + "  AND Rd.RCM_id not in (select distinct RCM_id from Receivable_D where check_pay_type='S')";
                }
                if (model.bad_type != "A")
                {
                    T_SQL = T_SQL + " and Rd.bad_debt_type = @bad_debt_type";
                    parameters.Add(new SqlParameter("@bad_debt_type", model.bad_type));
                }
                if (!string.IsNullOrEmpty(model.RC_count))
                {
                    T_SQL = T_SQL + " and Rd.RC_count = @RC_count";
                    parameters.Add(new SqlParameter("@RC_count", model.RC_count));
                }
                T_SQL = T_SQL + " ORDER BY Rd.RC_date";
                #endregion
                var result = _adoData.ExecuteQuery(T_SQL, parameters).AsEnumerable().Select(row=>new Receivable_res
                {
                    RCM_id = row.Field<decimal>("RCM_id"),
                    RCD_id = row.Field<decimal>("RCD_id"),
                    CS_name = row.Field<string>("CS_name"),
                    str_amount_total = row.Field<string>("str_amount_total"),
                    RC_count = row.Field<int>("RC_count"),
                    month_total = row.Field<int>("month_total"),
                    RC_date = FuncHandler.ConvertGregorianToROC(row.Field<string>("RC_date")),
                    str_RC_amount = row.Field<string>("str_RC_amount"),
                    str_interest = row.Field<string>("str_interest"),
                    str_Rmoney = row.Field<string>("str_Rmoney"),
                    str_RemainingAmount = row.Field<string>("str_RemainingAmount"),
                    str_PartiallySettled = row.Field<string>("str_PartiallySettled"),
                    DelayDay = row.Field<int>("DelayDay"),
                    check_pay_type = row.Field<string>("check_pay_type"),
                    check_pay_date = row.Field<string>("check_pay_date"),
                    check_pay_name = row.Field<string>("check_pay_name"),
                    RC_note = row.Field<string>("RC_note"),
                    bad_debt_type = row.Field<string>("bad_debt_type"),
                    bad_debt_date = row.Field<string>("bad_debt_date"),
                    bad_debt_name = row.Field<string>("bad_debt_name"),
                    invoice_no = row.Field<string>("invoice_no"),
                    invoice_date = row.Field<string>("invoice_date"),
                    interest_rate_pass = row.Field<string>("interest_rate_pass"),
                    isOver_RC = row.Field<string>("isOver_RC"),
                    isNewFun = row.Field<string>("isNewFun"),
                    RemainingPrincipal = row.Field<decimal?>("RemainingPrincipal"),
                    Fee = row.Field<decimal?>("Fee"),
                    EXrate = row.Field<decimal?>("EXrate")
                }).ToList();
                foreach (var item in result)
                {
                    //是否呈現延滯天數 延滯利息
                    if (item.isOver_RC == "Y")
                    {
                        int Fee = 0;

                        if (item.DelayDay >= 3 && item.DelayDay <= 6)
                            Fee = 100;
                        if (item.DelayDay >= 7 && item.DelayDay <= 14)
                            Fee = 200;
                        if (item.DelayDay > 15)
                            Fee = 300;

                        if (item.isNewFun == "Y")
                        {
                            decimal rcAmount = Convert.ToDecimal(item.str_RC_amount.Replace(",", ""));
                            item.Delaymoney = Math.Ceiling(Convert.ToDecimal(Fee + Math.Ceiling(Convert.ToDouble(rcAmount)) * 0.16 / 365 * item.DelayDay));
                        }
                        else
                        {
                            decimal remainingPrincipal = Convert.ToDecimal(item.RemainingPrincipal);
                            decimal rmoney = Convert.ToDecimal(item.str_Rmoney);
                            decimal exrate = Convert.ToDecimal(item.EXrate);
                            item.Delaymoney = Math.Ceiling(Convert.ToDecimal((Fee + item.Fee + Math.Ceiling((Math.Ceiling(remainingPrincipal) + Math.Ceiling(rmoney)) * exrate))));
                        }
                    }
                    else
                    {
                        item.DelayDay = null;
                        item.Delaymoney = null;
                    }

                    if (item.DelayDay == 0)
                        item.Delaymoney = 0;
                        
                }

                resultClass.ResultCode = "000";
                resultClass.objResult = JsonConvert.SerializeObject(result);
                return Ok(resultClass);
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                return StatusCode(500, resultClass);
            }
        }
        /// <summary>
        /// 應收帳款分期管理清單查詢Excel下載 RC_daily_Excel/RC_D_list.asp
        /// </summary>
        [HttpPost("RC_daily_Excel")]
        public IActionResult RC_daily_Excel(Receivable_req model)
        {
            try
            {
                model.str_Date_S = FuncHandler.ConvertROCToGregorian(model.str_Date_S);
                model.str_Date_E = FuncHandler.ConvertROCToGregorian(model.str_Date_E);
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"SELECT Rd.RCM_id,Rd.RCD_id,Ha.CS_name,FORMAT(Rm.amount_total, 'N0') AS str_amount_total,Rm.month_total,Rd.RC_count,
                    FORMAT(Rd.RC_date,'yyyy/MM/dd') AS RC_date,FORMAT(Rd.RC_amount, 'N0') AS str_RC_amount,FORMAT(Rd.interest, 'N0') AS str_interest,
                    FORMAT(Rd.Rmoney, 'N0') AS str_Rmoney,
                    FORMAT((Rm.amount_total - COALESCE((SELECT SUM(Rmoney) FROM Receivable_D WHERE RCM_ID = Rm.RCM_ID AND RC_count <= Rd.RC_count AND del_tag = '0'), 0)),'N0') AS str_RemainingAmount,
                    FORMAT(Rd.PartiallySettled, 'N0') AS str_PartiallySettled,
                    case when RC_date<SYSDATETIME() and check_pay_date is null then isnull(DATEDIFF(DAY, RC_date, SYSDATETIME()),0) else isnull(DATEDIFF(DAY, RC_date, check_pay_date),0) end DelayDay,
                    Rd.check_pay_type,FORMAT(Rd.check_pay_date, 'yyyy/MM/dd') as check_pay_date,
                    isnull((select U_name FROM User_M where U_num = Rd.check_pay_num AND del_tag = '0'),'') as check_pay_name,
                    Rd.RC_note,Rd.bad_debt_type,FORMAT(Rd.bad_debt_date, 'yyyy/MM/dd') as bad_debt_date,
                    isnull((select U_name FROM User_M where U_num = Rd.bad_debt_num AND del_tag = '0'),'') as bad_debt_name,Rd.invoice_no,
                    FORMAT(Rd.invoice_date, 'yyyy/MM/dd') as invoice_date,
                    interest_rate_pass,case when RC_date<SYSDATETIME() and check_pay_type = 'Y' then 'Y' else 'N' end isOver_RC,
                    case when Hs.sendcase_handle_date >='2023-04-24 00:00:00.000' then 'Y' else 'N' end isNewFun, Rd.RemainingPrincipal,
                    case when DATEDIFF(day, RC_date, check_pay_date) > 0 then CEILING(RC_amount*0.2/365*DATEDIFF(day, RC_date, check_pay_date)) else 0 end Fee,
                    case when DATEDIFF(day, RC_date, check_pay_date) > 0 then convert(decimal(12,6),convert(decimal(12,6),convert(decimal(4,2),convert(decimal(4,2),interest_rate_pass)/100)*5/10000*DATEDIFF(day, RC_date, check_pay_date))) else 1 end EXrate
                    FROM Receivable_D Rd
                    LEFT JOIN Receivable_M Rm ON Rm.RCM_id = Rd.RCM_id AND Rm.del_tag = '0'  
                    LEFT JOIN House_apply Ha ON Ha.HA_id = Rm.HA_id AND Ha.del_tag = '0'
                    LEFT JOIN (SELECT U_num, U_BC FROM User_M) Um ON Um.U_num = Ha.plan_num
                    LEFT JOIN House_sendcase Hs ON Hs.HS_id = Rm.HS_id AND Hs.del_tag = '0'
                    WHERE Rd.del_tag = '0' and Um.U_BC IN ('zz', 'BC0100', 'BC0200', 'BC0600', 'BC0900', 'BC0700', 'BC0800', 'BC0300', 'BC0500', 'BC0400', 'BC0800')
                    and Rd.cancel_type = 'N'                    
                    and Rd.RC_date >= @Date_S + ' 00:00:00' AND Rd.RC_date <= @Date_E + ' 23:59:59'";

                parameters.Add(new SqlParameter("@Date_S", model.str_Date_S));
                parameters.Add(new SqlParameter("@Date_E", model.str_Date_E));
                if (!string.IsNullOrEmpty(model.name))
                {
                    T_SQL = T_SQL + " and Ha.CS_name = @CS_name";
                    parameters.Add(new SqlParameter("@CS_name", model.name));
                }
                if (model.check_type != "A")
                {
                    T_SQL = T_SQL + " and Rd.check_pay_type = @check_pay_type";
                    parameters.Add(new SqlParameter("@check_pay_type", model.check_type));
                }
                if (!string.IsNullOrEmpty(model.NS_type))
                {
                    T_SQL = T_SQL + "  AND Rd.RCM_id not in (select distinct RCM_id from Receivable_D where check_pay_type='S')";
                }
                if (model.bad_type != "A")
                {
                    T_SQL = T_SQL + " and Rd.bad_debt_type = @bad_debt_type";
                    parameters.Add(new SqlParameter("@bad_debt_type", model.bad_type));
                }
                if (!string.IsNullOrEmpty(model.RC_count))
                {
                    T_SQL = T_SQL + " and Rd.RC_count = @RC_count";
                    parameters.Add(new SqlParameter("@RC_count", model.RC_count));
                }
                T_SQL = T_SQL + " ORDER BY Rd.RC_date";
                #endregion
                var result = _adoData.ExecuteQuery(T_SQL, parameters).AsEnumerable().Select(row => new Receivable_res
                {
                    RCM_id = row.Field<decimal>("RCM_id"),
                    RCD_id = row.Field<decimal>("RCD_id"),
                    CS_name = row.Field<string>("CS_name"),
                    str_amount_total = row.Field<string>("str_amount_total"),
                    RC_count = row.Field<int>("RC_count"),
                    month_total = row.Field<int>("month_total"),
                    RC_date = FuncHandler.ConvertGregorianToROC(row.Field<string>("RC_date")),
                    str_RC_amount = row.Field<string>("str_RC_amount"),
                    str_interest = row.Field<string>("str_interest"),
                    str_Rmoney = row.Field<string>("str_Rmoney"),
                    str_RemainingAmount = row.Field<string>("str_RemainingAmount"),
                    str_PartiallySettled = row.Field<string>("str_PartiallySettled"),
                    DelayDay = row.Field<int>("DelayDay"),
                    check_pay_type = row.Field<string>("check_pay_type"),
                    check_pay_date = row.Field<string>("check_pay_date"),
                    check_pay_name = row.Field<string>("check_pay_name"),
                    RC_note = row.Field<string>("RC_note"),
                    bad_debt_type = row.Field<string>("bad_debt_type"),
                    bad_debt_date = row.Field<string>("bad_debt_date"),
                    bad_debt_name = row.Field<string>("bad_debt_name"),
                    invoice_no = row.Field<string>("invoice_no"),
                    invoice_date = row.Field<string>("invoice_date"),
                    interest_rate_pass = row.Field<string>("interest_rate_pass"),
                    isOver_RC = row.Field<string>("isOver_RC"),
                    isNewFun = row.Field<string>("isNewFun"),
                    RemainingPrincipal = row.Field<decimal?>("RemainingPrincipal"),
                    Fee = row.Field<decimal?>("Fee"),
                    EXrate = row.Field<decimal?>("EXrate")
                }).ToList();
                foreach (var item in result)
                {
                   

                    //是否呈現延滯天數 延滯利息
                    if (item.isOver_RC == "Y")
                    {
                        int Fee = 0;

                        if (item.DelayDay >= 3 && item.DelayDay <= 6)
                            Fee = 100;
                        if (item.DelayDay >= 7 && item.DelayDay <= 14)
                            Fee = 200;
                        if (item.DelayDay > 15)
                            Fee = 300;

                        if (item.isNewFun == "Y")
                        {
                            decimal rcAmount = Convert.ToDecimal(item.str_RC_amount.Replace(",", ""));
                            item.Delaymoney = Math.Ceiling(Convert.ToDecimal(Fee + Math.Ceiling(Convert.ToDouble(rcAmount)) * 0.16 / 365 * item.DelayDay));
                        }
                        else
                        {
                            decimal remainingPrincipal = Convert.ToDecimal(item.RemainingPrincipal);
                            decimal rmoney = Convert.ToDecimal(item.str_Rmoney);
                            decimal exrate = Convert.ToDecimal(item.EXrate);
                            item.Delaymoney = Math.Ceiling(Convert.ToDecimal((Fee + item.Fee + Math.Ceiling((Math.Ceiling(remainingPrincipal) + Math.Ceiling(rmoney)) * exrate))));
                        }
                    }
                    else
                    {
                        item.DelayDay = null;
                        item.Delaymoney = null;
                    }

                    if (item.DelayDay <= 0)
                    {
                        item.DelayDay = 0;
                        item.Delaymoney = 0;
                    }
                        
                }

                List<Receivable_Excel> excelList = result.Select(x => new Receivable_Excel
                {
                    CS_name = x.CS_name,
                    amount_total = decimal.TryParse(x.str_amount_total, out var amountTotal) ? amountTotal : 0,
                    month_total = x.month_total,
                    RC_count = x.RC_count,
                    RC_date = x.RC_date,
                    RC_amount = decimal.TryParse(x.str_RC_amount, out var rcAmount) ? rcAmount : 0,
                    interest = decimal.TryParse(x.str_interest, out var interest) ? interest : 0,
                    Rmoney = decimal.TryParse(x.str_Rmoney, out var rmoney) ? rmoney : 0,
                    RemainingAmount = decimal.TryParse(x.str_RemainingAmount, out var remainingAmount) ? remainingAmount : 0,
                    PartiallySettled = int.TryParse(x.str_PartiallySettled, out var partiallySettled) ? partiallySettled : 0,
                    DelayDay = x.DelayDay,
                    Delaymoney = x.Delaymoney,
                    check_pay_type = x.check_pay_type,
                    check_pay_date = x.check_pay_date,
                    check_pay_name = x.check_pay_name,
                    RC_note = x.RC_note,
                    bad_debt_type = x.bad_debt_type,
                    bad_debt_date = x.bad_debt_date,
                    bad_debt_name = x.bad_debt_name,
                    invoice_no = x.invoice_no,
                    invoice_date = x.invoice_date,
                }).ToList();

                var Excel_Headers = new Dictionary<string, string>
                {
                    { "index", "序號" },
                    { "CS_name", "客戶姓名" },
                    { "str_amount_total", "總金額" },
                    { "month_total", "期數" },
                    { "RC_count", "第幾期" },
                    { "RC_date", "本期繳款日" },
                    { "str_RC_amount", "月付金" },
                    { "str_interest", "利息" },
                    { "str_Rmoney", "償還本金" },
                    { "str_RemainingAmount", "本金餘額" },
                    { "str_PartiallySettled", "部分結清" },
                    { "DelayDay", "延滯天數" },
                    { "Delaymoney", "延滯利息" },
                    { "check_pay_type", "沖銷狀態" },
                    { "check_pay_date", "沖銷日期" },
                    { "check_pay_name", "沖銷人員" },
                    { "RC_note", "備註" },
                    { "bad_debt_type", "轉呆狀態" },
                    { "bad_debt_date", "轉呆日期" },
                    { "bad_debt_name", "轉呆人員" },
                    { "invoice_no", "發票號碼" },
                    { "invoice_date", "發票日期" }
                };

                var fileBytes = FuncHandler.ReceivableExcel(excelList, Excel_Headers);
                var fileName = "應收款項" + DateTime.Now.ToString("yyyyMMddHHmm") + ".xlsx";
                return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            }
            catch (Exception ex)
            {
                ResultClass<string> resultClass = new ResultClass<string>();
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass);
            }
        }
        #endregion

        #region 應收帳款-催收 
        /// <summary>
        /// 應收帳款催收查詢 RC_D_New_LQuery/RC_D_list_New.asp
        /// </summary>
        [HttpPost("RC_D_New_LQuery")]
        public ActionResult<ResultClass<string>> RC_D_New_LQuery(Receivable_req_new model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                model.str_Date_S = FuncHandler.ConvertROCToGregorian(model.str_Date_S);
                model.str_Date_E = FuncHandler.ConvertROCToGregorian(model.str_Date_E);

                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"select Ha.CS_name,Rm.amount_total,Rm.month_total,Rd.RC_count,FORMAT(Rd.RC_date,'yyyy/MM/dd') AS RC_date,DATEDIFF(day,Rd.RC_date,SYSDATETIME()) DiffDay,
                    Rd.RC_amount,Hs.interest_rate_pass,Rm.loan_grace_num
                    from (
                           select bad_debt_type,check_pay_type,cancel_type,RC_amount,RCM_id,cancel_num,bad_debt_num,check_pay_num,add_num,
                    	   min(RC_count) RC_count,min(RC_date) RC_date 
                           from Receivable_D where del_tag = '0' and check_pay_type='N' and bad_debt_type='N' and cancel_type='N' 
                           group by bad_debt_type,check_pay_type,cancel_type,RCM_id,cancel_num,bad_debt_num,check_pay_num,add_num,RC_amount
                    	 ) Rd  
                    LEFT JOIN Receivable_M Rm on Rm.RCM_id = Rd.RCM_id AND Rm.del_tag='0'  
                    LEFT JOIN House_apply Ha on Ha.HA_id = Rm.HA_id AND Ha.del_tag='0'  
                    LEFT JOIN (select U_num ,U_BC FROM User_M) Um ON Um.U_num = Ha.plan_num  
                    LEFT JOIN House_sendcase Hs on Hs.HS_id = Rm.HS_id AND Hs.del_tag='0'  
                    where 1=1  AND (Rd.RC_date >= @Date_S + ' 00:00:00' AND Rd.RC_date <= @Date_E + ' 23:59:59')
                    and Um.U_BC in ('zz','BC0100','BC0200','BC0600','BC0900','BC0700','BC0800','BC0300','BC0500','BC0400','BC0800')";
                parameters.Add(new SqlParameter("@Date_S", model.str_Date_S));
                parameters.Add(new SqlParameter("@Date_E", model.str_Date_E));
                if (!string.IsNullOrEmpty(model.name))
                {
                    T_SQL = T_SQL + " and Ha.CS_name = @Cs_name";
                    parameters.Add(new SqlParameter("@Cs_name", model.name));
                }
                if (!string.IsNullOrEmpty(model.DiffDay_Type) && model.DiffDay_Type == "0")
                {
                    T_SQL = T_SQL + " and DATEDIFF(day,Rd.RC_date,SYSDATETIME()) between 1 and 29 ";
                }
                if (!string.IsNullOrEmpty(model.DiffDay_Type) && model.DiffDay_Type == "A")
                {
                    T_SQL = T_SQL + " and DATEDIFF(day,Rd.RC_date,SYSDATETIME()) between 30 and 59 ";
                }
                if (!string.IsNullOrEmpty(model.DiffDay_Type) && model.DiffDay_Type == "B")
                {
                    T_SQL = T_SQL + " and DATEDIFF(day,Rd.RC_date,SYSDATETIME()) between 60 and 89";
                }
                if (!string.IsNullOrEmpty(model.DiffDay_Type) && model.DiffDay_Type == "C")
                {
                    T_SQL = T_SQL + " and DATEDIFF(day,Rd.RC_date,SYSDATETIME()) >= 90";
                }
                T_SQL = T_SQL +" order by  Rd.RC_date";
                #endregion
                var result = _adoData.ExecuteQuery(T_SQL, parameters).AsEnumerable().Select(row => new Receivable_res_new
                {
                    CS_name = row.Field<string>("CS_name"),
                    amount_total = row.Field<decimal>("amount_total"),
                    month_total = row.Field<int>("month_total"),
                    RC_count = row.Field<int>("RC_count"),
                    RC_date = FuncHandler.ConvertGregorianToROC(row.Field<string>("RC_date")),
                    DiffDay = row.Field<int>("DiffDay"),
                    RC_amount = row.Field<decimal>("RC_amount"),
                    interest_rate_pass = row.Field<string>("interest_rate_pass"),
                    loan_grace_num = row.Field<int>("loan_grace_num")
                }).ToList();

                foreach (var item in result)
                {
                    decimal interest_rate_pass = Convert.ToInt32(item.interest_rate_pass);
                    decimal amountTotal = Math.Round(Convert.ToDecimal(item.amount_total), 2);
                    decimal interestRatePass = Convert.ToDecimal(item.interest_rate_pass) / 100 / 12;
                    int monthTotal = Convert.ToInt32(item.month_total);
                    decimal denominator = 1 - (1 / (decimal)Math.Pow(1 + (double)interestRatePass, monthTotal));
                    decimal realRCAmount = Math.Round(amountTotal * (interestRatePass / denominator), 2);

                    if ((item.RC_count - item.loan_grace_num) <= 1)
                    {
                        item.interest = Math.Round(item.amount_total, 2) * interest_rate_pass / 100 / 12;
                        if (item.RC_count > item.loan_grace_num)
                        {
                            item.Rmoney = Math.Round(realRCAmount - item.interest, 2);
                        }
                        else
                        {
                            item.Rmoney = 0;
                        }
                        item.RemainingPrincipal = Math.Round(item.amount_total - item.Rmoney,2);
                        item.RemainingPrincipal_1 = Math.Round(item.RemainingPrincipal + item.Rmoney,2);
                    }
                    else
                    {
                        item.RemainingPrincipal_1 = 0;
                        //第一期 數字
                        for (int i = item.loan_grace_num + 1; i <= item.RC_count; i++)
                        {
                            if (i == item.loan_grace_num + 1)
                            {
                                item.RemainingPrincipal_1 = Math.Round(item.amount_total, 2);
                            }
                            decimal monthlyInterestRate = interest_rate_pass / 100 / 12;
                            item.interest = Math.Round(Math.Round(item.RemainingPrincipal_1, 2) * monthlyInterestRate, 2);
                            item.Rmoney = Math.Round(Math.Round(realRCAmount, 2) - item.interest,2);
                            item.RemainingPrincipal = Math.Round(item.RemainingPrincipal_1 - item.Rmoney, 2);
                            item.RemainingPrincipal_1 = item.RemainingPrincipal;
                        }
                        item.RemainingPrincipal_1 = Math.Round(item.RemainingPrincipal + item.Rmoney, 2);
                    }
                    item.interest = Math.Round(item.interest, 0, MidpointRounding.AwayFromZero);
                    item.Rmoney = Math.Round(item.Rmoney, 0, MidpointRounding.AwayFromZero);
                    item.RemainingPrincipal = Math.Round(item.RemainingPrincipal, 0, MidpointRounding.AwayFromZero);
                    item.RemainingPrincipal_1 = Math.Round(item.RemainingPrincipal_1, 0, MidpointRounding.AwayFromZero);
                }

                if (result.Count > 0)
                {
                    resultClass.ResultCode = "000";
                    resultClass.objResult = JsonConvert.SerializeObject(result);
                    return Ok(resultClass);
                }
                else
                {
                    resultClass.ResultCode = "400";
                    resultClass.ResultMsg = "查無資料";
                    return BadRequest(resultClass);
                }
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                return StatusCode(500, resultClass);
            }
        }
        [HttpPost("RC_daily_New_Excel")]
        public IActionResult RC_daily_New_Excel(Receivable_req_new model)
        {
            try
            {
                model.str_Date_S = FuncHandler.ConvertROCToGregorian(model.str_Date_S);
                model.str_Date_E = FuncHandler.ConvertROCToGregorian(model.str_Date_E);

                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"select Ha.CS_name,Rm.amount_total,Rm.month_total,Rd.RC_count,FORMAT(Rd.RC_date,'yyyy/MM/dd') AS RC_date,DATEDIFF(day,Rd.RC_date,SYSDATETIME()) DiffDay,
                    Rd.RC_amount,Hs.interest_rate_pass,Rm.loan_grace_num
                    from (
                           select bad_debt_type,check_pay_type,cancel_type,RC_amount,RCM_id,cancel_num,bad_debt_num,check_pay_num,add_num,
                    	   min(RC_count) RC_count,min(RC_date) RC_date 
                           from Receivable_D where del_tag = '0' and check_pay_type='N' and bad_debt_type='N' and cancel_type='N' 
                           group by bad_debt_type,check_pay_type,cancel_type,RCM_id,cancel_num,bad_debt_num,check_pay_num,add_num,RC_amount
                    	 ) Rd  
                    LEFT JOIN Receivable_M Rm on Rm.RCM_id = Rd.RCM_id AND Rm.del_tag='0'  
                    LEFT JOIN House_apply Ha on Ha.HA_id = Rm.HA_id AND Ha.del_tag='0'  
                    LEFT JOIN (select U_num ,U_BC FROM User_M) Um ON Um.U_num = Ha.plan_num  
                    LEFT JOIN House_sendcase Hs on Hs.HS_id = Rm.HS_id AND Hs.del_tag='0'  
                    where 1=1  AND (Rd.RC_date >= @Date_S + ' 00:00:00' AND Rd.RC_date <= @Date_E + ' 23:59:59')
                    and Um.U_BC in ('zz','BC0100','BC0200','BC0600','BC0900','BC0700','BC0800','BC0300','BC0500','BC0400','BC0800')";
                parameters.Add(new SqlParameter("@Date_S", model.str_Date_S));
                parameters.Add(new SqlParameter("@Date_E", model.str_Date_E));
                if (!string.IsNullOrEmpty(model.name))
                {
                    T_SQL = T_SQL + " and Ha.CS_name = @Cs_name";
                    parameters.Add(new SqlParameter("@Cs_name", model.name));
                }
                if (!string.IsNullOrEmpty(model.DiffDay_Type) && model.DiffDay_Type == "0")
                {
                    T_SQL = T_SQL + " and DATEDIFF(day,Rd.RC_date,SYSDATETIME()) between 1 and 29 ";
                }
                if (!string.IsNullOrEmpty(model.DiffDay_Type) && model.DiffDay_Type == "A")
                {
                    T_SQL = T_SQL + " and DATEDIFF(day,Rd.RC_date,SYSDATETIME()) between 30 and 59 ";
                }
                if (!string.IsNullOrEmpty(model.DiffDay_Type) && model.DiffDay_Type == "B")
                {
                    T_SQL = T_SQL + " and DATEDIFF(day,Rd.RC_date,SYSDATETIME()) between 60 and 89";
                }
                if (!string.IsNullOrEmpty(model.DiffDay_Type) && model.DiffDay_Type == "C")
                {
                    T_SQL = T_SQL + " and DATEDIFF(day,Rd.RC_date,SYSDATETIME()) >= 90";
                }
                T_SQL = T_SQL + " order by  Rd.RC_date";
                #endregion
                var result = _adoData.ExecuteQuery(T_SQL, parameters).AsEnumerable().Select(row => new Receivable_res_new
                {
                    CS_name = row.Field<string>("CS_name"),
                    amount_total = row.Field<decimal>("amount_total"),
                    month_total = row.Field<int>("month_total"),
                    RC_count = row.Field<int>("RC_count"),
                    RC_date = FuncHandler.ConvertGregorianToROC(row.Field<string>("RC_date")),
                    DiffDay = row.Field<int>("DiffDay"),
                    RC_amount = row.Field<decimal>("RC_amount"),
                    interest_rate_pass = row.Field<string>("interest_rate_pass"),
                    loan_grace_num = row.Field<int>("loan_grace_num")
                }).ToList();
                foreach (var item in result)
                {
                    decimal interest_rate_pass = Convert.ToInt32(item.interest_rate_pass);
                    decimal amountTotal = Math.Round(Convert.ToDecimal(item.amount_total), 2);
                    decimal interestRatePass = Convert.ToDecimal(item.interest_rate_pass) / 100 / 12;
                    int monthTotal = Convert.ToInt32(item.month_total);
                    decimal denominator = 1 - (1 / (decimal)Math.Pow(1 + (double)interestRatePass, monthTotal));
                    decimal realRCAmount = Math.Round(amountTotal * (interestRatePass / denominator), 2);

                    if ((item.RC_count - item.loan_grace_num) <= 1)
                    {
                        item.interest = Math.Round(item.amount_total, 2) * interest_rate_pass / 100 / 12;
                        if (item.RC_count > item.loan_grace_num)
                        {
                            item.Rmoney = Math.Round(realRCAmount - item.interest, 2);
                        }
                        else
                        {
                            item.Rmoney = 0;
                        }
                        item.RemainingPrincipal = Math.Round(item.amount_total - item.Rmoney, 2);
                        item.RemainingPrincipal_1 = Math.Round(item.RemainingPrincipal + item.Rmoney, 2);
                    }
                    else
                    {
                        item.RemainingPrincipal_1 = 0;
                        //第一期 數字
                        for (int i = item.loan_grace_num + 1; i <= item.RC_count; i++)
                        {
                            if (i == item.loan_grace_num + 1)
                            {
                                item.RemainingPrincipal_1 = Math.Round(item.amount_total, 2);
                            }
                            decimal monthlyInterestRate = interest_rate_pass / 100 / 12;
                            item.interest = Math.Round(Math.Round(item.RemainingPrincipal_1, 2) * monthlyInterestRate, 2);
                            item.Rmoney = Math.Round(Math.Round(realRCAmount, 2) - item.interest, 2);
                            item.RemainingPrincipal = Math.Round(item.RemainingPrincipal_1 - item.Rmoney, 2);
                            item.RemainingPrincipal_1 = item.RemainingPrincipal;
                        }
                        item.RemainingPrincipal_1 = Math.Round(item.RemainingPrincipal + item.Rmoney, 2);
                    }
                    item.interest = Math.Round(item.interest, 0, MidpointRounding.AwayFromZero);
                    item.Rmoney = Math.Round(item.Rmoney, 0, MidpointRounding.AwayFromZero);
                    item.RemainingPrincipal = Math.Round(item.RemainingPrincipal, 0, MidpointRounding.AwayFromZero);
                    item.RemainingPrincipal_1 = Math.Round(item.RemainingPrincipal_1, 0, MidpointRounding.AwayFromZero);
                }
                
                var excelList = result.Select(x => new Receivable_New_Excel
                {
                    CS_name = x.CS_name,
                    amount_total = x.amount_total,
                    month_total = x.month_total,
                    RC_count = x.RC_count,
                    RC_date = x.RC_date,
                    DiffDay = x.DiffDay,
                    RC_amount = x.RC_amount,
                    interest = x.interest,
                    Rmoney = x.Rmoney,
                    RemainingPrincipal = x.RemainingPrincipal,
                    RemainingPrincipal_1 = x.RemainingPrincipal_1
                }).ToList();
                
                var Excel_Headers = new Dictionary<string, string>
                {
                    {"index","序號" },
                    { "CS_name", "客戶姓名" },
                    { "amount_total", "總金額" },
                    { "month_total", "期數" },
                    { "RC_count", "第幾期" },
                    { "RC_date", "本期繳款日" },
                    { "DiffDay", "逾期天數" },
                    { "RC_amount", "月付金" },
                    { "interest", "利息" },
                    { "Rmoney", "償還本金" },
                    { "RemainingPrincipal", "本金餘額" },
                    { "RemainingPrincipal_1", "實際本金餘額" }
                };

                var fileBytes = FuncHandler.ReceivableNewExcel(excelList, Excel_Headers);
                var fileName = "催收款項" + DateTime.Now.ToString("yyyyMMddHHmm") + ".xlsx";
                return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            }
            catch (Exception ex)
            {
                ResultClass<string> resultClass = new ResultClass<string>();
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass);
            }
        }
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
