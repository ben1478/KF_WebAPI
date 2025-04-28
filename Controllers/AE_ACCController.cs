using KF_WebAPI.BaseClass;
using KF_WebAPI.BaseClass.AE;
using KF_WebAPI.BaseClass.Max104;
using KF_WebAPI.FunctionHandler;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Http.Extensions;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Microsoft.SqlServer.Server;
using Newtonsoft.Json;
using System;
using System.Collections;
using System.Data;
using System.Data.SqlTypes;
using System.Reflection;
using System.Text.RegularExpressions;

namespace KF_WebAPI.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class AE_ACCController : ControllerBase
    {
        /// <summary>
        /// 抓取延滯利息金額
        /// </summary>
        /// <returns></returns>
        [HttpGet("GetDelayAMT")]
        public ActionResult<ResultClass<string>> GetDelayAMT(string RCM_id, string RC_date_E)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"select dbo.GetTotalDelay_AMT (@RCM_id,@RC_date_E) Delay_AMT";
                var parameters = new List<SqlParameter>
                {
                    new SqlParameter("@RCM_id", RCM_id),
                    new SqlParameter("@RC_date_E",RC_date_E)
                };
                #endregion
                var dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
                resultClass.ResultCode = "000";
                resultClass.objResult = JsonConvert.SerializeObject(dtResult);
                return Ok(resultClass);
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass);
            }
        }

        /// <summary>
        /// 提供分期人員清單
        /// </summary>
        /// <returns></returns>
        [HttpGet("Settle_Show_LQuery")]
        public ActionResult<ResultClass<string>> Settle_Show_LQuery(string? CS_name, string? CS_PID)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                var parameters = new List<SqlParameter>();
                #region SQL
                var T_SQL = @"select HA.CS_name,HA.CS_PID,RM.amount_total,RM.month_total,RM.amount_per_month,RM.RCM_id,HS.HS_id  
                              ,CASE WHEN sendcase_handle_date >='2023-04-24 00:00:00.000' THEN 'Y' ELSE 'N' END as isNewFun
                              from Receivable_M RM
                              INNER JOIN 
                              (
                                select * from 
                                (
                                  select ROW_NUMBER() OVER (PARTITION BY RCM_id ORDER BY RC_count ASC) AS rn,* from Receivable_D
                                  where bad_debt_type='N' AND cancel_type='N' AND del_tag = '0' AND isnull(RecPayAmt,0) = 0 and check_pay_type='N'
                                ) RD_filtered where RD_filtered.rn = 1
                              ) RD on RM.RCM_id = RD.RCM_id 
                              LEFT JOIN House_apply HA ON HA.HA_id = RM.HA_id AND HA.del_tag='0'
                              LEFT JOIN House_sendcase HS ON HS.HS_id = RM.HS_id
                              where RM.del_tag='0' AND RM.del_tag='0'";
                if (!string.IsNullOrEmpty(CS_name))
                {
                    T_SQL += " and CS_name=@CS_name ";
                    parameters.Add(new SqlParameter("@CS_name", CS_name));
                }

                if (!string.IsNullOrEmpty(CS_PID))
                {
                    T_SQL += " and CS_PID=@CS_PID ";
                    parameters.Add(new SqlParameter("@CS_PID", CS_PID));
                }
                T_SQL += " ORDER BY RD.RC_date";
                #endregion
                var dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
                resultClass.ResultCode = "000";
                resultClass.objResult = JsonConvert.SerializeObject(dtResult);
                return Ok(resultClass);
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass);
            }
        }

        /// <summary>
        /// 提供結清紀錄
        /// </summary>
        /// <param name="RCM_ID"></param>
        [HttpGet("Settle_Show_SQuery")]
        public ActionResult<ResultClass<string>> Settle_Show_SQuery(string RCM_ID)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                var isFirst = "N";
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = "";
                //先判定是否首期繳清
                var T_SQL_First = @"select Case When isnull(MAX(RC_count),0) <= 1 Then 'Y' ELSE 'N' End isFirst 
                                    from view_RC_Dtl where RCM_id = @RCM_ID";
                var parameters_f = new List<SqlParameter>
                {
                    new SqlParameter("@RCM_ID",RCM_ID)
                };
                var dtResult_f = _adoData.ExecuteQuery(T_SQL_First, parameters_f);
                isFirst = dtResult_f.Rows[0]["isFirst"].ToString();
                if (isFirst == "Y")
                {
                    T_SQL = @"SELECT M.amount_total RP_AMT,isnull(M.Break_AMT,CEILING(M.amount_total * 0.13)) Break_AMT,
                                     S.get_amount_date RC_date,D1.RC_count,convert(varchar(10),isnull(M.date_begin_settle,SYSDATETIME()),111)OffDate,
                                     D1.interest,Case When DATEPART(Day,D1.RC_date) = DATEPART(Day,isnull(M.date_begin_settle, SYSDATETIME())) and DATEDIFF(MONTH, D1.RC_date, isnull(M.date_begin_settle, SYSDATETIME())) = 1 
                                     THEN 30 ELSE DATEDIFF(Day, D1.RC_date, isnull(M.date_begin_settle, SYSDATETIME())) end interestDay,
                                     isnull(M.Interest_AMT,0) interest_AMT,
                              	     0 as Delay_AMT,'Y' as isFirst,amount_per_month,H.CS_name
                              FROM Receivable_M M
                              LEFT JOIN House_sendcase S ON M.HS_id=S.HS_id
                              LEFT JOIN Receivable_D D1 ON M.RCM_id=D1.RCM_id
                              LEFT JOIN House_apply H ON H.HA_id = M.HA_id
                              WHERE M.RCM_id =　@RCM_id AND D1.RC_count=1";
                }
                else
                {
                    T_SQL = @"SELECT D.Ex_RemainingPrincipal RP_AMT,isnull(M.Break_AMT,CEILING(D.Ex_RemainingPrincipal * 0.13)) Break_AMT,
                              convert(varchar(10),D.RC_date,111) RC_date,D.RC_count,convert(varchar(10),isnull(M.date_begin_settle,SYSDATETIME()),111) OffDate,
                              D1.interest,Case When DATEPART(Day,D.RC_date) = DATEPART(Day,isnull(M.date_begin_settle, SYSDATETIME())) and DATEDIFF(MONTH, D.RC_date, isnull(M.date_begin_settle, SYSDATETIME())) = 1 
                              THEN 30 ELSE DATEDIFF(Day, D.RC_date, isnull(M.date_begin_settle, SYSDATETIME())) end interestDay,
                              isnull(M.Interest_AMT,0) interest_AMT,
                              (select dbo.GetTotalDelay_AMT(@RCM_ID,SYSDATETIME())) Delay_AMT,'N' as isFirst,M.amount_per_month,H.CS_name
                              FROM Receivable_D D
                              LEFT JOIN Receivable_D D1 ON D.RCM_id=D1.RCM_id AND (D.RC_count+1) = D1.RC_count
                              LEFT JOIN Receivable_M M ON D.RCM_id=M.RCM_id
                              LEFT JOIN House_apply H ON H.HA_id = M.HA_id
                              WHERE D.RCM_id = @RCM_ID
                              AND D.RC_count IN
                              ( SELECT MAX(RC_count ) FROM Receivable_D
                              WHERE RCM_id = @RCM_ID
                              AND RC_note not like '%清償%'
                              AND check_pay_type <> 'S'
                              AND check_pay_type='Y')";
                }

                var parameters = new List<SqlParameter>
                {
                    new SqlParameter("@RCM_ID",RCM_ID)
                };
                #endregion
                var dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
                resultClass.ResultCode = "000";
                resultClass.objResult = JsonConvert.SerializeObject(dtResult);
                return Ok(resultClass);
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass);
            }
        }

        /// <summary>
        /// 取得法院相關費用總金額
        /// </summary>
        [HttpGet("GetCourtAmt")]
        public ActionResult<ResultClass<string>> GetCourtAmt (string HS_id)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"SELECT HS_ID,sum(cast(isnull(L.charge_AMT, 0) AS decimal))TotAmt FROM Item_list I
                              LEFT JOIN(SELECT HS_ID,charge_AMT,charge_type_M,charge_type_D FROM House_sendcase_other_charge 
                              WHERE isnull(del_tag, '0')='0' AND charge_type_M='Court_exe'
                              AND HS_ID=@HS_id) L ON item_M_code=charge_type_M AND item_D_code=charge_type_D 
                              WHERE item_M_code ='Court_exe' AND item_M_type='N' AND charge_AMT IS NOT NULL GROUP BY HS_ID";
                var parameters = new List<SqlParameter>
                {
                    new SqlParameter("@HS_id",HS_id)
                };
                #endregion
                var dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
                if (dtResult.Rows.Count > 0)
                {
                    resultClass.ResultCode = "000";
                    resultClass.objResult = JsonConvert.SerializeObject(dtResult);
                    return Ok(resultClass);
                }
                else
                {
                    DataTable dtDefault = new DataTable();
                    dtDefault.Columns.Add("HS_ID", typeof(string));
                    dtDefault.Columns.Add("TotAmt", typeof(decimal));

                    DataRow dr = dtDefault.NewRow();
                    dr["HS_ID"] = HS_id;  
                    dr["TotAmt"] = 0;
                    dtDefault.Rows.Add(dr);

                    resultClass.ResultCode = "000";
                    resultClass.objResult = JsonConvert.SerializeObject(dtDefault);
                    return Ok(resultClass);
                }
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass);
            }
        }

        /// <summary>
        /// 取得法院相關費用明細
        /// </summary>
        [HttpGet("GetCourtDetail")]
        public ActionResult<ResultClass<string>> GetCourtDetail(string HS_id)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"SELECT Li.item_D_name as showName,FORMAT(Hs.charge_AMT,'N0') as showAmt FROM House_sendcase_other_charge Hs
                              LEFT JOIN Item_list Li on Li.item_M_code='Court_exe' and item_D_code =Hs.charge_type_D
                              WHERE isnull(Hs.del_tag, '0')='0'
                              AND charge_type_M='Court_exe'
                              AND HS_ID=@HS_id";
                var parameters = new List<SqlParameter>
                {
                    new SqlParameter("@HS_id",HS_id)
                };
                #endregion
                var dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
                if (dtResult.Rows.Count > 0)
                {
                    resultClass.ResultCode = "000";
                    resultClass.objResult = JsonConvert.SerializeObject(dtResult);
                    return Ok(resultClass);
                }
                else
                {
                    resultClass.ResultCode = "400";
                    resultClass.objResult = JsonConvert.SerializeObject(dtResult);
                    return Ok(resultClass);
                }
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass);
            }
        }

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
                resultClass.ResultMsg = $" response: {ex.Message}";
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
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass);
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
                resultClass.ResultMsg = $" response: {ex.Message}";
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
                resultClass.ResultMsg = $" response: {ex.Message}";
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
                resultClass.ResultMsg = $" response: {ex.Message}";
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
                resultClass.ResultMsg = $" response: {ex.Message}";
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
                resultClass.ResultMsg = $" response: {ex.Message}";
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
                resultClass.ResultMsg = $" response: {ex.Message}";
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
                var cknum = FuncHandler.GetCheckNum();
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
                            var cknum_d = FuncHandler.GetCheckNum();
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
                resultClass.ResultMsg = $" response: {ex.Message}";
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
                resultClass.ResultMsg = $" response: {ex.Message}";
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
        [HttpPost("RC_D_Coll_LQuery")]
        public ActionResult<ResultClass<string>> RC_D_Coll_LQuery(Receivable_Coll_req model)
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
                T_SQL = T_SQL + " order by  Rd.RC_date";
                #endregion
                var result = _adoData.ExecuteQuery(T_SQL, parameters).AsEnumerable().Select(row => new Receivable_Coll_res
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
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass);
            }
        }
        [HttpPost("RC_daily_Coll_Excel")]
        public IActionResult RC_daily_Coll_Excel(Receivable_Coll_req model)
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
                var result = _adoData.ExecuteQuery(T_SQL, parameters).AsEnumerable().Select(row => new Receivable_Coll_res
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

                var excelList = result.Select(x => new Receivable_Coll_Excel
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

                var fileBytes = FuncHandler.ReceivableCollExcel(excelList, Excel_Headers);
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
        /// <summary>
        /// 逾期繳款清單查詢 RC_D_Late_Pay_LQuery/RC_D_list_New1.asp
        /// </summary>
        [HttpPost("RC_D_Late_Pay_LQuery")]
        public ActionResult<ResultClass<string>> RC_D_Late_Pay_LQuery(Receivable_Late_Pay_req model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                model.str_Date_E = FuncHandler.ConvertROCToGregorian(model.str_Date_E);

                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"select Rd.RCM_id,Rd.RCD_id,Ha.CS_name,Rm.amount_total,Rm.month_total,Rd.RC_count,FORMAT(Rd.RC_date,'yyyy/MM/dd') AS RC_date,
                    Rd.RC_amount,Rd.interest,Rd.Rmoney,Rd.RemainingPrincipal,
                    isnull(DATEDIFF(DAY, RC_date,case when check_pay_date is null then @Date_E +' 00:00:00' else check_pay_date end),0) DelayDay,
                    Hs.interest_rate_pass,Rm.loan_grace_num
                    from Receivable_D Rd
                    LEFT JOIN Receivable_M Rm on Rm.RCM_id = Rd.RCM_id AND Rm.del_tag = '0'
                    LEFT JOIN House_apply Ha on Ha.HA_id = Rm.HA_id AND Ha.del_tag = '0'
                    LEFT JOIN (select U_num, U_BC FROM User_M) Um ON Um.U_num = Ha.plan_num
                    LEFT JOIN House_sendcase Hs on Hs.HS_id = Rm.HS_id AND Hs.del_tag = '0'  
                    where Rd.del_tag = '0' AND Rd.RC_date <= @Date_E +' 00:00:00' AND (Rd.bad_debt_type = 'N')
                    AND(Rd.cancel_type = 'N') AND Rm.RCM_id is not null AND Rm.RCM_note not like '%提前結清%'";
                if (!string.IsNullOrEmpty(model.pay_type))
                {
                    if (model.pay_type == "Y")
                    {
                        T_SQL = T_SQL + " AND Rd.check_pay_type ='Y' AND (Rd.check_pay_date is null or Rd.check_pay_date <= @Date_E + ' 00:00:00')";
                    }
                    if (model.pay_type == "N")
                    {
                        T_SQL = T_SQL + " AND Rd.check_pay_date is null AND (Rd.check_pay_date is null or Rd.check_pay_date >= @Date_E + ' 00:00:00')";
                    }
                }
                if (!string.IsNullOrEmpty(model.name))
                {
                    T_SQL = T_SQL + " AND Ha.CS_name=@CS_name";
                    parameters.Add(new SqlParameter("@CS_name", model.name));
                }
                if (!string.IsNullOrEmpty(model.delay_type))
                {
                    switch (model.delay_type)
                    {
                        case "0":
                            T_SQL = T_SQL + " AND isnull(DATEDIFF(DAY, RC_date,case when check_pay_date is null then @Date_E + ' 00:00:00' else check_pay_date end),0) between 1 and 29";
                            break;
                        case "1":
                            T_SQL = T_SQL + " AND isnull(DATEDIFF(DAY, RC_date,case when check_pay_date is null then @Date_E + ' 00:00:00' else check_pay_date end),0) >=30";
                            break;
                        case "A":
                            T_SQL = T_SQL + " AND isnull(DATEDIFF(DAY, RC_date,case when check_pay_date is null then @Date_E + ' 00:00:00' else check_pay_date end),0) between 30 and 59";
                            break;
                        case "B":
                            T_SQL = T_SQL + " AND isnull(DATEDIFF(DAY, RC_date,case when check_pay_date is null then @Date_E + ' 00:00:00' else check_pay_date end),0) between 60 and 89";
                            break;
                        case "C":
                            T_SQL = T_SQL + " AND isnull(DATEDIFF(DAY, RC_date,case when check_pay_date is null then @Date_E + ' 00:00:00' else check_pay_date end),0) >=90";
                            break;
                        default:
                            break;
                    }
                }
                else
                {
                    T_SQL = T_SQL + " AND isnull(DATEDIFF(DAY, RC_date,case when check_pay_date is null then @Date_E + ' 00:00:00' else check_pay_date end),0) > 0 ";
                }
                T_SQL = T_SQL + @" AND Um.U_BC in ('zz','BC0100','BC0200','BC0600','BC0900','BC0700','BC0800','BC0300','BC0500','BC0400','BC0800')
                    AND convert(varchar(15),Rd.RCM_ID)+'-'+ convert(varchar(3),(RC_count)) 
                    in (
                         SELECT convert(varchar(15),Receivable_D.RCM_ID)+'-'+ convert(varchar(3),Min(RC_count)) RC_count
                         FROM Receivable_D
                         LEFT JOIN Receivable_M ON Receivable_M.RCM_id = Receivable_D.RCM_id AND Receivable_M.del_tag= '0'
                         LEFT JOIN House_apply ON House_apply.HA_id = Receivable_M.HA_id AND House_apply.del_tag= '0'
                         LEFT JOIN (SELECT U_num, U_BC FROM User_M) User_M ON User_M.U_num = House_apply.plan_num
                         LEFT JOIN House_sendcase ON House_sendcase.HS_id = Receivable_M.HS_id AND House_sendcase.del_tag= '0'
                         WHERE Receivable_D.del_tag = '0' AND Receivable_D.RC_date <= @Date_E +' 00:00:00' AND (Receivable_D.bad_debt_type= 'N')
                         AND(Receivable_D.cancel_type= 'N') AND Receivable_M.RCM_id is not null AND Receivable_M.RCM_note not like '%提前結清%'";
                if (!string.IsNullOrEmpty(model.pay_type))
                {
                    if (model.pay_type == "Y")
                    {
                        T_SQL = T_SQL + " AND Receivable_D.check_pay_type ='Y' AND (Receivable_D.check_pay_date is null or Receivable_D.check_pay_date <= @Date_E + ' 00:00:00')";
                    }
                    if (model.pay_type == "N")
                    {
                        T_SQL = T_SQL + " AND Receivable_D.check_pay_date is null AND (Receivable_D.check_pay_date is null or Receivable_D.check_pay_date >= @Date_E + ' 00:00:00')";
                    }
                }
                if (!string.IsNullOrEmpty(model.name))
                {
                    T_SQL = T_SQL + " AND House_apply.CS_name=@CS_name";
                    parameters.Add(new SqlParameter("@CS_name", model.name));
                }
                if (!string.IsNullOrEmpty(model.delay_type))
                {
                    switch (model.delay_type)
                    {
                        case "0":
                            T_SQL = T_SQL + " AND isnull(DATEDIFF(DAY, RC_date,case when check_pay_date is null then @Date_E + ' 00:00:00' else check_pay_date end),0) between 1 and 29";
                            break;
                        case "1":
                            T_SQL = T_SQL + " AND isnull(DATEDIFF(DAY, RC_date,case when check_pay_date is null then @Date_E + ' 00:00:00' else check_pay_date end),0) >=30";
                            break;
                        case "A":
                            T_SQL = T_SQL + " AND isnull(DATEDIFF(DAY, RC_date,case when check_pay_date is null then @Date_E + ' 00:00:00' else check_pay_date end),0) between 30 and 59";
                            break;
                        case "B":
                            T_SQL = T_SQL + " AND isnull(DATEDIFF(DAY, RC_date,case when check_pay_date is null then @Date_E + ' 00:00:00' else check_pay_date end),0) between 60 and 89";
                            break;
                        case "C":
                            T_SQL = T_SQL + " AND isnull(DATEDIFF(DAY, RC_date,case when check_pay_date is null then @Date_E + ' 00:00:00' else check_pay_date end),0) >=90";
                            break;
                        default:
                            break;
                    }
                }
                else
                {
                    T_SQL = T_SQL + " AND isnull(DATEDIFF(DAY, RC_date,case when check_pay_date is null then @Date_E + ' 00:00:00' else check_pay_date end),0) > 0 ";
                }

                T_SQL = T_SQL + @"AND User_M.U_BC in ('zz','BC0100','BC0200','BC0600','BC0900','BC0700','BC0800','BC0300','BC0500','BC0400','BC0800')
                         group by Receivable_D.RCM_ID) order by Rd.RC_date,CS_name";
                parameters.Add(new SqlParameter("@Date_E", model.str_Date_E));
                #endregion
                var result = _adoData.ExecuteQuery(T_SQL, parameters).AsEnumerable().Select(row => new Receivable_Late_Pay_res
                {
                    RCM_id = row.Field<decimal>("RCM_id"),
                    RCD_id = row.Field<decimal>("RCD_id"),
                    CS_name = row.Field<string>("CS_name"),
                    amount_total = row.Field<decimal>("amount_total"),
                    month_total = row.Field<int>("month_total"),
                    RC_count = row.Field<int>("RC_count"),
                    RC_date = row.Field<string>("RC_date"),
                    RC_amount = row.Field<decimal>("RC_amount"),
                    interest = row.Field<decimal>("interest"),
                    Rmoney = row.Field<decimal>("Rmoney"),
                    RemainingPrincipal = row.Field<decimal?>("RemainingPrincipal"),
                    DelayDay = row.Field<int>("DelayDay"),
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
                    }
                    else
                    {
                        decimal RemainingPrincipal_1 = 0;
                        //第一期 數字
                        for (int i = item.loan_grace_num + 1; i <= item.RC_count; i++)
                        {
                            if (i == item.loan_grace_num + 1)
                            {
                                RemainingPrincipal_1 = Math.Round(item.amount_total, 2);
                            }
                            decimal monthlyInterestRate = interest_rate_pass / 100 / 12;
                            item.interest = Math.Round(Math.Round(RemainingPrincipal_1, 2) * monthlyInterestRate, 2);
                            item.Rmoney = Math.Round(Math.Round(realRCAmount, 2) - item.interest, 2);
                            item.RemainingPrincipal = Math.Round(RemainingPrincipal_1 - item.Rmoney, 2);
                            RemainingPrincipal_1 = Convert.ToDecimal(item.RemainingPrincipal);
                        }
                    }
                    item.interest = Math.Round(item.interest, 0, MidpointRounding.AwayFromZero);
                    item.Rmoney = Math.Round(item.Rmoney, 0, MidpointRounding.AwayFromZero);
                    item.RemainingPrincipal = Math.Round(Convert.ToDecimal(item.RemainingPrincipal), 0, MidpointRounding.AwayFromZero);
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
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass);
            }
        }
        /// <summary>
        /// 逾期繳款清單查詢Excel下載 RC_D_Late_Pay_LQuery/RC_D_list_New1.asp
        /// </summary>
        [HttpPost("RC_D_Late_Pay_Excel")]
        public IActionResult RC_D_Late_Pay_Excel(Receivable_Late_Pay_req model)
        {
            try
            {
                model.str_Date_E = FuncHandler.ConvertROCToGregorian(model.str_Date_E);

                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"select Rd.RCM_id,Rd.RCD_id,Ha.CS_name,Rm.amount_total,Rm.month_total,Rd.RC_count,FORMAT(Rd.RC_date,'yyyy/MM/dd') AS RC_date,
                    Rd.RC_amount,Rd.interest,Rd.Rmoney,Rd.RemainingPrincipal,
                    isnull(DATEDIFF(DAY, RC_date,case when check_pay_date is null then @Date_E +' 00:00:00' else check_pay_date end),0) DelayDay,
                    Hs.interest_rate_pass,Rm.loan_grace_num
                    from Receivable_D Rd
                    LEFT JOIN Receivable_M Rm on Rm.RCM_id = Rd.RCM_id AND Rm.del_tag = '0'
                    LEFT JOIN House_apply Ha on Ha.HA_id = Rm.HA_id AND Ha.del_tag = '0'
                    LEFT JOIN (select U_num, U_BC FROM User_M) Um ON Um.U_num = Ha.plan_num
                    LEFT JOIN House_sendcase Hs on Hs.HS_id = Rm.HS_id AND Hs.del_tag = '0'  
                    where Rd.del_tag = '0' AND Rd.RC_date <= @Date_E +' 00:00:00' AND (Rd.bad_debt_type = 'N')
                    AND(Rd.cancel_type = 'N') AND Rm.RCM_id is not null AND Rm.RCM_note not like '%提前結清%'";
                if (!string.IsNullOrEmpty(model.pay_type))
                {
                    if (model.pay_type == "Y")
                    {
                        T_SQL = T_SQL + " AND Rd.check_pay_type ='Y' AND (Rd.check_pay_date is null or Rd.check_pay_date <= @Date_E + ' 00:00:00')";
                    }
                    if (model.pay_type == "N")
                    {
                        T_SQL = T_SQL + " AND Rd.check_pay_date is null AND (Rd.check_pay_date is null or Rd.check_pay_date >= @Date_E + ' 00:00:00')";
                    }
                }
                if (!string.IsNullOrEmpty(model.name))
                {
                    T_SQL = T_SQL + " AND Ha.CS_name=@CS_name";
                    parameters.Add(new SqlParameter("@CS_name", model.name));
                }
                if (!string.IsNullOrEmpty(model.delay_type))
                {
                    switch (model.delay_type)
                    {
                        case "0":
                            T_SQL = T_SQL + " AND isnull(DATEDIFF(DAY, RC_date,case when check_pay_date is null then @Date_E + ' 00:00:00' else check_pay_date end),0) between 1 and 29";
                            break;
                        case "1":
                            T_SQL = T_SQL + " AND isnull(DATEDIFF(DAY, RC_date,case when check_pay_date is null then @Date_E + ' 00:00:00' else check_pay_date end),0) >=30";
                            break;
                        case "A":
                            T_SQL = T_SQL + " AND isnull(DATEDIFF(DAY, RC_date,case when check_pay_date is null then @Date_E + ' 00:00:00' else check_pay_date end),0) between 30 and 59";
                            break;
                        case "B":
                            T_SQL = T_SQL + " AND isnull(DATEDIFF(DAY, RC_date,case when check_pay_date is null then @Date_E + ' 00:00:00' else check_pay_date end),0) between 60 and 89";
                            break;
                        case "C":
                            T_SQL = T_SQL + " AND isnull(DATEDIFF(DAY, RC_date,case when check_pay_date is null then @Date_E + ' 00:00:00' else check_pay_date end),0) >=90";
                            break;
                        default:
                            break;
                    }
                }
                else
                {
                    T_SQL = T_SQL + " AND isnull(DATEDIFF(DAY, RC_date,case when check_pay_date is null then @Date_E + ' 00:00:00' else check_pay_date end),0) > 0 ";
                }
                T_SQL = T_SQL + @" AND Um.U_BC in ('zz','BC0100','BC0200','BC0600','BC0900','BC0700','BC0800','BC0300','BC0500','BC0400','BC0800')
                    AND convert(varchar(15),Rd.RCM_ID)+'-'+ convert(varchar(3),(RC_count)) 
                    in (
                         SELECT convert(varchar(15),Receivable_D.RCM_ID)+'-'+ convert(varchar(3),Min(RC_count)) RC_count
                         FROM Receivable_D
                         LEFT JOIN Receivable_M ON Receivable_M.RCM_id = Receivable_D.RCM_id AND Receivable_M.del_tag= '0'
                         LEFT JOIN House_apply ON House_apply.HA_id = Receivable_M.HA_id AND House_apply.del_tag= '0'
                         LEFT JOIN (SELECT U_num, U_BC FROM User_M) User_M ON User_M.U_num = House_apply.plan_num
                         LEFT JOIN House_sendcase ON House_sendcase.HS_id = Receivable_M.HS_id AND House_sendcase.del_tag= '0'
                         WHERE Receivable_D.del_tag = '0' AND Receivable_D.RC_date <= @Date_E +' 00:00:00' AND (Receivable_D.bad_debt_type= 'N')
                         AND(Receivable_D.cancel_type= 'N') AND Receivable_M.RCM_id is not null AND Receivable_M.RCM_note not like '%提前結清%'";
                if (!string.IsNullOrEmpty(model.pay_type))
                {
                    if (model.pay_type == "Y")
                    {
                        T_SQL = T_SQL + " AND Receivable_D.check_pay_type ='Y' AND (Receivable_D.check_pay_date is null or Receivable_D.check_pay_date <= @Date_E + ' 00:00:00')";
                    }
                    if (model.pay_type == "N")
                    {
                        T_SQL = T_SQL + " AND Receivable_D.check_pay_date is null AND (Receivable_D.check_pay_date is null or Receivable_D.check_pay_date >= @Date_E + ' 00:00:00')";
                    }
                }
                if (!string.IsNullOrEmpty(model.name))
                {
                    T_SQL = T_SQL + " AND House_apply.CS_name=@CS_name";
                    parameters.Add(new SqlParameter("@CS_name", model.name));
                }
                if (!string.IsNullOrEmpty(model.delay_type))
                {
                    switch (model.delay_type)
                    {
                        case "0":
                            T_SQL = T_SQL + " AND isnull(DATEDIFF(DAY, RC_date,case when check_pay_date is null then @Date_E + ' 00:00:00' else check_pay_date end),0) between 1 and 29";
                            break;
                        case "1":
                            T_SQL = T_SQL + " AND isnull(DATEDIFF(DAY, RC_date,case when check_pay_date is null then @Date_E + ' 00:00:00' else check_pay_date end),0) >=30";
                            break;
                        case "A":
                            T_SQL = T_SQL + " AND isnull(DATEDIFF(DAY, RC_date,case when check_pay_date is null then @Date_E + ' 00:00:00' else check_pay_date end),0) between 30 and 59";
                            break;
                        case "B":
                            T_SQL = T_SQL + " AND isnull(DATEDIFF(DAY, RC_date,case when check_pay_date is null then @Date_E + ' 00:00:00' else check_pay_date end),0) between 60 and 89";
                            break;
                        case "C":
                            T_SQL = T_SQL + " AND isnull(DATEDIFF(DAY, RC_date,case when check_pay_date is null then @Date_E + ' 00:00:00' else check_pay_date end),0) >=90";
                            break;
                        default:
                            break;
                    }
                }
                else
                {
                    T_SQL = T_SQL + " AND isnull(DATEDIFF(DAY, RC_date,case when check_pay_date is null then @Date_E + ' 00:00:00' else check_pay_date end),0) > 0 ";
                }

                T_SQL = T_SQL + @"AND User_M.U_BC in ('zz','BC0100','BC0200','BC0600','BC0900','BC0700','BC0800','BC0300','BC0500','BC0400','BC0800')
                         group by Receivable_D.RCM_ID) order by Rd.RC_date,CS_name";
                parameters.Add(new SqlParameter("@Date_E", model.str_Date_E));
                #endregion
                var result = _adoData.ExecuteQuery(T_SQL, parameters).AsEnumerable().Select(row => new Receivable_Late_Pay_res
                {
                    RCM_id = row.Field<decimal>("RCM_id"),
                    RCD_id = row.Field<decimal>("RCD_id"),
                    CS_name = row.Field<string>("CS_name"),
                    amount_total = row.Field<decimal>("amount_total"),
                    month_total = row.Field<int>("month_total"),
                    RC_count = row.Field<int>("RC_count"),
                    RC_date = row.Field<string>("RC_date"),
                    RC_amount = row.Field<decimal>("RC_amount"),
                    interest = row.Field<decimal>("interest"),
                    Rmoney = row.Field<decimal>("Rmoney"),
                    RemainingPrincipal = row.Field<decimal?>("RemainingPrincipal"),
                    DelayDay = row.Field<int>("DelayDay"),
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
                    }
                    else
                    {
                        decimal RemainingPrincipal_1 = 0;
                        //第一期 數字
                        for (int i = item.loan_grace_num + 1; i <= item.RC_count; i++)
                        {
                            if (i == item.loan_grace_num + 1)
                            {
                                RemainingPrincipal_1 = Math.Round(item.amount_total, 2);
                            }
                            decimal monthlyInterestRate = interest_rate_pass / 100 / 12;
                            item.interest = Math.Round(Math.Round(RemainingPrincipal_1, 2) * monthlyInterestRate, 2);
                            item.Rmoney = Math.Round(Math.Round(realRCAmount, 2) - item.interest, 2);
                            item.RemainingPrincipal = Math.Round(RemainingPrincipal_1 - item.Rmoney, 2);
                            RemainingPrincipal_1 = Convert.ToDecimal(item.RemainingPrincipal);
                        }
                    }
                    item.interest = Math.Round(item.interest, 0, MidpointRounding.AwayFromZero);
                    item.Rmoney = Math.Round(item.Rmoney, 0, MidpointRounding.AwayFromZero);
                    item.RemainingPrincipal = Math.Round(Convert.ToDecimal(item.RemainingPrincipal), 0, MidpointRounding.AwayFromZero);
                }

                var excelList = result.Select(x => new Receivable_Late_Pay_Excel
                {
                    CS_name = x.CS_name,
                    amount_total = x.amount_total,
                    month_total = x.month_total,
                    RC_count = x.RC_count,
                    RC_date = FuncHandler.ConvertGregorianToROC(x.RC_date),
                    RC_amount = x.RC_amount,
                    interest = x.interest,
                    Rmoney = x.Rmoney,
                    RemainingPrincipal = Convert.ToDecimal(x.RemainingPrincipal),
                    DelayDay = x.DelayDay
                }).ToList();

                var Excel_Headers = new Dictionary<string, string>
                {
                    {"index","序號" },
                    { "CS_name", "客戶姓名" },
                    { "amount_total", "總金額" },
                    { "month_total", "期數" },
                    { "RC_count", "第幾期" },
                    { "RC_date", "本期繳款日" },
                    { "RC_amount", "月付金" },
                    { "interest", "利息" },
                    { "Rmoney", "償還本金" },
                    { "RemainingPrincipal", "本金餘額" },
                    { "DelayDay", "延滯天數" }
                };

                var fileBytes = FuncHandler.ReceivableLatePayExcel(excelList, Excel_Headers);
                var fileName = "逾期繳款" + DateTime.Now.ToString("yyyyMMddHHmm") + ".xlsx";
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

        #region 應收帳款-逾放比
        /// <summary>
        /// 逾放比明細資料查詢 RC_Over_Rel_Detail_LQuery/_Ajaxhandler.asp?method=GetOvDtl
        /// </summary>
        /// <param name="overDay">60</param>
        /// <param name="planNum">K0051</param>
        /// <returns></returns>
        [HttpGet("RC_Over_Rel_Detail_LQuery")]
        public ActionResult<ResultClass<string>> RC_Over_Rel_Detail_LQuery(int overDay, string planNum)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"SELECT RC_count,HA.CS_name,isnull(DATEDIFF(DAY,RC_date,CASE WHEN check_pay_date IS NULL THEN SYSDATETIME() ELSE check_pay_date END),0) DelayDay, 
                    RM.amount_total,RM.month_total,RM.amount_per_month 
                    FROM Receivable_D RD 
                    LEFT JOIN Receivable_M RM ON RM.RCM_id = RD.RCM_id AND RM.del_tag='0' 
                    LEFT JOIN House_apply HA ON HA.HA_id = RM.HA_id AND HA.del_tag='0' 
                    LEFT JOIN (SELECT U_num,U_BC FROM User_M) User_M ON User_M.U_num = HA.plan_num 
                    LEFT JOIN House_sendcase HS ON HS.HS_id = RM.HS_id AND HS.del_tag='0' WHERE RD.del_tag = '0' 
                    AND RD.RC_date <= SYSDATETIME() AND (RD.bad_debt_type='N') 
                    AND (RD.cancel_type='N') AND RM.RCM_id IS NOT NULL 
                    AND check_pay_type='N' AND RD.check_pay_date IS NULL/*未繳款*/ 
                    AND (RD.check_pay_date IS NULL OR RD.check_pay_date >= SYSDATETIME()) 
                    AND isnull(DATEDIFF(DAY,RC_date,CASE WHEN check_pay_date IS NULL THEN SYSDATETIME() ELSE check_pay_date END), 0) > @overDay
                    AND convert(varchar(15), RD.RCM_ID)+'-'+ convert(varchar(3), (RC_count)) 
                    in (
                         SELECT convert(varchar(15),RD.RCM_ID)+'-'+ convert(varchar(3),Min(RC_count)) RC_count 
                          FROM Receivable_D RD 
                          LEFT JOIN Receivable_M RM ON RM.RCM_id = RD.RCM_id AND RM.del_tag='0' 
                          LEFT JOIN House_apply ON HA.HA_id = RM.HA_id AND HA.del_tag='0' 
                          LEFT JOIN (SELECT U_num,U_BC FROM User_M) User_M ON User_M.U_num = HA.plan_num 
                          LEFT JOIN House_sendcase HS ON HS.HS_id = RM.HS_id AND HS.del_tag='0' 
                          WHERE RD.del_tag = '0' AND RD.RC_date <= SYSDATETIME() AND (RD.bad_debt_type='N') 
                          AND (RD.cancel_type='N') AND RM.RCM_id IS NOT NULL 
                          AND check_pay_type='N' AND RD.check_pay_date IS NULL/*未繳款*/ 
                          AND (RD.check_pay_date IS NULL OR RD.check_pay_date >= SYSDATETIME() ) 
                          AND isnull(DATEDIFF(DAY, RC_date, CASE WHEN check_pay_date IS NULL THEN SYSDATETIME() ELSE check_pay_date END),0) > @overDay
                          GROUP BY RD.RCM_ID
                       ) and plan_num=@plan_num
                    ORDER BY RD.RC_date,CS_name ";
                parameters.Add(new SqlParameter("@overDay", overDay));
                parameters.Add(new SqlParameter("@plan_num", planNum));
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
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass);
            }
        }
        /// <summary>
        /// 逾放比查詢 RC_Over_Rel_LQuery/RC_over_release.asp
        /// </summary>
        /// <param name="overDay">60</param>
        /// <returns></returns>
        [HttpGet("RC_Over_Rel_LQuery")]
        public ActionResult<ResultClass<string>> RC_Over_Rel_LQuery(int overDay)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"SELECT BC_name,Tol.*,isnull(OV.OV_Count,0) OV_Count,isnull(OV_total, 0) OV_total,M.u_name,
                    FORMAT(ROUND(isnull(OV_total, 0)/isnull(Tol.amount_total, 0)*100,2),'N2') + '%' OV_Rate 
                    FROM (
                           SELECT plan_num,count(plan_num) TOT_Count,sum(amount_total) amount_total    
                    	   FROM ( 
                    	          SELECT HA.plan_num,DATEDIFF(DAY,RD.RC_date,SYSDATETIME()) DiffDay,RM.amount_total 
                    			  FROM (
                    			          SELECT RCM_id,min(RC_count) RC_count,min(RC_date) RC_date 
                    					  FROM Receivable_D 			
                    					  WHERE del_tag = '0' AND check_pay_type='N' AND bad_debt_type='N' AND cancel_type='N' GROUP BY RCM_id 		
                    				   ) RD 		
                    			  LEFT JOIN Receivable_M RM ON RM.RCM_id = RD.RCM_id AND RM.del_tag='0' 		
                    			  LEFT JOIN House_apply HA ON HA.HA_id = RM.HA_id AND HA.del_tag='0' 		
                    			  WHERE RM.RCM_id IS NOT NULL
                    		    ) OV GROUP BY plan_num 
                    	) Tol 
                    LEFT JOIN ( 
                                 SELECT plan_num,count(plan_num) OV_Count,sum(amount_total) OV_total 
                    			 FROM (
                    			        SELECT HA.plan_num,DATEDIFF(DAY,RD.RC_date,SYSDATETIME()) DiffDay,RM.amount_total 
                    					FROM ( 
                    					        SELECT RCM_id,min(RC_count) RC_count,min(RC_date) RC_date FROM Receivable_D 			
                    							WHERE del_tag = '0' AND check_pay_type='N' AND bad_debt_type='N' AND cancel_type='N' GROUP BY RCM_id 		 
                    						 ) RD 			
                    				    LEFT JOIN Receivable_M RM ON RM.RCM_id = RD.RCM_id AND RM.del_tag='0' 			
                    				    LEFT JOIN House_apply HA ON HA.HA_id = RM.HA_id AND HA.del_tag='0' 			
                    				    WHERE RM.RCM_id IS NOT NULL AND (DATEDIFF(DAY,RD.RC_date,SYSDATETIME()) > @overDay) 	  
                    				 ) OV GROUP BY plan_num    
                    		  ) OV ON Tol.plan_num=OV.plan_num  
                    Left Join User_M M on Tol.plan_num=M.u_num    
                    LEFT JOIN (
                    		    SELECT a.item_D_name BC_name,a.item_D_code FROM Item_list a WHERE a.item_M_code = 'branch_company' AND a.item_D_type='Y'
                    		  ) U on M.U_BC =U.item_D_code  
                    where isnull(OV.OV_Count, 0) <> 0  ORDER BY M.U_BC,ROUND(isnull(OV_total, 0)/isnull(Tol.amount_total, 0)*100, 2) desc,Tol.plan_num";
                parameters.Add(new SqlParameter("@overDay", overDay));
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
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass);
            }
        }
        /// <summary>
        /// 逾放比查詢Excel下載 RC_Over_Rel_Excel/RC_over_release.asp
        /// </summary>
        [HttpPost("RC_Over_Rel_Excel")]
        public IActionResult RC_Over_Rel_Excel(int overDay)
        {
            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"SELECT BC_name,Tol.*,isnull(OV.OV_Count,0) OV_Count,isnull(OV_total, 0) OV_total,M.u_name,
                    FORMAT(ROUND(isnull(OV_total, 0)/isnull(Tol.amount_total, 0)*100,2),'N2') + '%' OV_Rate 
                    FROM (
                           SELECT plan_num,count(plan_num) TOT_Count,sum(amount_total) amount_total    
                    	   FROM ( 
                    	          SELECT HA.plan_num,DATEDIFF(DAY,RD.RC_date,SYSDATETIME()) DiffDay,RM.amount_total 
                    			  FROM (
                    			          SELECT RCM_id,min(RC_count) RC_count,min(RC_date) RC_date 
                    					  FROM Receivable_D 			
                    					  WHERE del_tag = '0' AND check_pay_type='N' AND bad_debt_type='N' AND cancel_type='N' GROUP BY RCM_id 		
                    				   ) RD 		
                    			  LEFT JOIN Receivable_M RM ON RM.RCM_id = RD.RCM_id AND RM.del_tag='0' 		
                    			  LEFT JOIN House_apply HA ON HA.HA_id = RM.HA_id AND HA.del_tag='0' 		
                    			  WHERE RM.RCM_id IS NOT NULL
                    		    ) OV GROUP BY plan_num 
                    	) Tol 
                    LEFT JOIN ( 
                                 SELECT plan_num,count(plan_num) OV_Count,sum(amount_total) OV_total 
                    			 FROM (
                    			        SELECT HA.plan_num,DATEDIFF(DAY,RD.RC_date,SYSDATETIME()) DiffDay,RM.amount_total 
                    					FROM ( 
                    					        SELECT RCM_id,min(RC_count) RC_count,min(RC_date) RC_date FROM Receivable_D 			
                    							WHERE del_tag = '0' AND check_pay_type='N' AND bad_debt_type='N' AND cancel_type='N' GROUP BY RCM_id 		 
                    						 ) RD 			
                    				    LEFT JOIN Receivable_M RM ON RM.RCM_id = RD.RCM_id AND RM.del_tag='0' 			
                    				    LEFT JOIN House_apply HA ON HA.HA_id = RM.HA_id AND HA.del_tag='0' 			
                    				    WHERE RM.RCM_id IS NOT NULL AND (DATEDIFF(DAY,RD.RC_date,SYSDATETIME()) > @overDay) 	  
                    				 ) OV GROUP BY plan_num    
                    		  ) OV ON Tol.plan_num=OV.plan_num  
                    Left Join User_M M on Tol.plan_num=M.u_num    
                    LEFT JOIN (
                    		    SELECT a.item_D_name BC_name,a.item_D_code FROM Item_list a WHERE a.item_M_code = 'branch_company' AND a.item_D_type='Y'
                    		  ) U on M.U_BC =U.item_D_code  
                    where isnull(OV.OV_Count, 0) <> 0  ORDER BY M.U_BC,ROUND(isnull(OV_total, 0)/isnull(Tol.amount_total, 0)*100, 2) desc,Tol.plan_num";
                parameters.Add(new SqlParameter("@overDay", overDay));
                #endregion
                var excelList = _adoData.ExecuteQuery(T_SQL, parameters).AsEnumerable().Select(row => new Receivable_Over_Rel_Excel
                {
                    BC_name = row.Field<string>("BC_name"),
                    u_name = row.Field<string>("BC_name"),
                    ToT_Count = row.Field<int>("ToT_Count"),
                    amount_total = row.Field<decimal>("amount_total"),
                    OV_Count = row.Field<int>("OV_Count"),
                    OV_total = row.Field<decimal>("OV_total"),
                    OV_Rate = row.Field<string>("OV_Rate")
                }).ToList();

                var Excel_Headers = new Dictionary<string, string>
                {
                    {"index","序號" },
                    { "BC_name", "客戶姓名" },
                    { "u_name", "總金額" },
                    { "ToT_Count", "期數" },
                    { "amount_total", "第幾期" },
                    { "OV_Count", "本期繳款日" },
                    { "OV_total", "月付金" },
                    { "OV_Rate", "利息" }
                };

                var fileBytes = FuncHandler.ReceivableOverRelExcel(excelList, Excel_Headers);
                var fileName = "逾放比" + DateTime.Now.ToString("yyyyMMddHHmm") + ".xlsx";
                return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            }
            catch (Exception ex)
            {
                ResultClass<string> resultClass = new ResultClass<string>();
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass);
            }
        }
        /// <summary>
        /// 各區逾放比明細資料查詢 RC_Over_Rel_Area_LQuery/RC_over_release.asp
        /// </summary>
        [HttpGet("RC_Over_Rel_Area_LQuery")]
        public ActionResult<ResultClass<string>> RC_Over_Rel_Area_LQuery(int overDay)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"SELECT BC_name,Tol.*,isnull(OV.OV_Count, 0)OV_Count,isnull(OV_total, 0)OV_total,
                    FORMAT(ROUND(isnull(OV_total, 0)/isnull(Tol.amount_total, 0)*100, 2),'N2') + '%' OV_Rate
                    FROM (
                            SELECT U_BC,count(U_BC)TOT_Count,sum(amount_total)amount_total  
                            FROM (
                    	           SELECT M.U_BC,DATEDIFF(DAY, RD.RC_date, SYSDATETIME()) DiffDay,RM.amount_total  
                                   FROM (
                    		              SELECT RCM_id,min(RC_count)RC_count,min(RC_date)RC_date  
                                          FROM Receivable_D  
                                          WHERE del_tag = '0' AND check_pay_type='N' AND bad_debt_type='N' AND cancel_type='N'  
                                          GROUP BY RCM_id
                    				    ) RD  
                                   LEFT JOIN Receivable_M RM ON RM.RCM_id = RD.RCM_id  
                                   LEFT JOIN House_sendcase HS ON RM.HA_id = HS.HA_id AND RM.HS_id = HS.HS_id LEFT JOIN House_apply HA ON HS.HA_id = HA.HA_id   
                     	           LEFT JOIN User_M M ON HA.plan_num=M.u_num WHERE RM.RCM_id IS NOT NULL AND RM.del_tag='0' AND HA.del_tag='0'   
                    	         ) OV  
                            GROUP BY U_BC
                    	  ) Tol  
                    LEFT JOIN (   
                     	        SELECT U_BC,count(U_BC) OV_Count,sum(amount_total) OV_total 
                    			FROM (
                    			       SELECT M.U_BC,DATEDIFF(DAY, RD.RC_date, SYSDATETIME()) DiffDay,RM.amount_total 
                    				   FROM (
                    					      SELECT RCM_id,min(RC_count)RC_count,min(RC_date)RC_date  
                                              FROM Receivable_D  
                                              WHERE del_tag = '0' AND check_pay_type='N' AND bad_debt_type='N' AND cancel_type='N'  
                                              GROUP BY RCM_id
                    						) RD  
                                       LEFT JOIN Receivable_M RM ON RM.RCM_id = RD.RCM_id AND RM.del_tag='0'  
                                       LEFT JOIN House_apply HA ON HA.HA_id = RM.HA_id AND HA.del_tag='0'  
                     	               LEFT JOIN User_M M ON HA.plan_num=M.u_num WHERE RM.RCM_id IS NOT NULL 
                                       AND (DATEDIFF(DAY,RD.RC_date,SYSDATETIME()) > @overDay ) 
                    			     ) OV  
                                GROUP BY U_BC
                    		  ) OV ON Tol.U_BC=OV.U_BC   
                    LEFT JOIN (
                    		    SELECT a.item_D_name BC_name,a.item_D_code  
                                FROM Item_list a   
                                WHERE a.item_M_code = 'branch_company' AND a.item_D_type='Y'  
                              ) U on Tol.U_BC =U.item_D_code  
                    ORDER BY Tol.U_BC, ROUND(isnull(OV_total, 0)/isnull(Tol.amount_total, 0)*100, 2) DESC ";
                parameters.Add(new SqlParameter("@overDay", overDay));
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
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass);
            }
        }
        /// <summary>
        /// 各專案逾放比明細資料查詢 RC_Over_Rel_Area_LQuery/RC_over_release.asp
        /// </summary>
        [HttpGet("RC_Over_Rel_Case_LQuery")]
        public ActionResult<ResultClass<string>> RC_Over_Rel_Case_LQuery(int overDay)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"SELECT Tol.*,isnull(OV.OV_Count, 0)OV_Count,isnull(OV_total, 0)OV_total,
                    FORMAT(ROUND(isnull(OV_total, 0)/isnull(Tol.amount_total, 0)*100, 2),'N2') + '%' OV_Rate								  
                    FROM (
                           SELECT pro_name,count(pro_name)TOT_Count,sum(amount_total)amount_total													  
                           FROM	(
                    	          SELECT pro_name,DATEDIFF(DAY, RD.RC_date, SYSDATETIME()) DiffDay,RM.amount_total																				  
                                  FROM (
                    			         SELECT RCM_id,min(RC_count)RC_count,min(RC_date)RC_date											  
                                         FROM Receivable_D																					  
                                         WHERE del_tag = '0'AND check_pay_type='N' AND bad_debt_type='N'AND cancel_type='N'					  
                                         GROUP BY RCM_id
                    				   ) RD																				  
                                  LEFT JOIN Receivable_M RM ON RM.RCM_id = RD.RCM_id													  
                                  LEFT JOIN House_sendcase HS ON RM.HA_id = HS.HA_id and  RM.HS_id = HS.HS_id 							  
                                  LEFT JOIN House_apply HA ON HS.HA_id = HA.HA_id														  
                     	          LEFT JOIN (																										  
                     			              SELECT HP_project_id,P.project_title,item_D_name pro_name 
                    						  FROM House_pre_project P				  
                     			              left join (																								  
                     				                      SELECT item_D_code, item_D_name 
                    									  FROM Item_list												  
                     				                      WHERE item_M_code = 'project_title' AND item_D_type='Y'										  
                     			                        ) I on case when P.project_title='PJ00001' then 'PJ00005'else P.project_title end = I.item_D_code WHERE P.del_tag='0'										  
                     	                    ) P on HS.HP_project_id=P.HP_project_id																  
                                   WHERE RM.RCM_id IS NOT NULL AND RM.del_tag='0' AND HA.del_tag='0'  AND HS.del_tag='0'	
                    			 ) OV									  
                           GROUP BY pro_name) Tol																					  
                    LEFT JOIN (																			
                     	        SELECT pro_name,count(pro_name)OV_Count,sum(amount_total)OV_total										  
                                FROM (
                    			       SELECT DATEDIFF(DAY, RD.RC_date, SYSDATETIME()) DiffDay,RM.amount_total,pro_name						  
                                       FROM	(
                    				          SELECT RCM_id,min(RC_count)RC_count,min(RC_date)RC_date											  
                                              FROM Receivable_D																					  
                                              WHERE del_tag = '0'AND check_pay_type='N' AND bad_debt_type='N'									  
                     		                  AND cancel_type='N' GROUP BY RCM_id
                    						) RD															  
                                　　　　LEFT JOIN Receivable_M RM ON RM.RCM_id = RD.RCM_id													  
                     	        　　　　LEFT JOIN House_sendcase HS ON RM.HA_id = HS.HA_id and  RM.HS_id = HS.HS_id 							  
                                　　　　LEFT JOIN House_apply HA ON HS.HA_id = HA.HA_id														  
                                　　　　LEFT JOIN (																										  
                     			            　　　  SELECT HP_project_id,P.project_title,item_D_name pro_name FROM House_pre_project P				  
                     			            　　　  left join (																								  
                     				                           SELECT item_D_code, item_D_name FROM Item_list 												  
                     				                           WHERE item_M_code = 'project_title' AND item_D_type='Y'										  
                     			                             ) I on case when P.project_title='PJ00001' then 'PJ00005'else P.project_title end = I.item_D_code WHERE P.del_tag='0'										  
                     	                         ) P on HS.HP_project_id=P.HP_project_id																  
                    　　　　　　　       WHERE RM.RCM_id IS NOT NULL AND RM.del_tag='0' AND HA.del_tag='0' AND HS.del_tag='0'									  
                    　　　　　　　       AND (DATEDIFF(DAY, RD.RC_date, SYSDATETIME()) > 60 )
                    　　　　　        ) OV 
                                group by pro_name) OV ON Tol.pro_name=OV.pro_name																				   
                    ORDER BY Tol.pro_name,ROUND(isnull(OV_total, 0)/isnull(Tol.amount_total, 0)*100, 2) DESC";
                parameters.Add(new SqlParameter("@overDay", overDay));
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
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass);
            }
        }
        #endregion

        #region 應收帳款-本金餘額比
        /// <summary>
        /// 提供應繳查詢年月 GetSendcaseYYYMM/RC_repay.asp
        /// </summary>
        [HttpGet("GetRDYYYMM")]
        public ActionResult<ResultClass<string>> GetRDYYYMM()
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                Receivable_ROC_YYYMM_SE resultData = new Receivable_ROC_YYYMM_SE();
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters_s = new List<SqlParameter>();
                var parameters_e = new List<SqlParameter>();
                var T_SQL_S = @"select distinct convert(varchar(4),(convert(varchar(4),RC_date,126)-1911)) + '-' 
                    + convert(varchar(2),month(RC_date)) yyyMM,convert(varchar(7),D.RC_date,126) yyyymm 
                    from Receivable_D D where (check_pay_type='Y' OR check_pay_type='S' AND check_pay_date IS NOT NULL) 
                    and D.RC_date > DATEADD(month,-12,SYSDATETIME()) 
                    order by convert(varchar(7),D.RC_date,126) desc";
                var T_SQL = @"select distinct top 1 convert(varchar(7),RC_date,126) yyyymm 
                    from Receivable_D 
                    where check_pay_type='Y' 
                    order by convert(varchar(7),RC_date,126) desc";
                var dateM = _adoData.ExecuteSQuery(T_SQL).AsEnumerable().Select(x => x.Field<string>("yyyymm")).First();
                var T_SQL_E = @"select CAST(year(DATEADD(month,3,convert(datetime,@dateM + '-01'))) - 1911 as varchar) + '-' 
                    + CAST(month(DATEADD(month,4,convert(datetime,@dateM + '-01'))) as varchar) yyyMM,
                    convert(varchar(7),DATEADD(month,4,convert(datetime,@dateM + '-01')),126) yyyymm 
                    union all 
                    select CAST(year(DATEADD(month,3,convert(datetime,@dateM + '-01'))) - 1911 as varchar) + '-' 
                    + CAST(month(DATEADD(month,3,convert(datetime,@dateM + '-01'))) as varchar) yyyMM,
                    convert(varchar(7),DATEADD(month,3,convert(datetime,@dateM + '-01')),126) yyyymm 
                    union all 
                    select CAST(year(DATEADD(month,2,convert(datetime,@dateM + '-01'))) - 1911 as varchar) + '-' 
                    + CAST(month(DATEADD(month,2,convert(datetime,@dateM + '-01'))) as varchar) yyyMM,
                    convert(varchar(7),DATEADD(month,2, convert(datetime,@dateM + '-01')),126)yyyymm 
                    union all
                    select CAST(year(DATEADD(month,1, convert(datetime,@dateM + '-01')))-1911 as varchar) + '-' 
                    + CAST(month(DATEADD(month,1,convert(datetime,@dateM + '-01'))) as varchar) yyyMM,
                    convert(varchar(7),DATEADD(month,1,convert(datetime,@dateM + '-01')),126) yyyymm
                    union all
                    select distinct convert(varchar(4),(convert(varchar(4),RC_date,126)-1911)) + '-' 
                    + convert(varchar(2),month(RC_date)) yyyMM,convert(varchar(7),D.RC_date,126) yyyymm 
                    from Receivable_D D where (check_pay_type='Y' OR check_pay_type='S' AND check_pay_date IS NOT NULL) 
                    and D.RC_date > DATEADD(month,-12,SYSDATETIME())";
                parameters_e.Add(new SqlParameter("@dateM", dateM));
                #endregion
                var dateS = _adoData.ExecuteSQuery(T_SQL_S).AsEnumerable().Select(row => new Receivable_ROC_YYYMM_S
                {
                    ROC_YYYMM = row.Field<string>("yyyMM"),
                    Gre_YYYYMM = row.Field<string>("yyyymm")
                }).ToList();
                var dateE = _adoData.ExecuteQuery(T_SQL_E, parameters_e).AsEnumerable().Select(row => new Receivable_ROC_YYYMM_E
                {
                    ROC_YYYMM = row.Field<string>("yyyMM"),
                    Gre_YYYYMM = row.Field<string>("yyyymm")
                }).OrderByDescending(x => x.Gre_YYYYMM).ToList();

                resultData.ROC_Date_S = dateS;
                resultData.ROC_Date_E = dateE;

                resultClass.ResultCode = "000";
                resultClass.objResult = JsonConvert.SerializeObject(resultData);
                return Ok(resultClass);
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass);
            }
        }
        /// <summary>
        /// 本金餘額表查詢 RC_Repay_LQuery/RC_repay.asp
        /// </summary>
        /// <param name="DateS">2024-08</param>
        /// <param name="DateE">2024-11</param>
        [HttpGet("RC_Repay_LQuery")]
        public ActionResult<ResultClass<string>> RC_Repay_LQuery(string DateS, string DateE)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL_SP = @"exec UpdateRemainingPrincipal";
                var T_SQL = @"SELECT sum(Rmoney_T) - sum(Rmoney) Rmoney_U,sum(interest_T) - sum(interest) interest_U,convert(varchar(7),RC_date,126) RC_date, 
                    convert(varchar(4),(convert(varchar(4),RC_date, 126) - 1911)) + '-' + convert(varchar(2),month(RC_date)) yyyMM,         
                    sum(ToCount) ToCount,sum(YCount) YCount,sum(NCount) NCount,sum(BCount) BCount,sum(SCount) SCount,sum(Rmoney) Rmoney,
                    sum(Rmoney_T) Rmoney_T,sum(interest) interest,sum(interest_T) interest_T,sum(RemainingPrincipal) RemainingPrincipal,                                              
                    sum(S_AMT) S_AMT,sum(RemainingPrincipal_BB) RemainingPrincipal_BB                                                   
                    FROM (                                                                                                                        
                     	   SELECT  [RCM_id],1 ToCount,case when check_pay_Type ='Y' then 1 else 0 end YCount,case when check_pay_Type ='N' then 1 else 0 end NCount     
                     	   ,case when check_pay_Type ='S' then 1 else 0 end SCount,case when check_pay_Type ='B' then 1 else 0 end BCount     
                           ,[cs_name],[check_pay_Type],[check_pay_date],[RC_date],[RC_count],RemainingPrincipal,[interest] interest_T,[Rmoney] Rmoney_T                                                                                           
                     	   ,[interest] - (case when [check_pay_Type] ='Y'then 0 when [check_pay_Type] ='N' then [interest] else 0 end) interest    
                           ,[Rmoney]-(case when [check_pay_Type]='Y'then 0 when [check_pay_Type]='N' then [Rmoney] else 0 end) Rmoney          
                           ,[bad_debt_date],[S_AMT],case when check_pay_Type ='B' then S_AMT else 0 end RemainingPrincipal_BB                  
                           FROM [dbo].[view_ACC_Receivable]                                                                                 
                         ) S  
                    where convert(varchar(7), RC_date, 126)between @DateS and @DateE
                    GROUP BY convert(varchar(7), RC_date, 126)                                                                                 
                    ,convert(varchar(4),(convert(varchar(4), RC_date, 126)-1911))+'-'+convert(varchar(2), month(RC_date))                      
                    ORDER BY convert(varchar(7), RC_date, 126)";
                parameters.Add(new SqlParameter("@DateS", DateS));
                parameters.Add(new SqlParameter("@DateE", DateE));
                #endregion
                //優先執行資料更新
                _adoData.ExecuteSQuery(T_SQL_SP);
                var resultList = _adoData.ExecuteQuery(T_SQL, parameters).AsEnumerable().Select(row => new Receivable_Repay_res
                {
                    RC_date = row.Field<string>("RC_date"),
                    yyyMM = row.Field<string>("yyyMM"),
                    ToCount = row.Field<int>("ToCount"),
                    YCount = row.Field<int>("YCount"),
                    NCount = row.Field<int>("NCount"),
                    BCount = row.Field<int>("BCount"),
                    SCount = row.Field<int>("SCount"),
                    str_interest_T = row.Field<decimal>("interest_T").ToString("N0"),
                    str_interest = row.Field<decimal>("interest").ToString("N0"),
                    str_interest_U = row.Field<decimal>("interest_U").ToString("N0"),
                    str_Rmoney_T = row.Field<decimal>("Rmoney_T").ToString("N0"),
                    str_Rmoney = row.Field<decimal>("Rmoney").ToString("N0"),
                    str_Rmoney_U = row.Field<decimal>("Rmoney_U").ToString("N0"),
                    str_RemainingPrincipal_BB = row.Field<decimal>("RemainingPrincipal_BB").ToString("N0"),
                    str_S_AMT = row.Field<decimal>("S_AMT").ToString("N0"),
                    str_RemainingPrincipal = row.Field<decimal>("RemainingPrincipal").ToString("N0"),
                }).ToList();

                resultClass.ResultCode = "000";
                resultClass.objResult = JsonConvert.SerializeObject(resultList);
                return Ok(resultClass);
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass);
            }
        }
        /// <summary>
        /// 本金餘額表查詢Excel下載 RC_Repay_LQuery/RC_repay.asp
        /// </summary>
        [HttpPost("RC_Repay_Excel")]
        public IActionResult RC_Repay_Excel(string DateS, string DateE)
        {
            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL_SP = @"exec UpdateRemainingPrincipal";
                var T_SQL = @"SELECT sum(Rmoney_T) - sum(Rmoney) Rmoney_U,sum(interest_T) - sum(interest) interest_U,convert(varchar(7),RC_date,126) RC_date, 
                    convert(varchar(4),(convert(varchar(4),RC_date, 126) - 1911)) + '-' + convert(varchar(2),month(RC_date)) yyyMM,         
                    sum(ToCount) ToCount,sum(YCount) YCount,sum(NCount) NCount,sum(BCount) BCount,sum(SCount) SCount,sum(Rmoney) Rmoney,
                    sum(Rmoney_T) Rmoney_T,sum(interest) interest,sum(interest_T) interest_T,sum(RemainingPrincipal) RemainingPrincipal,                                              
                    sum(S_AMT) S_AMT,sum(RemainingPrincipal_BB) RemainingPrincipal_BB                                                   
                    FROM (                                                                                                                        
                     	   SELECT  [RCM_id],1 ToCount,case when check_pay_Type ='Y' then 1 else 0 end YCount,case when check_pay_Type ='N' then 1 else 0 end NCount     
                     	   ,case when check_pay_Type ='S' then 1 else 0 end SCount,case when check_pay_Type ='B' then 1 else 0 end BCount     
                           ,[cs_name],[check_pay_Type],[check_pay_date],[RC_date],[RC_count],RemainingPrincipal,[interest] interest_T,[Rmoney] Rmoney_T                                                                                           
                     	   ,[interest] - (case when [check_pay_Type] ='Y'then 0 when [check_pay_Type] ='N' then [interest] else 0 end) interest    
                           ,[Rmoney]-(case when [check_pay_Type]='Y'then 0 when [check_pay_Type]='N' then [Rmoney] else 0 end) Rmoney          
                           ,[bad_debt_date],[S_AMT],case when check_pay_Type ='B' then S_AMT else 0 end RemainingPrincipal_BB                  
                           FROM [dbo].[view_ACC_Receivable]                                                                                 
                         ) S  
                    where convert(varchar(7), RC_date, 126)between @DateS and @DateE
                    GROUP BY convert(varchar(7), RC_date, 126)                                                                                 
                    ,convert(varchar(4),(convert(varchar(4), RC_date, 126)-1911))+'-'+convert(varchar(2), month(RC_date))                      
                    ORDER BY convert(varchar(7), RC_date, 126)";
                parameters.Add(new SqlParameter("@DateS", DateS));
                parameters.Add(new SqlParameter("@DateE", DateE));
                #endregion
                var excelList = _adoData.ExecuteQuery(T_SQL, parameters).AsEnumerable().Select(row => new Receivable_Repay_Excel
                {
                    yyyMM = row.Field<string>("yyyMM"),
                    ToCount = row.Field<int>("ToCount"),
                    YCount = row.Field<int>("YCount"),
                    NCount = row.Field<int>("NCount"),
                    BCount = row.Field<int>("BCount"),
                    SCount = row.Field<int>("SCount"),
                    interest_T = row.Field<decimal>("interest_T"),
                    interest = row.Field<decimal>("interest"),
                    interest_U = row.Field<decimal>("interest_U"),
                    Rmoney_T = row.Field<decimal>("Rmoney_T"),
                    Rmoney = row.Field<decimal>("Rmoney"),
                    Rmoney_U = row.Field<decimal>("Rmoney_U"),
                    RemainingPrincipal_BB = row.Field<decimal>("RemainingPrincipal_BB"),
                    S_AMT = row.Field<decimal>("S_AMT"),
                    RemainingPrincipal = row.Field<decimal>("RemainingPrincipal")
                }).ToList();

                var Excel_Headers = new Dictionary<string, string>
                {
                    {"index","序號" },
                    { "yyyMM", "應繳年月" },
                    { "ToCount", "總件數" },
                    { "YCount", "已繳件數" },
                    { "NCount", "未繳件數" },
                    { "BCount", "呆帳件數" },
                    { "SCount", "提前清償件數" },
                    { "interest_T", "應收利息" },
                    {"interest","已收利息" },
                    {"interest_U","未收利息" },
                    {"Rmoney_T","應償還本金" },
                    {"Rmoney","已償還本金" },
                    {"Rmoney_U","未償還本金" },
                    {"RemainingPrincipal_BB","呆帳金額" },
                    {"S_AMT","提前清償金" },
                    {"RemainingPrincipal","總本金餘額" }
                };

                var fileBytes = FuncHandler.ReceivableRepayExcel(excelList, Excel_Headers);
                var fileName = "本金餘額表" + DateTime.Now.ToString("yyyyMMddHHmm") + ".xlsx";
                return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);

            }
            catch (Exception ex)
            {
                ResultClass<string> resultClass = new ResultClass<string>();
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass);
            }
        }
        /// <summary>
        /// 本金餘額表明細查詢 RC_Repay_Dateil_LQuery/_Ajaxhandler.asp?method=GetRePayDtl
        /// </summary>
        /// <param name="type">N:未繳,B:呆帳,S:提前清償</param>
        /// <param name="date">2024-11</param>
        /// <returns></returns>
        [HttpGet("RC_Repay_Dat-l_LQuery")]
        public ActionResult<ResultClass<string>> RC_Repay_Dateil_LQuery(string type, string date)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();

                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = "";
                switch (type)
                {
                    case "N":
                        T_SQL = @"select cs_name,convert(varchar(10),RC_date,126) RC_date,convert(varchar(10),check_pay_date,126) check_pay_date,RC_count,
                                  FORMAT(interest,'N0'),FORMAT(Rmoney,'N0')
                                  from view_ACC_Receivable 
                                  where check_pay_type='N' AND convert(varchar(7),RC_date, 126) = @date 
                                  order by RC_date";
                        break;
                    case "B":
                        T_SQL = @"select cs_name,convert(varchar(10),bad_debt_date,126) bad_debt_date,FORMAT(RemainingPrincipal_BB,'N0') 
                                  from 
                                  (
                                  SELECT D.RCM_id,check_pay_type,HA_id,B.bad_debt_date,RemainingPrincipal RemainingPrincipal_BB
                                  FROM Receivable_D D
                                  LEFT JOIN Receivable_M M ON M.RCM_id = D.RCM_id
                                  LEFT JOIN (
                                              select distinct RCM_id,bad_debt_date from Receivable_D 
                                              where convert(varchar(7),bad_debt_date,126) = @date
                                  			) B on M.RCM_id =B.RCM_id
                                  WHERE convert(varchar,D.RCM_id) + '-' + convert(varchar,(RCD_id)) 
                                  IN (
                                  	   SELECT convert(varchar,RCM_id) + '-' + convert(varchar,MAX(RCD_id)) RCD_id 
                                  	   FROM Receivable_D 
                                  	   WHERE RCM_id IN (SELECT DISTINCT RCM_id FROM Receivable_D WHERE bad_debt_type = 'Y') 
                                  	   AND bad_debt_type='N' GROUP BY RCM_id
                                  	 )
                                  ) M
                                  LEFT JOIN House_apply HA ON HA.HA_id = M.HA_id
                                  where HA.del_tag='0'";
                        break;
                    case "S":
                        T_SQL = @"select convert(varchar(10),D.RC_date,126) RC_date,convert(varchar(10),D.check_pay_date,126) check_pay_date,cs_name,
                                  A.RCM_id,D.RC_count,A.RCD_id,FORMAT(A.RemainingPrincipal,'N0') S_AMT,
                                  replace(replace(REPLACE(replace(replace(cast(RCM_note as nvarchar(60)),'[',''),']',''),char(10),''),CHAR(13),''),CHAR(32),'') RCM_note   
                                  from 
                                  (/*當期有繳款的資料*/																																																													
                                    select RCM_id,RCD_id,RemainingPrincipal,convert(varchar(7),D.RC_date,126) RC_date,
                                    convert(varchar(7),D.check_pay_date,126) check_pay_date 
                                    from Receivable_D D 
                                    where cast(RCM_id as varchar(20)) + '-' + cast((RCD_id)as varchar(20)) 
                                    in ( SELECT cast(RCM_id as varchar(20)) + '-' + cast(max(RCD_id)as varchar(20)) 
                                  	     from Receivable_D 
                                  	     where RemainingPrincipal is not null and check_pay_type ='S'																				
                                  	     group by RCM_id )
                                    union all	
                                    select D.RCM_id,RCD_id,M.amount_total RemainingPrincipal,convert(varchar(7),D.RC_date,126) RC_date,
                                    convert(varchar(7),D.check_pay_date, 126) check_pay_date 
                                    from Receivable_D D 
                                    left join Receivable_M M   
                                    left join House_apply H on M.HA_id=H.HA_id on D.RCM_id = M.RCM_id  
                                    WHERE check_pay_type ='S' and RC_count=1 and check_pay_date is not null   
                                    union all
                                    select RCM_id,RCD_id,RemainingPrincipal,convert(varchar(7),D.RC_date,126) RC_date,
                                    convert(varchar(7),D.check_pay_date,126) check_pay_date 
                                    from Receivable_D D  
                                    where cast(RCM_id as varchar(20)) + '-' + cast((RCD_id)as varchar(20)) 
                                    in (/*有結清但當期無繳款的資料*/																																																					
                                         SELECT cast(M.RCM_id as varchar(20)) + '-' + cast(max(RCD_id) as varchar(20)) RCD_id 
                                         from Receivable_D D																																			
                                         Left Join (
                                  	                 select * from Receivable_M 
                                  				     where RCM_id 
                                  				     in (
                                                          select M.RCM_id 
                                                          from Receivable_D D																																																
                                                          Left Join (
                                                                      select * from Receivable_M																																																	
                                  　　			　　　                 where RCM_id not in (/*排除結清當期有繳款的資料*/																																																		
                                                                                            select RCM_id from Receivable_D 
                                  														    where cast(RCM_id as varchar(20)) + '-' + cast((RCD_id)as varchar(20)) 
                                                                                            in ( SELECT cast(RCM_id as varchar(20)) + '-' + cast(max(RCD_id)as varchar(20)) 
                                                                                                 from Receivable_D 
                                  　　			　　　                                            where RemainingPrincipal is not null 
                                                                                                 and check_pay_type ='S'																
                                                                                                 group by RCM_id )																																																						
                                                                                          )																																																							
                                                                    ) M on D.RCM_id=M.RCM_id																																																		
                                                          where cast(D.RCM_id as varchar(20)) + '-' + cast((RCD_id)as varchar(20)) 
                                                          in (																																																						
                                                               SELECT cast(RCM_id as varchar(20)) + '-' + cast(max(RCD_id)as varchar(20)) 
                                                               from Receivable_D 
                                                               where RemainingPrincipal is null and check_pay_type ='S'																	
                                                               group by RCM_id ) 
                                                               and M.RCM_id is not null																																																
                                                             )							0																																																				
                                                   ) M on D.RCM_id=M.RCM_id 
                                         where check_pay_type='Y' and D.del_tag = '0' AND M.del_tag='0'																																																			
                                  　　　　group by M.RCM_id 
                                       )
                                  ) A 																																																												
                                  LEFT JOIN Receivable_D D ON A.RCM_id = D.RCM_id and  A.RCD_id = D.RCD_id																																											
                                  LEFT JOIN Receivable_M M ON M.RCM_id = D.RCM_id																																																		
                                  LEFT JOIN House_apply HA ON HA.HA_id = M.HA_id																																																		
                                  LEFT JOIN House_sendcase HS ON HS.HS_id = M.HS_id																																																	
                                  where convert(varchar(7), D.RC_date,126)=@date order by D.RC_date";
                        break;
                }
                parameters.Add(new SqlParameter("@date", date));
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
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass);
            }
        }
        #endregion

        #region 資產數據分析 
        /// <summary>
        /// 資產數據分析查詢 RC_Excess_LQuery/RC_Excess.asp
        /// </summary>
        /// <param name="Forec">1</param>
        [HttpGet("RC_Excess_LQuery")]
        public ActionResult<ResultClass<string>> RC_Excess_LQuery(string? Forec)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parametes = new List<SqlParameter>();
                var T_SQL = "";
                //差在where court_sale = ''條件是否存在
                if (!string.IsNullOrEmpty(Forec) && Forec == "1")
                {
                    T_SQL = @"select V.diffType,AmtTypeDesc,AmtType,count(V.diffType) Count,Format(sum(amount_total),'N0') amount_total,Tot_Amt,rowspan,
                    Format(convert (decimal(5,2),ROUND(sum(amount_total)/Tot_Amt*100,2)),'N2') + '%' Rate
                    from view_excess_base V
                    left join (select sum(isnull(case when try_convert(int, get_amount) is NULL
                        then 0
                        else convert(int, get_amount)
                        end,0)*10000) Tot_Amt
                    from House_sendcase H
                    left join House_apply A ON A.HA_id = H.HA_id
                    where H.del_tag = '0' AND A.del_tag= '0'
                    AND H.sendcase_handle_type= 'Y' AND isnull(H.Send_amount, '') <>''     
        			AND H.fund_company='FDCOM003' AND get_amount_type = 'GTAT002'
                    ) T on 1=1    
                    left join(select diffType, count(diffType) rowspan
                    from (select diffType, AmtType, count(diffType) rowspan
                          from view_excess_base
                          group by diffType, AmtType
                         ) a
                    group by diffType  
                    ) R on V.diffType=R.diffType
                    group by AmtType,AmtTypeDesc,V.diffType,Tot_Amt,rowspan
                    order by V.diffType, AmtType";
                }
                else
                {
                    T_SQL = @"select V.diffType,AmtTypeDesc,AmtType,count(V.diffType) Count,Format(sum(amount_total),'N0') amount_total,Tot_Amt,rowspan,
                    Format(convert (decimal(5,2),ROUND(sum(amount_total)/Tot_Amt*100,2)),'N2') + '%' Rate
                    from view_excess_base V
                    left join (select sum(isnull(case when try_convert(int, get_amount) is NULL
                        then 0
                        else convert(int, get_amount)
                        end,0)*10000) Tot_Amt
                    from House_sendcase H
                    left join House_apply A ON A.HA_id = H.HA_id
                    where H.del_tag = '0' AND A.del_tag= '0'
                    AND H.sendcase_handle_type= 'Y' AND isnull(H.Send_amount, '') <>''     
        			AND H.fund_company='FDCOM003' AND get_amount_type = 'GTAT002'
                    ) T on 1=1    
                    left join (select diffType, count(diffType) rowspan
                    from (select diffType, AmtType, count(diffType) rowspan
                          from view_excess_base
                          where court_sale = ''
                          group by diffType, AmtType
                         ) a
                    group by diffType  
                    ) R on V.diffType=R.diffType
                    where court_sale=''
                    group by AmtType,AmtTypeDesc,V.diffType,Tot_Amt,rowspan
                    order by V.diffType, AmtType";
                }
                #endregion
                DataTable dtResult = _adoData.ExecuteSQuery(T_SQL);
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
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass);
            }
        }
        /// <summary>
        /// 資產數據分析查詢Excel下載 RC_Excess_Excel/RC_Excess.asp
        /// </summary>
        [HttpPost("RC_Excess_Excel")]
        public IActionResult RC_Excess_Excel(string? Forec)
        {
            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parametes = new List<SqlParameter>();
                var T_SQL = "";
                //差在where court_sale = ''條件是否存在
                if (!string.IsNullOrEmpty(Forec) && Forec == "1")
                {
                    T_SQL = @"select V.diffType,AmtTypeDesc,AmtType,count(V.diffType) Count,sum(amount_total) amount_total,Tot_Amt,rowspan,
                    convert (decimal(5,2),ROUND(sum(amount_total)/Tot_Amt*100,2)) Rate
                    from view_excess_base V
                    left join (select sum(isnull(convert(int, get_amount),0)*10000) Tot_Amt
                    from House_sendcase H
                    left join House_apply A ON A.HA_id = H.HA_id
                    where H.del_tag = '0' AND A.del_tag= '0'
                    AND H.sendcase_handle_type= 'Y' AND isnull(H.Send_amount, '') <>''     
        			AND H.fund_company='FDCOM003' AND get_amount_type = 'GTAT002'
                    ) T on 1=1    
                    left join(select diffType, count(diffType) rowspan
                    from (select diffType, AmtType, count(diffType) rowspan
                          from view_excess_base
                          group by diffType, AmtType
                         ) a
                    group by diffType  
                    ) R on V.diffType=R.diffType
                    group by AmtType,AmtTypeDesc,V.diffType,Tot_Amt,rowspan
                    order by V.diffType, AmtType";
                }
                else
                {
                    T_SQL = @"select V.diffType,AmtTypeDesc,AmtType,count(V.diffType) Count,sum(amount_total) amount_total,Tot_Amt,rowspan,
                    convert (decimal(5,2),ROUND(sum(amount_total)/Tot_Amt*100,2)) Rate
                    from view_excess_base V
                    left join (select sum(isnull(convert(int, get_amount),0)*10000) Tot_Amt
                    from House_sendcase H
                    left join House_apply A ON A.HA_id = H.HA_id
                    where H.del_tag = '0' AND A.del_tag= '0'
                    AND H.sendcase_handle_type= 'Y' AND isnull(H.Send_amount, '') <>''     
        			AND H.fund_company='FDCOM003' AND get_amount_type = 'GTAT002'
                    ) T on 1=1    
                    left join (select diffType, count(diffType) rowspan
                    from (select diffType, AmtType, count(diffType) rowspan
                          from view_excess_base
                          where court_sale = ''
                          group by diffType, AmtType
                         ) a
                    group by diffType  
                    ) R on V.diffType=R.diffType
                    where court_sale=''
                    group by AmtType,AmtTypeDesc,V.diffType,Tot_Amt,rowspan
                    order by V.diffType, AmtType";
                }
                #endregion
                var excelList = _adoData.ExecuteSQuery(T_SQL).AsEnumerable().Select(row => new Receivable_Excess_Excel
                {
                    diffType = row.Field<string>("diffType"),
                    AmtTypeDesc = row.Field<string>("AmtTypeDesc"),
                    Count = row.Field<int>("Count"),
                    amount_total = row.Field<decimal>("amount_total"),
                    Rate = row.Field<decimal>("Rate"),
                }).ToList();

                var Excel_Headers = new Dictionary<string, string>
                {
                    { "diffType","逾放期數" },
                    { "AmtTypeDesc", "放款金額" },
                    { "Count", "件數" },
                    { "amount_total", "逾放金額" },
                    { "Rate", "逾放比%" }
                };

                var fileBytes = FuncHandler.ReceivableExcessExcel(excelList, Excel_Headers);
                var fileName = "資產數據分析" + DateTime.Now.ToString("yyyyMMddHHmm") + ".xlsx";
                return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            }
            catch (Exception ex)
            {
                ResultClass<string> resultClass = new ResultClass<string>();
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass);
            }
        }
        /// <summary>
        /// 資產數據分析明細查詢 RC_Excess_Detail_LQuery/_Ajaxhandler.asp?method=GetExcess_base
        /// </summary>
        [HttpPost("RC_Excess_Detail_LQuery")]
        public ActionResult<ResultClass<string>> RC_Excess_Detail_LQuery(Receivable_Excess_req model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = "";
                if(!string.IsNullOrEmpty(model.Forec) && model.Forec == "1")
                {
                    T_SQL = @"SELECT cs_name,DiffDay, amount_total,RC_count,month_total,court_sale,AmtTypeDesc,AmtType,DiffType
                    ,replace(replace(REPLACE(replace(replace(cast(RCM_note as nvarchar(max)),'[',''),']',''),char(10),''),CHAR(13),''),CHAR(32),'') RCM_note
                    FROM dbo.view_excess_base
                    where 1 = 1 
                    and DiffType=@DiffType and  AmtType=@AmtType
                    ORDER BY DiffType,AmtType,DiffDay";
                }
                else
                {
                    T_SQL = @"SELECT cs_name,DiffDay, amount_total,RC_count,month_total,court_sale,AmtTypeDesc,AmtType,DiffType
                    ,replace(replace(REPLACE(replace(replace(cast(RCM_note as nvarchar(max)),'[',''),']',''),char(10),''),CHAR(13),''),CHAR(32),'') RCM_note
                    FROM dbo.view_excess_base
                    where 1 = 1 
                    and DiffType=@DiffType and  AmtType=@AmtType
                    and court_sale=''
                    ORDER BY DiffType,AmtType,DiffDay";
                }
                parameters.Add(new SqlParameter("@DiffType", model.DiffType));
                parameters.Add(new SqlParameter("@AmtType", model.AmtType));
                #endregion
                DataTable dtResult=_adoData.ExecuteQuery(T_SQL, parameters);
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
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass);
            }
        }
        /// <summary>
        /// 資產數據分析查詢明細Excel下載 RC_Excess_Detail_Excel//RC_Excess.asp
        /// </summary>
        [HttpPost("RC_Excess_Detail_Excel")]
        public IActionResult RC_Excess_Detail_Excel(string? Forec)
        {
            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = "";
                if(!string.IsNullOrEmpty(Forec) && Forec == "1")
                {
                    T_SQL = @"SELECT cs_name,DiffDay,amount_total,RC_count,month_total,court_sale,AmtTypeDesc,AmtType,DiffType
                    ,replace(replace(REPLACE(replace(replace(cast(RCM_note as nvarchar(max)),'[',''),']',''),char(10),''),CHAR(13),''),CHAR(32),'') RCM_note
                    FROM dbo.view_excess_base
                    where 1 = 1 
                    ORDER BY DiffType,AmtType,DiffDay";
                }
                else
                {
                    T_SQL = @"SELECT cs_name,DiffDay,amount_total,RC_count,month_total,court_sale,AmtTypeDesc,AmtType,DiffType
                    ,replace(replace(REPLACE(replace(replace(cast(RCM_note as nvarchar(max)),'[',''),']',''),char(10),''),CHAR(13),''),CHAR(32),'') RCM_note
                    FROM dbo.view_excess_base
                    where 1 = 1 
                    and court_sale=''
                    ORDER BY DiffType,AmtType,DiffDay";
                }
                #endregion
                var excelList =_adoData.ExecuteSQuery(T_SQL).AsEnumerable().Select(row=>new Receivable_Excess_Detail_Excel 
                {
                    diffType = row.Field<string>("diffType"),
                    AmtTypeDesc = row.Field<string>("AmtTypeDesc"),
                    Cs_name = row.Field<string>("Cs_name"),
                    DiffDay = row.Field<int>("DiffDay"),
                    amount_total = row.Field<decimal>("amount_total"),
                    RCM_note = row.Field<string>("RCM_note"),
                }).ToList();

                var Excel_Headers = new Dictionary<string, string>
                {
                    { "diffType","逾放期數" },
                    { "AmtTypeDesc", "放款金額" },
                    { "Cs_name", "客戶名稱" },
                    { "DiffDay", "延滯天數" },
                    { "amount_total", "逾放金額" },
                    { "RCM_note", "備註" }
                };

                var fileBytes = FuncHandler.ReceivableExcessDetailExcel(excelList, Excel_Headers);
                var fileName = "資產數據分析(明細資料)" + DateTime.Now.ToString("yyyyMMddHHmm") + ".xlsx";
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

        #region 債權憑證
        /// <summary>
        /// 債權憑證查詢 Debt_Certificate_LQuery/Debt_certificate_list.asp
        /// </summary>
        [HttpGet("Debt_Certificate_LQuery")]
        public ActionResult<ResultClass<string>> Debt_Certificate_LQuery(string? cs_name)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var User_Num = HttpContext.Session.GetString("UserID");
            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"select Debt_ID, cs_name, CS_PID, FORMAT(loan_amount,'N0') loan_amount,certificate_date_S,certificate_date_E,Remark 
                    from Debt_certificate
                    where del_tag = '0'
                    order by certificate_date_S";
                #endregion
                var result = _adoData.ExecuteSQuery(T_SQL).AsEnumerable().Select(row => new Debt_Certificate_Lres
                {
                    Debt_ID = row.Field<int>("Debt_ID"),
                    cs_name = row.Field<string>("cs_name"),
                    CS_PID = row.Field<string>("CS_PID"),
                    str_loan_amount = row.Field<string>("loan_amount"),
                    str_certificate_date_S = FuncHandler.ConvertGregorianToROC(row.Field<DateTime>("certificate_date_S").ToString("yyyy/MM/dd")),
                    str_certificate_date_E = FuncHandler.ConvertGregorianToROC(row.Field<DateTime>("certificate_date_E").ToString("yyyy/MM/dd")),
                    Remark = row.Field<string>("Remark")
                }).ToList();
                if (!string.IsNullOrEmpty(cs_name))
                {
                    result = result.Where(x => x.cs_name.Equals(cs_name)).ToList();
                }

                if(result != null && result.Count > 0)
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
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass);
            }
        }
        /// <summary>
        /// 債權憑證查詢Excel下載 Debt_Certificate_Exccel/Debt_certificate_list.asp
        /// </summary>
        [HttpPost("Debt_Certificate_Exccel")]
        public IActionResult Debt_Certificate_Exccel(string? cs_name)
        {
            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"select cs_name,CS_PID,loan_amount,certificate_date_S,certificate_date_E,Remark 
                    from Debt_certificate
                    where del_tag = '0'
                    order by certificate_date_S";
                #endregion
                var excelList = _adoData.ExecuteSQuery(T_SQL).AsEnumerable().Select(row=>new Debt_Certificate_Excel 
                {
                    cs_name = row.Field<string>("cs_name"),
                    CS_PID = row.Field<string>("CS_PID"),
                    loan_amount = row.Field<decimal>("loan_amount"),
                    str_certificate_date_S = FuncHandler.ConvertGregorianToROC(row.Field<DateTime>("certificate_date_S").ToString("yyyy/MM/dd")),
                    str_certificate_date_E = FuncHandler.ConvertGregorianToROC(row.Field<DateTime>("certificate_date_E").ToString("yyyy/MM/dd")),
                    Remark = row.Field<string>("Remark")
                }).ToList();
                if(!string.IsNullOrEmpty(cs_name))
                {
                    excelList = excelList.Where(x=>x.cs_name.Equals(cs_name)).ToList();
                }

                var Excel_Headers = new Dictionary<string, string>
                {
                    { "cs_name","客戶名稱" },
                    { "CS_PID", "身分證字號" },
                    { "loan_amount", "貸款金額" },
                    { "str_certificate_date_S", "憑證起始時間" },
                    { "str_certificate_date_E", "憑證到期時間" },
                    { "Remark", "備註" }
                };

                var fileBytes = FuncHandler.ExportToExcel(excelList, Excel_Headers);
                var fileName = "資產數據分析(明細資料)" + DateTime.Now.ToString("yyyyMMddHHmm") + ".xlsx";
                return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            }
            catch (Exception ex)
            {
                ResultClass<string> resultClass = new ResultClass<string>();
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass);
            }
        }
        /// <summary>
        /// 債權憑證資料查詢 Debt_Certificate_DetQuery/Debt_certificate_list.asp
        /// </summary>
        [HttpGet("Debt_Certificate_DetQuery")]
        public ActionResult<ResultClass<string>> Debt_Certificate_DetQuery(string DebtID)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"select cs_name,CS_PID,loan_amount,certificate_date_S,certificate_date_E,Remark 
                    from Debt_certificate
                    where del_tag = '0' and Debt_ID = @Debt_ID
                    order by certificate_date_S";
                parameters.Add(new SqlParameter("@Debt_ID", DebtID));
                #endregion
                var result = _adoData.ExecuteQuery(T_SQL,parameters).AsEnumerable().Select(row => new Debt_Certificate_res
                {
                    cs_name = row.Field<string>("cs_name"),
                    CS_PID = row.Field<string>("CS_PID"),
                    loan_amount = row.Field<decimal>("loan_amount"),
                    str_certificate_date_S = FuncHandler.ConvertGregorianToROC(row.Field<DateTime>("certificate_date_S").ToString("yyyy/MM/dd")),
                    str_certificate_date_E = FuncHandler.ConvertGregorianToROC(row.Field<DateTime>("certificate_date_E").ToString("yyyy/MM/dd")),
                    Remark = row.Field<string>("Remark")
                }).ToList();

                if (result != null && result.Count > 0)
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
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass);
            }
        }
        /// <summary>
        /// 債權憑證資料修改 Debt_Certificate_DetUpd/Debt_certificate_list.asp
        /// </summary>
        [HttpPost("Debt_Certificate_DetUpd")]
        public ActionResult<ResultClass<string>> Debt_Certificate_DetUpd(Debt_Certificate_req model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            var User_Num = HttpContext.Session.GetString("UserID");
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"Update Debt_certificate set cs_name = @cs_name,CS_PID = @CS_PID,loan_amount = @loan_amount,
                    certificate_date_S = @certificate_date_S,certificate_date_E = @certificate_date_E,Remark = @Remark, 
                    edit_date = GETDATE(),edit_num = @edit_num,edit_ip = @edit_ip where Debt_ID = @Debt_ID";
                parameters.Add(new SqlParameter("@cs_name", model.cs_name));
                parameters.Add(new SqlParameter("@CS_PID", model.CS_PID));
                parameters.Add(new SqlParameter("@loan_amount", model.loan_amount));
                parameters.Add(new SqlParameter("@certificate_date_S", DateTime.Parse(FuncHandler.ConvertROCToGregorian(model.str_certificate_date_S))));
                parameters.Add(new SqlParameter("@certificate_date_E", DateTime.Parse(FuncHandler.ConvertROCToGregorian(model.str_certificate_date_E))));
                parameters.Add(new SqlParameter("@Remark", model.Remark));
                parameters.Add(new SqlParameter("@edit_num", model.edit_num)); // User_Num
                parameters.Add(new SqlParameter("@edit_ip", clientIp)); // 
                parameters.Add(new SqlParameter("@Debt_ID", model.Debt_ID));
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
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass);
            }
        }
        /// <summary>
        /// 債權憑證資料刪除 Debt_Certificate_DetDel/Debt_certificate_list.asp
        /// </summary>
        [HttpPost("Debt_Certificate_DetDel")]
        public ActionResult<ResultClass<string>> Debt_Certificate_DetDel(string DebtID, string del_num)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            var User_Num = HttpContext.Session.GetString("UserID");
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"Update Debt_certificate set del_tag = '1',del_date = GETDATE(),del_num = @del_num,del_ip = @del_ip where Debt_ID = @Debt_ID";
                parameters.Add(new SqlParameter("@del_num", del_num));
                parameters.Add(new SqlParameter("@del_ip", clientIp));
                parameters.Add(new SqlParameter("@Debt_ID", DebtID));
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
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass);
            }
        }
        /// <summary>
        /// 新增債權憑證資料 Debt_Certificate_DetIns/Debt_certificate_list.asp
        /// </summary>
        [HttpPost("Debt_Certificate_DetIns")]
        public ActionResult<ResultClass<string>> Debt_Certificate_DetIns(Debt_Certificate_req model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            var User_Num = HttpContext.Session.GetString("UserID");
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"Insert into Debt_certificate (cs_name,CS_PID,loan_amount,certificate_date_S,certificate_date_E,Remark,add_date,add_num,add_ip,del_tag)
                    Values (@cs_name,@CS_PID,@loan_amount,@certificate_date_S,@certificate_date_E,@Remark,GETDATE(),@add_num,@add_ip,@del_tag)";
                parameters.Add(new SqlParameter("@cs_name", model.cs_name));
                parameters.Add(new SqlParameter("@CS_PID", model.CS_PID));
                parameters.Add(new SqlParameter("@loan_amount", model.loan_amount));
                parameters.Add(new SqlParameter("@certificate_date_S", DateTime.Parse(FuncHandler.ConvertROCToGregorian(model.str_certificate_date_S))));
                parameters.Add(new SqlParameter("@certificate_date_E", DateTime.Parse(FuncHandler.ConvertROCToGregorian(model.str_certificate_date_E))));
                parameters.Add(new SqlParameter("@Remark", model.Remark));
                parameters.Add(new SqlParameter("@add_num", model.add_num));
                parameters.Add(new SqlParameter("@add_ip",clientIp));
                parameters.Add(new SqlParameter("@del_tag", "0"));
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
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass);
            }
        }
        #endregion
    }
}
