using KF_WebAPI.BaseClass;
using KF_WebAPI.BaseClass.AE;
using KF_WebAPI.FunctionHandler;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Newtonsoft.Json;
using OfficeOpenXml.Sorting;
using System;
using System.Data;
using System.Text;

namespace KF_WebAPI.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class AE_SALController : ControllerBase
    {
        /// <summary>
        /// 車貸列表查詢 House_othercase_LQuery
        /// </summary>
        [HttpPost("House_othercase_LQuery")]
        public ActionResult<ResultClass<string>> House_othercase_LQuery(House_othercase_Req model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            try
            {
                ADOData _adoData = new ADOData(); // 測試:"Test" / 正式:""

                var sqlBuilder = new StringBuilder("SELECT " +
                    " show_fund_company, show_project_title, cs_name, cs_id, get_amount, period, interest_rate_pass, FORMAT(get_amount_date,'yyyy/MM/dd') as get_amount_date, comm_amt, case_remark, case_id " +
                    ",(select U_name FROM User_M where U_num = House_othercase.plan_num AND del_tag='0') as plan_name " +
                    ",(SELECT COUNT(*) FROM ASP_UpLoad WHERE cknum = House_othercase.case_id and del_tag = '0') AS case_id_count "+
                    " FROM House_othercase WHERE 1 = 1 ");

                var parameters = new List<SqlParameter>();
                
                if(!string.IsNullOrEmpty(model.Date_S) && !string.IsNullOrEmpty(model.Date_E))
                {
                    sqlBuilder.Append(" AND get_amount_date between @Date_S and @Date_E ");
                    parameters.Add(new SqlParameter("@Date_S", FuncHandler.ConvertROCToGregorian(model.Date_S)));
                    parameters.Add(new SqlParameter("@Date_E", FuncHandler.ConvertROCToGregorian(model.Date_E)));
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


        /// <summary>
        /// 廠商資料單筆查詢 Manufacturer_SQuery
        /// </summary>
        [HttpGet("House_othercase_SQuery")]
        public ActionResult<ResultClass<string>> House_othercase_SQuery(string case_id)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            try
            {
                ADOData _adoData = new ADOData(); // 測試:"Test" / 正式:""
                #region SQL
                var T_SQL = @"SELECT " +
                    " show_fund_company, show_project_title, cs_name, cs_id, get_amount, period, interest_rate_pass, FORMAT(get_amount_date,'yyyy/MM/dd') as get_amount_date, comm_amt, case_remark, act_perf_amt, case_id, confirm_num " +
                    " ,plan_num,(select U_name FROM User_M where U_num = House_othercase.plan_num AND del_tag='0') as plan_name " +
                    ",(SELECT COUNT(*) FROM ASP_UpLoad WHERE cknum = House_othercase.case_id and del_tag = '0') AS case_id_count "+
                    " FROM House_othercase WHERE case_id = @ID";
                var parameters = new List<SqlParameter>
                {
                    new SqlParameter("@ID", case_id)
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
        /// 新增車貸資料 House_othercase_Ins
        /// </summary>
        [HttpPost("House_othercase_Ins")]
        public ActionResult<ResultClass<string>> House_othercase_Ins(House_othercase_Ins model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();
            string case_id =  FuncHandler.GetCheckNum();
            case_id = case_id.Substring(0, case_id.Length - 6);

            try
              {
                  ADOData _adoData = new ADOData();
                  #region SQL

                  var T_SQL = @"INSERT INTO [House_othercase]
                       ([case_id]
                       ,[CaseType]
                       ,[show_fund_company]
                       ,[show_project_title]
                       ,[cs_name]
                       ,[cs_id]
                       ,[get_amount]
                       ,[period]
                       ,[interest_rate_pass]
                       ,[get_amount_date]
                       ,[case_remark]
                       ,[comm_amt]
                       ,[act_perf_amt]
                       ,[plan_num]
                       ,[add_date]
                       ,[add_num]
                       ,[add_ip])
                 VALUES
                       (@case_id
                       ,@CaseType
                       ,@show_fund_company
                       ,@show_project_title
                       ,@cs_name
                       ,@cs_id
                       ,@get_amount
                       ,@period
                       ,@interest_rate_pass
                       ,@get_amount_date
                       ,@case_remark
                       ,@comm_amt
                       ,@act_perf_amt
                       ,@plan_num
                       ,@add_date
                       ,@add_num
                       ,@add_ip)";

                  var parameters = new List<SqlParameter>
                  {
                      

                      new SqlParameter("@case_id", case_id),
                      new SqlParameter("@CaseType", string.IsNullOrEmpty(model.CaseType) ? DBNull.Value : model.CaseType),
                      new SqlParameter("@show_fund_company", string.IsNullOrEmpty(model.show_fund_company) ? DBNull.Value : model.show_fund_company),
                      new SqlParameter("@show_project_title", string.IsNullOrEmpty(model.show_project_title) ? DBNull.Value : model.show_project_title),
                      new SqlParameter("@cs_name", string.IsNullOrEmpty(model.cs_name) ? DBNull.Value : model.cs_name),
                      new SqlParameter("@cs_id", string.IsNullOrEmpty(model.cs_id) ? DBNull.Value : model.cs_id),
                      new SqlParameter("@get_amount", string.IsNullOrEmpty(model.get_amount) ? DBNull.Value : model.get_amount),
                      new SqlParameter("@period", string.IsNullOrEmpty(model.period) ? DBNull.Value : model.period),
                      new SqlParameter("@interest_rate_pass", string.IsNullOrEmpty(model.interest_rate_pass) ? DBNull.Value : model.interest_rate_pass),
                      
                      // Datetime Nullable 要參考其他人的做法，
                      new SqlParameter("@get_amount_date", string.IsNullOrEmpty(model.get_amount_date) ? DBNull.Value : model.get_amount_date),
                      
                      new SqlParameter("@case_remark", string.IsNullOrEmpty(model.case_remark) ? DBNull.Value : model.case_remark),
                      // 數值Nullable 要參考其他人的做法，
                      new SqlParameter("@comm_amt", model.comm_amt),
                      new SqlParameter("@act_perf_amt", model.act_perf_amt),

                      new SqlParameter("@plan_num", string.IsNullOrEmpty(model.plan_num) ? DBNull.Value : model.plan_num),
                      new SqlParameter("@add_date", DateTime.Today),
                      new SqlParameter("@add_num", model.add_num),
                      new SqlParameter("@add_ip", clientIp)
                  };
             

                    #endregion
                    int result = _adoData.ExecuteNonQuery(T_SQL, parameters);

                    if (result == 0)
                    {
                        resultClass.ResultCode = "400";
                        resultClass.ResultMsg = "新增失敗";
                        return BadRequest(resultClass);
                    }
                    else
                    {
                        resultClass.ResultCode = "000";
                        resultClass.ResultMsg = "新增成功";
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
        /// 修改廠商資料 House_othercase_Upd
        /// </summary>
        [HttpPost("House_othercase_Upd")]
        public ActionResult<ResultClass<string>> House_othercase_Upd(House_othercase_Ins model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();
            
            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"UPDATE [dbo].[House_othercase]
                               SET 
                                  [show_fund_company] = @show_fund_company
                                  ,[show_project_title] = @show_project_title
                                  ,[cs_name] = @cs_name
                                  ,[cs_id] = @cs_id
                                  ,[get_amount] = @get_amount
                                  ,[period] = @period
                                  ,[interest_rate_pass] = @interest_rate_pass
                                  ,[get_amount_date] = @get_amount_date
                                  ,[case_remark] = @case_remark
                                  ,[comm_amt] = @comm_amt
                                  ,[act_perf_amt] = @act_perf_amt
                                  ,[plan_num] = @plan_num
                                  ,[edit_date] = @edit_date
                                  ,[edit_num] = @edit_num
                                  ,[edit_ip] = @edit_ip
                                where [case_id] = @case_id";
                var parameters = new List<SqlParameter>
                {
                      new SqlParameter("@case_id", model.case_id),
                      new SqlParameter("@show_fund_company", string.IsNullOrEmpty(model.show_fund_company) ? DBNull.Value : model.show_fund_company),
                      new SqlParameter("@show_project_title", string.IsNullOrEmpty(model.show_project_title) ? DBNull.Value : model.show_project_title),
                      new SqlParameter("@cs_name", string.IsNullOrEmpty(model.cs_name) ? DBNull.Value : model.cs_name),
                      new SqlParameter("@cs_id", string.IsNullOrEmpty(model.cs_id) ? DBNull.Value : model.cs_id),
                      new SqlParameter("@get_amount", string.IsNullOrEmpty(model.get_amount) ? DBNull.Value : model.get_amount),
                      new SqlParameter("@period", string.IsNullOrEmpty(model.period) ? DBNull.Value : model.period),
                      new SqlParameter("@interest_rate_pass", string.IsNullOrEmpty(model.interest_rate_pass) ? DBNull.Value : model.interest_rate_pass),
                      
                      // Datetime Nullable 要參考其他人的做法，
                      new SqlParameter("@get_amount_date", string.IsNullOrEmpty(model.get_amount_date) ? DBNull.Value : model.get_amount_date),

                      new SqlParameter("@case_remark", string.IsNullOrEmpty(model.case_remark) ? DBNull.Value : model.case_remark),
                      // 數值Nullable 要參考其他人的做法，
                      new SqlParameter("@comm_amt", model.comm_amt),
                      new SqlParameter("@act_perf_amt", model.act_perf_amt),

                      new SqlParameter("@plan_num", string.IsNullOrEmpty(model.plan_num) ? DBNull.Value : model.plan_num),
                      new SqlParameter("@edit_date", DateTime.Today),
                      new SqlParameter("@edit_num", model.edit_num),
                      new SqlParameter("@edit_ip", clientIp)
                };
                #endregion
                int result = _adoData.ExecuteNonQuery(T_SQL, parameters);

                if (result == 0)
                {
                    resultClass.ResultCode = "400";
                    resultClass.ResultMsg = "修改失敗";
                    return BadRequest(resultClass);
                }
                else
                {
                    resultClass.ResultCode = "000";
                    resultClass.ResultMsg = "修改成功";
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
        /// 刪除單筆 House_othercase_Del
        /// </summary>
        [HttpDelete("House_othercase_Del")]
        public ActionResult<ResultClass<string>> House_othercase_Del(string case_id)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"Delete House_othercase where case_id=@case_id";
                var parameters = new List<SqlParameter>
                {
                    new SqlParameter("@case_id",case_id)
                };
                #endregion
                int result = _adoData.ExecuteNonQuery(T_SQL, parameters);
                if (result == 0)
                {
                    resultClass.ResultCode = "400";
                    resultClass.ResultMsg = "明細刪除失敗";
                    return BadRequest(resultClass);
                }
                else
                {
                    resultClass.ResultCode = "000";
                    resultClass.ResultMsg = "明細刪除成功";
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

    }
}
