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

        # region 車貸
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
        /// 車貸單筆查詢 Manufacturer_SQuery
        /// </summary>
        [HttpGet("House_othercase_SQuery")]
        public ActionResult<ResultClass<string>> House_othercase_SQuery(string Id)
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
                    new SqlParameter("@ID", Id)
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
        /// 修改車貸資料 House_othercase_Upd
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
        /// 刪除車貸單筆 House_othercase_Del
        /// </summary>
        [HttpDelete("House_othercase_Del")]
        public ActionResult<ResultClass<string>> House_othercase_Del(string Id)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"Delete House_othercase where case_id=@Id";
                var parameters = new List<SqlParameter>
                {
                    new SqlParameter("@Id",Id)
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
        #endregion

        #region 委對單
        /// <summary>
        /// 委對單列表查詢 House_agency_LQuery
        /// </summary>
        [HttpPost("House_agency_LQuery")]
        public ActionResult<ResultClass<string>> House_agency_LQuery(House_agency_Req model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            try
            {
                ADOData _adoData = new ADOData(); // 測試:"Test" / 正式:""

                var sqlBuilder = new StringBuilder("select AG_id, FORMAT(add_date,'yyyy/MM/dd') add_date, case_com, agency_com, check_process_type, close_type " +
                    ", (select U_name FROM User_M where U_num = agcy.check_leader_num AND del_tag = '0') as check_leader_name " +
                    ",(select U_name FROM User_M where U_num = agcy.add_num AND del_tag = '0') as add_name " +
                    ",(select U_name FROM User_M where U_num = agcy.check_process_num AND del_tag='0') as check_process_name " +
                    "from House_agency agcy WHERE 1 = 1 ");

                var parameters = new List<SqlParameter>();
                if (model.AG_id != 0)
                {
                    sqlBuilder.Append(" AND AG_id = @AG_id ");
                    parameters.Add(new SqlParameter("@AG_id", model.AG_id));
                }

                if (!string.IsNullOrEmpty(model.Date_S) && !string.IsNullOrEmpty(model.Date_E))
                {
                    sqlBuilder.Append(" AND add_date between @Date_S and @Date_E ");
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
        /// 委對單單筆查詢 Manufacturer_SQuery
        /// </summary>
        [HttpGet("House_agency_SQuery")]
        public ActionResult<ResultClass<string>> House_agency_SQuery(string Id)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            try
            {
                ADOData _adoData = new ADOData(); // 測試:"Test" / 正式:""
                #region SQL
                var T_SQL = @"select AG_id, FORMAT(add_date,'yyyy/MM/dd') add_date, case_com, agency_com, case_text,CS_text,check_date,check_address, pass_amount,set_amount, print_data,get_data,process_charge,AG_note, check_leader_num, check_process_type, check_process_date, check_process_note,check_process_num,set_process_date, close_type " +
                    ", (select U_name FROM User_M where U_num = agcy.check_leader_num AND del_tag = '0') as check_leader_name " +
                    ",(select U_name FROM User_M where U_num = agcy.add_num AND del_tag = '0') as add_name " +
                    ",(select U_name FROM User_M where U_num = agcy.check_process_num AND del_tag='0') as check_process_name " +
                    "from House_agency agcy WHERE AG_id = @ID ";
                var parameters = new List<SqlParameter>
                {
                    new SqlParameter("@ID", Id)
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
        /// 新增委對單資料 House_agency_Ins
        /// </summary>
        [HttpPost("House_agency_Ins")]
        public ActionResult<ResultClass<string>> House_agency_Ins(House_agency_Ins model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();


            try
            {
                ADOData _adoData = new ADOData();
                #region SQL

                var T_SQL = @"INSERT INTO [dbo].[House_agency]
                               ([AG_cknum],[case_com],[agency_com],[case_text],[CS_text],[check_date],[check_address]
                               ,[pass_amount],[set_amount],[print_data],[get_data],[process_charge],[AG_note],[check_leader_num]
                               ,[check_process_num]
                               ,[check_process_type]
                               ,[check_process_note]
                               ,[close_type],[close_type_date],[del_tag],[add_date],[add_num],[add_ip]
                               )
                         VALUES
                               (@AG_cknum, @case_com, @agency_com, @case_text, @CS_text, @check_date, @check_address, 
                               @pass_amount,@set_amount, @print_data,@get_data,@process_charge, @AG_note,  @check_leader_num, 
                               @check_process_num,  
                                @check_process_type, 
                                @check_process_note, 
                               @close_type, @close_type_date, @del_tag, @add_date, @add_num, @add_ip 
                               )";

                var parameters = new List<SqlParameter>
                {
                    new SqlParameter("@AG_cknum",  FuncHandler.GetCheckNum()),
                    new SqlParameter("@case_com",  string.IsNullOrEmpty(model.case_com) ? DBNull.Value : model.case_com),
                    new SqlParameter("@agency_com",  string.IsNullOrEmpty(model.agency_com) ? DBNull.Value : model.agency_com),
                    new SqlParameter("@case_text",  string.IsNullOrEmpty(model.case_text) ? DBNull.Value : model.case_text),
                    new SqlParameter("@CS_text",  string.IsNullOrEmpty(model.CS_text) ? DBNull.Value : model.CS_text),
                    new SqlParameter("@check_date",  string.IsNullOrEmpty(model.check_date) ? DBNull.Value : model.check_date),
                    new SqlParameter("@check_address",  string.IsNullOrEmpty(model.check_address) ? DBNull.Value : model.check_address),
                    new SqlParameter("@pass_amount",  string.IsNullOrEmpty(model.pass_amount) ? DBNull.Value : model.pass_amount),
                    new SqlParameter("@set_amount",  string.IsNullOrEmpty(model.set_amount) ? DBNull.Value : model.set_amount),
                    new SqlParameter("@print_data",  string.IsNullOrEmpty(model.print_data) ? DBNull.Value : model.print_data),
                    new SqlParameter("@get_data",  string.IsNullOrEmpty(model.get_data) ? DBNull.Value : model.get_data),
                    new SqlParameter("@process_charge",  string.IsNullOrEmpty(model.process_charge) ? DBNull.Value : model.process_charge),
                    new SqlParameter("@AG_note",  string.IsNullOrEmpty(model.AG_note) ? DBNull.Value : model.AG_note),
                    new SqlParameter("@check_leader_num",  string.IsNullOrEmpty(model.check_leader_num) ? DBNull.Value : model.check_leader_num),
                    new SqlParameter("@check_process_num",  string.IsNullOrEmpty(model.check_process_num) ? DBNull.Value : model.check_process_num),
                    //new SqlParameter("@set_process_date",  string.IsNullOrEmpty(model.set_process_date) ? DBNull.Value : model.set_process_date),
                    new SqlParameter("@check_process_type",  string.IsNullOrEmpty(model.check_process_type) ? "N" : model.check_process_type),
                    //new SqlParameter("@check_process_date",  string.IsNullOrEmpty(model.check_process_date) ? DBNull.Value : model.check_process_date),
                    new SqlParameter("@check_process_note",  string.IsNullOrEmpty(model.check_process_note) ? DBNull.Value : model.check_process_note),
                    new SqlParameter("@close_type",  string.IsNullOrEmpty(model.close_type) ? "N" : model.close_type),
                    new SqlParameter("@close_type_date",  string.IsNullOrEmpty(model.close_type_date) ? DBNull.Value : model.close_type_date),
                    new SqlParameter("@del_tag",  string.IsNullOrEmpty(model.del_tag) ? DBNull.Value : model.del_tag),

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
        /// 修改委對單資料 House_agency_Upd
        /// </summary>
        [HttpPost("House_agency_Upd")]
        public ActionResult<ResultClass<string>> House_agency_Upd(House_agency_Ins model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"UPDATE [dbo].[House_agency]
                               SET 
                                    [case_com] = @case_com,
                                    [agency_com] = @agency_com,
                                    [case_text] = @case_text,
                                    [CS_text] = @CS_text,
                                    [check_date] = @check_date,
                                    [check_address] = @check_address,
                                    [pass_amount] = @pass_amount,
                                    [set_amount] = @set_amount,
                                    [print_data] = @print_data,
                                    [get_data] = @get_data,
                                    [process_charge] = @process_charge,
                                    [AG_note] = @AG_note,
                                    [check_leader_num] = @check_leader_num,
                                    [check_process_num] = @check_process_num,
                                    [set_process_date] = @set_process_date,
                                    [check_process_type] = @check_process_type,
                                    [check_process_date] = @check_process_date,
                                    [check_process_note] = @check_process_note,
                                    [close_type] = @close_type,
                                    [close_type_date] = @close_type_date,
                                    [del_tag] = @del_tag,
                                    [edit_date] = @edit_date,
                                    [edit_num] = @edit_num,
                                    [edit_ip] = @edit_ip
                                where [AG_id] = @AG_id";
                var parameters = new List<SqlParameter>
                {
                    new SqlParameter("@AG_id", model.AG_id),
                    new SqlParameter("@case_com",  string.IsNullOrEmpty(model.case_com) ? DBNull.Value : model.case_com),
                    new SqlParameter("@agency_com",  string.IsNullOrEmpty(model.agency_com) ? DBNull.Value : model.agency_com),
                    new SqlParameter("@case_text",  string.IsNullOrEmpty(model.case_text) ? DBNull.Value : model.case_text),
                    new SqlParameter("@CS_text",  string.IsNullOrEmpty(model.CS_text) ? DBNull.Value : model.CS_text),
                    new SqlParameter("@check_date",  string.IsNullOrEmpty(model.check_date) ? DBNull.Value : model.check_date),
                    new SqlParameter("@check_address",  string.IsNullOrEmpty(model.check_address) ? DBNull.Value : model.check_address),
                    new SqlParameter("@pass_amount",  string.IsNullOrEmpty(model.pass_amount) ? DBNull.Value : model.pass_amount),
                    new SqlParameter("@set_amount",  string.IsNullOrEmpty(model.set_amount) ? DBNull.Value : model.set_amount),
                    new SqlParameter("@print_data",  string.IsNullOrEmpty(model.print_data) ? DBNull.Value : model.print_data),
                    new SqlParameter("@get_data",  string.IsNullOrEmpty(model.get_data) ? DBNull.Value : model.get_data),
                    new SqlParameter("@process_charge",  string.IsNullOrEmpty(model.process_charge) ? DBNull.Value : model.process_charge),
                    new SqlParameter("@AG_note",  string.IsNullOrEmpty(model.AG_note) ? DBNull.Value : model.AG_note),
                    new SqlParameter("@check_leader_num",  string.IsNullOrEmpty(model.check_leader_num) ? DBNull.Value : model.check_leader_num),
                    new SqlParameter("@check_process_num",  string.IsNullOrEmpty(model.check_process_num) ? DBNull.Value : model.check_process_num),
                    new SqlParameter("@set_process_date",  string.IsNullOrEmpty(model.set_process_date) ? DBNull.Value : model.set_process_date),
                    new SqlParameter("@check_process_type",  string.IsNullOrEmpty(model.check_process_type) ? DBNull.Value : model.check_process_type),
                    new SqlParameter("@check_process_date",  string.IsNullOrEmpty(model.check_process_date) ? DBNull.Value : model.check_process_date),
                    new SqlParameter("@check_process_note",  string.IsNullOrEmpty(model.check_process_note) ? DBNull.Value : model.check_process_note),
                    new SqlParameter("@close_type",  string.IsNullOrEmpty(model.close_type) ? DBNull.Value : model.close_type),
                    new SqlParameter("@close_type_date",  string.IsNullOrEmpty(model.close_type_date) ? DBNull.Value : model.close_type_date),
                    new SqlParameter("@del_tag",  string.IsNullOrEmpty(model.del_tag) ? DBNull.Value : model.del_tag),
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
        /// 刪除委對單單筆 House_agency_Del
        /// </summary>
        [HttpDelete("House_agency_Del")]
        public ActionResult<ResultClass<string>> House_agency_Del(string Id)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"Delete House_agency where AG_id=@Id";
                var parameters = new List<SqlParameter>
                {
                    new SqlParameter("@Id",Id)
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


        #endregion

        #region 撥款及費用確認書
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
                    parameters.Add(new SqlParameter("@CS_name", "%" + model.CS_name + "%"));
                }

                if (!string.IsNullOrEmpty(model.Date_S) && !string.IsNullOrEmpty(model.Date_E))
                {
                    sqlBuilder.Append(" AND get_amount_date between @Date_S and @Date_E ");
                    parameters.Add(new SqlParameter("@Date_S", FuncHandler.ConvertROCToGregorian(model.Date_S)));
                    parameters.Add(new SqlParameter("@Date_E", FuncHandler.ConvertROCToGregorian(model.Date_E)));
                }

                if (!string.IsNullOrEmpty(model.OrderByStr))
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

        [HttpGet("House_SendCase_SQuery")]
        public ActionResult<ResultClass<string>> House_SendCase_SQuery(string Id)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            return Ok(resultClass);
        }

        #endregion

    }
}
