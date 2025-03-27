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
using System.Xml.Linq;

namespace KF_WebAPI.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class AE_AgcyController : ControllerBase
    {
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
                    ",(select U_name FROM User_M where U_num = agcy.check_process_num AND del_tag='0') as check_process_name "+
                    "from House_agency agcy WHERE 1 = 1 ");

                var parameters = new List<SqlParameter>();
                if (model.AG_id != 0 )
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
        /// 廠商資料單筆查詢 Manufacturer_SQuery
        /// </summary>
        [HttpGet("House_agency_SQuery")]
        public ActionResult<ResultClass<string>> House_agency_SQuery(string Id)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            try
            {
                ADOData _adoData = new ADOData(); // 測試:"Test" / 正式:""
                #region SQL
                var T_SQL = @"select AG_id, FORMAT(add_date,'yyyy/MM/dd') add_date, case_com, agency_com, case_text,CS_text,check_date,check_address, pass_amount,set_amount, print_data,get_data,process_charge,AG_note, check_leader_num, check_process_type, close_type " +
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
                               ([AG_cknum]
                               ,[HS_id]
                               ,[case_com]
                               ,[agency_com]
                               ,[case_text]
                               ,[CS_text]
                               ,[check_date]
                               ,[check_address]
                               ,[pass_amount]
                               ,[set_amount]
                               ,[print_data]
                               ,[get_data]
                               ,[process_charge]
                               ,[AG_note]
                               ,[check_leader_num]
                               ,[check_process_num]
                               ,[set_process_date]
                               ,[check_process_type]
                               ,[check_process_date]
                               ,[check_process_note]
                               ,[close_type]
                               ,[close_type_date]
                               ,[del_tag]
                               ,[add_date]
                               ,[add_num]
                               ,[add_ip]
                               )
                         VALUES
                               (@AG_cknum, 
                               @HS_id, 
                               @case_com, 
                               @agency_com, 
                               @case_text, 
                               @CS_text, 
                               @check_date, 
                               @check_address, 
                               @pass_amount, 
                               @set_amount, 
                               @print_data, 
                               @get_data, 
                               @process_charge, 
                               @AG_note, ntext,>
                               @check_leader_num, 
                               @check_process_num, 
                               @set_process_date, 
                               @check_process_type, 
                               @check_process_date, 
                               @check_process_note, 
                               @close_type, 
                               @close_type_date, 
                               @del_tag, 
                               @add_date, 
                               @add_num, 
                               @add_ip 
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
                    new SqlParameter("@set_process_date",  string.IsNullOrEmpty(model.set_process_date) ? DBNull.Value : model.set_process_date),
                    new SqlParameter("@check_process_type",  string.IsNullOrEmpty(model.check_process_type) ? DBNull.Value : model.check_process_type),
                    new SqlParameter("@check_process_date",  string.IsNullOrEmpty(model.check_process_date) ? DBNull.Value : model.check_process_date),
                    new SqlParameter("@check_process_note",  string.IsNullOrEmpty(model.check_process_note) ? DBNull.Value : model.check_process_note),
                    new SqlParameter("@close_type",  string.IsNullOrEmpty(model.close_type) ? DBNull.Value : model.close_type),
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
        /// 刪除單筆 House_agency_Del
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

    }
}
