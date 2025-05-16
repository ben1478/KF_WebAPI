using KF_WebAPI.BaseClass;
using KF_WebAPI.BaseClass.AE;
using KF_WebAPI.FunctionHandler;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Newtonsoft.Json;
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
                    "   ,(select U_name FROM User_M where U_num = ho.plan_num AND del_tag='0') as plan_name " +
                    "   ,(SELECT COUNT(*) FROM ASP_UpLoad WHERE cknum = ho.case_id and del_tag = '0') AS case_id_count "+
                    "   ,ub.item_D_name BC_name, ub.item_D_code BC_code, confirm_num " +
                    " FROM House_othercase ho " +
                    "   LEFT JOIN User_M M1 on ho.plan_num=M1.u_num " +
                    "   LEFT JOIN Item_list ub ON ub.item_M_code='branch_company'  AND ub.item_D_code=M1.U_BC " +
                    "   WHERE 1 = 1 ");

                var parameters = new List<SqlParameter>();

                //區
                if (!string.IsNullOrEmpty(model.BC_code))
                {
                    sqlBuilder.Append(" AND item_D_code = @BC_code ");
                    parameters.Add(new SqlParameter("@BC_code", model.BC_code));
                }

                if (!string.IsNullOrEmpty(model.selYear_S))
                {
                    string y = model.selYear_S.Split('-')[0];
                    string m = model.selYear_S.Split('-')[1];
                    sqlBuilder.Append(" AND year(get_amount_date) = @y and month(get_amount_date) = @m ");
                    parameters.Add(new SqlParameter("@y", y));
                    parameters.Add(new SqlParameter("@m", m));
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
        /// 車貸.撥款日期.Group年月
        /// </summary>
        [HttpGet("House_othercase_GYMQuery")]
        public ActionResult<ResultClass<string>> House_othercase_GYMQuery()
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            try
            {
                ADOData _adoData = new ADOData(); // 測試:"Test" / 正式:""
                #region SQL
                var T_SQL = @"SELECT  
                    CONVERT(VARCHAR(7), get_amount_date, 126) as v,
                    RIGHT('0'+CONVERT(VARCHAR(3), DATEPART(year,get_amount_date)-1911), 3) + '-'+
                    RIGHT('0'+CONVERT(VARCHAR(2), month(get_amount_date), 112),2) as t
                    FROM House_othercase
                    where  get_amount_date is not null
                    group by CONVERT(VARCHAR(7), get_amount_date, 126),RIGHT('0'+CONVERT(VARCHAR(3), DATEPART(year,get_amount_date)-1911), 3) + '-'+
                    RIGHT('0'+CONVERT(VARCHAR(2), month(get_amount_date), 112),2)
                    order by v desc";
                
                #endregion
                var dtResult = _adoData.ExecuteSQuery(T_SQL);

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
            var get_amount_date = DateTime.Parse(FuncHandler.ConvertROCToGregorian(model.get_amount_date));

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

                      new SqlParameter("@get_amount_date", string.IsNullOrEmpty(model.get_amount_date) ? DBNull.Value : get_amount_date),
                      
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
            var get_amount_date = DateTime.Parse(FuncHandler.ConvertROCToGregorian(model.get_amount_date));

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
                      new SqlParameter("@get_amount_date", string.IsNullOrEmpty(model.get_amount_date) ? DBNull.Value : get_amount_date),

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

            string SC = SpecialCkeck(model.U_num);
            string isDel = (SC.Contains("7005") || SC.Contains("7036") )? "Y" : "N";
            ResultClass<string> resultClass = new ResultClass<string>();
            try
            {
                ADOData _adoData = new ADOData(); // 測試:"Test" / 正式:""

                var sqlBuilder = new StringBuilder("select AG_id " +
                    ",FORMAT(agcy.add_date, 'yyyy/MM/dd', 'en-US') + ' ' +  CASE WHEN DATEPART(HOUR, agcy.add_date) < 12 THEN '上午' ELSE '下午' END + ' '  + FORMAT(agcy.add_date, 'hh:mm:ss', 'en-US') AS add_date " +
                    ",case_com, agency_com, check_process_type, close_type, check_leader_num " +
                    ",(select U_name FROM User_M where U_num = agcy.check_leader_num AND del_tag = '0') as check_leader_name " +
                    ",addUser.U_name as add_num_name, addUser.U_BC as add_U_BC " +
                    ",(select U_name FROM User_M where U_num = agcy.add_num AND del_tag = '0') as add_name " +
                    ",(select U_name FROM User_M where U_num = agcy.check_process_num AND del_tag='0') as check_process_name " +
                    ",isEdit = case when agcy.add_num = @U_num then 'Y' else 'N' end " +
                    ",@isDel as isDel "+
                    " from House_agency agcy " +
                    " join User_M as addUser on addUser.U_num = agcy.add_num AND addUser.del_tag='0' " +
                    " WHERE 1 = 1 "
                    );

                var parameters = new List<SqlParameter>
                {
                    new SqlParameter("@U_num", model.U_num),
                    new SqlParameter("@isDel", isDel)
                };

                // 委對單號
                if (model.AG_id != 0)
                {
                    sqlBuilder.Append(" AND AG_id = @AG_id ");
                    parameters.Add(new SqlParameter("@AG_id", model.AG_id));
                }

                #region 依管理作業\特定權限設定判斷
                DateTime DateS = DateTime.MinValue;
                DateTime.TryParse(FuncHandler.ConvertROCToGregorian(model.Date_S), out DateS);
                DateTime DateE = DateTime.MinValue;
                DateTime.TryParse(FuncHandler.ConvertROCToGregorian(model.Date_E), out DateE);
                

                bool isCanQuery = true;
                // 7036:業務流程區	[特定權限全區, 全時段開放] 可查詢全部人員資料且不限3個月內
                if (!(SC.Contains("7005") || SC.Contains("7036")))
                {
                    // 7018:業務流程區  [委對單]可查詢全部人員資料,資料限3個月內
                    DateS = DateE.AddMonths(-3);
                    DateE = DateTime.Today;
                    if (!SC.Contains("7018"))
                    {
                        // 角色管理:業務主管
                        if (isDepManager(model.U_num))
                        {
                            sqlBuilder.Append(" AND ((check_process_num = @U_num or agcy.add_num = @U_num or check_leader_num = @U_num) or (addUser.U_BC = @U_BC))");
                            parameters.Add(new SqlParameter("@U_BC", model.U_BC));
                        }
                        else
                        {
                            // 0:本人
                            if (SC.Contains(model.U_num))
                            {
                                // 1.add_num/check_leader_num=本人
                                sqlBuilder.Append(" AND (check_process_num = @U_num or agcy.add_num = @U_num or check_leader_num = @U_num) ");

                            }
                        }
                    }
                }
                #endregion

                // 申請日期
                //if (!string.IsNullOrEmpty(model.Date_S) && !string.IsNullOrEmpty(model.Date_E))
                //{
                    sqlBuilder.Append(" AND agcy.add_date between @Date_S and @Date_E ");
                    parameters.Add(new SqlParameter("@Date_S", DateS));
                    parameters.Add(new SqlParameter("@Date_E", DateE.AddDays(1)));
                //}
                sqlBuilder.Append(" order by AG_id desc ");
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
                    new SqlParameter("@check_leader_num",  string.IsNullOrEmpty(model.check_leader_num) ? DBNull.Value : model.check_leader_num.ToUpper()),
                    new SqlParameter("@check_process_num",  string.IsNullOrEmpty(model.check_process_num) ? DBNull.Value : model.check_process_num.ToUpper()),
                    //new SqlParameter("@set_process_date",  string.IsNullOrEmpty(model.set_process_date) ? DBNull.Value : model.set_process_date),
                    new SqlParameter("@check_process_type",  string.IsNullOrEmpty(model.check_process_type) ? "N" : model.check_process_type),
                    //new SqlParameter("@check_process_date",  string.IsNullOrEmpty(model.check_process_date) ? DBNull.Value : model.check_process_date),
                    new SqlParameter("@check_process_note",  string.IsNullOrEmpty(model.check_process_note) ? DBNull.Value : model.check_process_note),
                    new SqlParameter("@close_type",  string.IsNullOrEmpty(model.close_type) ? "N" : model.close_type),
                    new SqlParameter("@close_type_date",  string.IsNullOrEmpty(model.close_type_date) ? DBNull.Value : model.close_type_date),
                    new SqlParameter("@del_tag",  "0"),

                    new SqlParameter("@add_date", DateTime.Now),
                    new SqlParameter("@add_num", model.add_num.ToUpper()),
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
                    new SqlParameter("@check_leader_num",  string.IsNullOrEmpty(model.check_leader_num) ? DBNull.Value : model.check_leader_num.ToUpper()),
                    new SqlParameter("@check_process_num",  string.IsNullOrEmpty(model.check_process_num) ? DBNull.Value : model.check_process_num.ToUpper()),
                    new SqlParameter("@set_process_date",  string.IsNullOrEmpty(model.set_process_date) || string.IsNullOrEmpty(model.check_process_num) ? DBNull.Value : model.set_process_date),
                    new SqlParameter("@check_process_type",  string.IsNullOrEmpty(model.check_process_type) ? DBNull.Value : model.check_process_type),
                    new SqlParameter("@check_process_date",  string.IsNullOrEmpty(model.check_process_date) ? DBNull.Value : model.check_process_date),
                    new SqlParameter("@check_process_note",  string.IsNullOrEmpty(model.check_process_note) ? DBNull.Value : model.check_process_note),
                    new SqlParameter("@close_type",  string.IsNullOrEmpty(model.close_type) ? DBNull.Value : model.close_type),
                    new SqlParameter("@close_type_date",  string.IsNullOrEmpty(model.close_type_date) ? DBNull.Value : model.close_type_date),
                    
                    new SqlParameter("@edit_date", DateTime.Today),
                    new SqlParameter("@edit_num", model.edit_num.ToUpper()),
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

        /// <summary>
        /// 委對單修改權限判定 ChkEditor
        /// </summary>
        [HttpGet("ChkEditor")]
        public ActionResult<ResultClass<string>> ChkEditor(string User)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"SELECT [Map_id]
                              FROM [AE_DB_TEST].[dbo].[Menu_set]
                              where menu_id = 1034 -- 委對單
                                and u_num = @u_num
                                and per_edit = 1 -- 可修改";
                var parameters = new List<SqlParameter>
                {
                    new SqlParameter("@u_num", User)
                };
                #endregion
                var dtresult = _adoData.ExecuteQuery(T_SQL, parameters);
                resultClass.ResultCode = "000";
                if (dtresult.Rows.Count == 0)
                {
                    resultClass.objResult = "N";
                }
                else
                {
                    resultClass.objResult = "Y";
                }
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
        /// 管理作業\特定權限設定
        /// </summary>
        private string SpecialCkeck(string User)
        {
            // 7018:業務流程區  [委對單]可查詢全部人員資料
            // 7036:業務流程區	[特定權限全區, 全時段開放] 可查詢全部人員資料且不限3個月內
            // 7005	開發人員區	管理者	最高權限
            // 0:本人
            ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"select sp_id from [dbo].[Special_set] as ss
                                where sp_type = 1 and SS.U_num = @U_num and SS.sp_id in  (7005,7018,7036) ";
                var parameters = new List<SqlParameter>
                {
                    new SqlParameter("@u_num", User)
                };
                #endregion
                var dtresult = _adoData.ExecuteQuery(T_SQL, parameters);
                
                string sResult = "";
                foreach(DataRow dr in dtresult.Rows)
                {
                    sResult = dr["sp_id"].ToString() + "/"+sResult ; 
                }
                if(string.IsNullOrEmpty(sResult))
                {
                    sResult = User; //回傳本人員編
                }
                else
                {
                    sResult = sResult.Substring(0, sResult.Length - 1); //回傳:7018/7036
                }
                
                return sResult;
            
        }

        private bool isDepManager(string User)
        {
            //業務主管
            ADOData _adoData = new ADOData();
            #region SQL
            var T_SQL = @"select u_name, u_num, u_bc, Role_num " +
                "from user_M as u " +
                "where u_num = @u_num ";

            var parameters = new List<SqlParameter>
                {
                    new SqlParameter("@u_num", User)
                };
            #endregion
            var dtresult = _adoData.ExecuteQuery(T_SQL, parameters);
            bool isDep = false;
            foreach (DataRow dr in dtresult.Rows)
            {
                isDep = dr["Role_num"].ToString() == "1008"; // 業務主管:1008
            }
            return isDep;

        }
        #endregion

        #region 撥款及費用確認書
        /// <summary>
        /// 撥款及費用確認書列表查詢
        /// </summary>
        [HttpPost("House_SendCase_LQuery")]
        public ActionResult<ResultClass<string>> House_SendCase_LQuery(House_sendcase_Req model)
        {
            
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
                //申請人
                if (!string.IsNullOrEmpty(model.CS_name))
                {
                    sqlBuilder.Append(" AND A.CS_name like  @CS_name ");
                    parameters.Add(new SqlParameter("@CS_name", "%" + model.CS_name + "%"));
                }

                //區
                if (!string.IsNullOrEmpty(model.BC_code))
                {
                    sqlBuilder.Append(" AND M.U_BC = @BC_code ");
                    parameters.Add(new SqlParameter("@BC_code", model.BC_code));
                }

                //撥款年月
                if (!string.IsNullOrEmpty(model.selYear_S))
                {
                    string y = model.selYear_S.Split('-')[0];
                    string m = model.selYear_S.Split('-')[1];
                    sqlBuilder.Append(" AND year(get_amount_date) = @y and month(get_amount_date) = @m ");
                    parameters.Add(new SqlParameter("@y", y));
                    parameters.Add(new SqlParameter("@m", m));
                }

                //業務
                if (!string.IsNullOrEmpty(model.plan_name))
                {
                    sqlBuilder.Append(" AND M.U_name = @plan_name ");
                    parameters.Add(new SqlParameter("@plan_name", model.plan_name));
                    
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


        /// <summary>
        /// 撥款及費用確認書.撥款日期.Group年月
        /// </summary>
        [HttpGet("House_SendCase_GYMQuery")]
        public ActionResult<ResultClass<string>> House_SendCase_GYMQuery()
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            try
            {
                ADOData _adoData = new ADOData(); // 測試:"Test" / 正式:""
                #region SQL
                var T_SQL = @"SELECT  
                    CONVERT(VARCHAR(7), get_amount_date, 126) as v,
                    RIGHT('0'+CONVERT(VARCHAR(3), DATEPART(year,get_amount_date)-1911), 3) + '-'+
                    RIGHT('0'+CONVERT(VARCHAR(2), month(get_amount_date), 112),2) as t
                    FROM House_SendCase
                    where  get_amount_date is not null
                    group by CONVERT(VARCHAR(7), get_amount_date, 126),RIGHT('0'+CONVERT(VARCHAR(3), DATEPART(year,get_amount_date)-1911), 3) + '-'+
                    RIGHT('0'+CONVERT(VARCHAR(2), month(get_amount_date), 112),2)
                    order by v desc";

                #endregion
                var dtResult = _adoData.ExecuteSQuery(T_SQL);

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

        #region 客訴
        /// <summary>
        /// 客訴資料列表查詢 House_othercase_LQuery
        /// </summary>
        [HttpPost("Complaint_LQuery")]
        public ActionResult<ResultClass<string>> Complaint_LQuery(Complaint_Req model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            try
            {
                ADOData _adoData = new ADOData(); // 測試:"Test" / 正式:""

                var sqlBuilder = new StringBuilder(
                    "SELECT Comp_Id, CS_name, M1.U_name Sales_name, ub.item_D_name BC_name, Complaint, CompDate, Remark, M2.U_name add_name " +
                        "FROM dbo.Complaint C " +
                        "LEFT JOIN User_M M1 on C.Sales_num = M1.u_num " +
                        "LEFT JOIN Item_list ub ON ub.item_M_code = 'branch_company'  AND ub.item_D_code = M1.U_BC " +
                        "LEFT JOIN User_M M2 on C.Add_num = M2.u_num " +
                        "where 1 = 1 "
                        );

                var parameters = new List<SqlParameter>();

                //區
                if (!string.IsNullOrEmpty(model.BC_code))
                {
                    sqlBuilder.Append(" AND M1.U_BC = @BC_code ");
                    parameters.Add(new SqlParameter("@BC_code", model.BC_code));
                }

                //客訴時間(年月)
                if (!string.IsNullOrEmpty(model.selYear_S))
                {
                    sqlBuilder.Append($" AND left(C.compDate,{model.selYear_S.Length}) = @selYear_S ");
                    parameters.Add(new SqlParameter("@selYear_S", model.selYear_S));
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
        /// 客訴資料單筆查詢 Manufacturer_SQuery
        /// </summary>
        [HttpGet("Complaint_SQuery")]
        public ActionResult<ResultClass<string>> Complaint_SQuery(string Id)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            try
            {
                ADOData _adoData = new ADOData(); // 測試:"Test" / 正式:""
                #region SQL

                var T_SQL =
                    @"SELECT Comp_Id, CS_name, M1.U_name Sales_name, ub.item_D_name BC_name, Complaint, CompDate, Remark, M2.U_name add_name,M3.U_name edit_name " +
                        "FROM dbo.Complaint C " +
                        "LEFT JOIN User_M M1 on C.Sales_num = M1.u_num " +
                        "LEFT JOIN Item_list ub ON ub.item_M_code = 'branch_company'  AND ub.item_D_code = M1.U_BC " +
                        "LEFT JOIN User_M M2 on C.Add_num = M2.u_num " +
                        "LEFT JOIN User_M M3 on C.Edit_num = M3.u_num " +
                        "where Comp_Id = @Id ";


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
        /// 新增客訴資料資料 House_othercase_Ins
        /// </summary>
        [HttpPost("Complaint_Ins")]
        public ActionResult<ResultClass<string>> Complaint_Ins(Complaint_Ins model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();
            string case_id = FuncHandler.GetCheckNum();
            case_id = case_id.Substring(0, case_id.Length - 6);

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL

                var T_SQL = @"INSERT INTO [Complaint]
                       ([CS_name] ,
                        [Sales_num] ,
                        [Complaint] ,
                        [CompDate] ,
                        [CompTime] ,
                        [Remark] ,
                        [add_date] ,
                        [add_num]
                        )
                 VALUES
                       (
                        @CS_name,
                        @Sales_num,
                        @Complaint,
                        @CompDate,
                        @CompTime,
                        @Remark,
                        @add_date,
                        @add_num                     
                        )";

                var parameters = new List<SqlParameter>
                  {
                        new SqlParameter("@CS_name",  string.IsNullOrEmpty(model.CS_name) ? DBNull.Value : model.CS_name),
                        new SqlParameter("@Sales_num",  string.IsNullOrEmpty(model.Sales_num) ? DBNull.Value : model.Sales_num),
                        new SqlParameter("@Complaint",  string.IsNullOrEmpty(model.Complaint) ? DBNull.Value : model.Complaint),
                        new SqlParameter("@CompDate",  string.IsNullOrEmpty(model.CompDate) ? DBNull.Value : model.CompDate.Replace("-0","/")),
                        new SqlParameter("@CompTime",  string.IsNullOrEmpty(model.CompTime) ? DBNull.Value : model.CompTime),
                        new SqlParameter("@Remark",  string.IsNullOrEmpty(model.Remark) ? DBNull.Value : model.Remark),
                        new SqlParameter("@add_date", DateTime.Today),
                        new SqlParameter("@add_num", model.add_num)
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
        /// 修改客訴資料資料 House_othercase_Upd
        /// </summary>
        [HttpPost("Complaint_Upd")]
        public ActionResult<ResultClass<string>> Complaint_Upd(Complaint_Ins model)
        {

            // CompDate 傳入格式 國曆：yyy-MM-dd 需轉成國曆：yyy/MM/dd 並且要去零 。例如:113-01-01 -> 113/1/1

            ResultClass<string> resultClass = new ResultClass<string>();
            
            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"UPDATE [dbo].[Complaint]
                               SET 
                                    [CS_name] = @CS_name,
                                    [Sales_num] = @Sales_num,
                                    [Complaint] = @Complaint,
                                    [CompDate] = @CompDate,
                                    [CompTime] = @CompTime,
                                    [Remark] = @Remark,
                                    [edit_date] = @edit_date,
                                    [edit_num] = @edit_num
                                    where[Comp_Id] = @Comp_Id";
                var parameters = new List<SqlParameter>
                {
                        new SqlParameter("@Comp_Id",  model.Comp_Id),
                        new SqlParameter("@CS_name",  string.IsNullOrEmpty(model.CS_name) ? DBNull.Value : model.CS_name),
                        new SqlParameter("@Sales_num",  string.IsNullOrEmpty(model.Sales_num) ? DBNull.Value : model.Sales_num),
                        new SqlParameter("@Complaint",  string.IsNullOrEmpty(model.Complaint) ? DBNull.Value : model.Complaint),
                        new SqlParameter("@CompDate",  string.IsNullOrEmpty(model.CompDate) ? DBNull.Value : model.CompDate.Replace("-0","/")),
                        new SqlParameter("@CompTime",  string.IsNullOrEmpty(model.CompTime) ? DBNull.Value : model.CompTime),
                        new SqlParameter("@Remark",  string.IsNullOrEmpty(model.Remark) ? DBNull.Value : model.Remark),
                        new SqlParameter("@edit_date", DateTime.Today),
                        new SqlParameter("@edit_num",  model.edit_num)
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
        [HttpDelete("Complaint_Del")]
        public ActionResult<ResultClass<string>> Complaint_Del(string Id)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"Delete Complaint where Comp_Id=@Id";
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

        /// <summary>
        /// 客訴.日期.Group年月
        /// </summary>
        [HttpGet("Complaint_GYMQuery")]
        public ActionResult<ResultClass<string>> Complaint_GYMQuery()
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            try
            {
                ADOData _adoData = new ADOData(); // 測試:"Test" / 正式:""
                #region SQL
                var T_SQL = @"  select
                    y+'-'+replace(m,'0','') as v
                    from (  
                        SELECT  
                            CompDate,
                            case 
                                when charindex('/',CompDate) > 3 then
                                    left(CompDate, charindex('/',CompDate)-1)
                                else
                                    null
                                END as y,
                            case 
                                when  charindex('/',right(CompDate,len(CompDate)-charindex('/',CompDate))) > 0 then
                                    SUBSTRING(CompDate, charindex('/',CompDate)+1, 
                                        charindex('/',right(CompDate,len(CompDate)-charindex('/',CompDate)))-1)
                                else
                                    null
                                END as m
                            FROM Complaint
                            where  CompDate is not null
                    ) as ym
                    where y is not NULL and  m is not null
                    group by  y+'-'+replace(m,'0','')
                    order by v desc";

                #endregion
                var dtResult = _adoData.ExecuteSQuery(T_SQL);

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
    }
}
