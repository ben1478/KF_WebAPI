using KF_WebAPI.BaseClass;
using KF_WebAPI.BaseClass.AE;
using KF_WebAPI.FunctionHandler;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Primitives;
using Newtonsoft.Json;
using System;
using System.Data;
using Microsoft.Data.SqlClient;
using System.Drawing.Drawing2D;
using System.Reflection;
using Microsoft.AspNetCore.Mvc.Infrastructure;
using Microsoft.IdentityModel.Tokens;
using Microsoft.AspNetCore.Mvc.Routing;
using System.Linq;

namespace KF_WebAPI.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class AE_RPTController : Controller
    {
        #region 業績報表_業務
        /// <summary>
        /// 取得show_more_type判定是否為7011權限
        /// </summary>
        [HttpGet("GetShowMoreTypeUser")]
        public ActionResult<ResultClass<string>> GetShowMoreTypeUser()
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var User_Num = HttpContext.Session.GetString("UserID");

            try
            {
                string[] str = new string[] { "7011" };
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
        /// 提供業務名單資料(當show_more_type='Y') GetTeamUsersList/select_team_more.asp
        /// </summary>
        /// <returns></returns>
        [HttpGet("GetTeamUsersList")]
        public ActionResult<ResultClass<string>> GetTeamUsersList()
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var User_Num = HttpContext.Session.GetString("UserID");
            var U_Check_BC_txt = "xxx";

            try
            {
                ADOData _adoData = new ADOData();

                #region SQL
                var parameters_sc = new List<SqlParameter>();
                var T_SQL_SC = "Select U_BC,U_Check_BC From User_M Where del_tag = '0' AND U_num=@U_num";
                parameters_sc.Add(new SqlParameter("@U_num", User_Num));
                #endregion
                DataTable dtResult_sc = _adoData.ExecuteQuery(T_SQL_SC, parameters_sc);
                DataRow row = dtResult_sc.Rows[0];
                if (!string.IsNullOrEmpty(row["U_Check_BC"].ToString()))
                {
                    U_Check_BC_txt = U_Check_BC_txt + row["U_Check_BC"].ToString().Replace('#', ',');
                }
                else
                {
                    U_Check_BC_txt = U_Check_BC_txt + "," + row["U_BC"].ToString();
                }

                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"
                    SELECT bc.item_D_name AS U_BC_name,um.U_num,um.U_name,pft.item_D_name AS U_PFT_name FROM User_M um
                    LEFT JOIN Item_list bc ON bc.item_M_code = 'branch_company'  AND bc.item_D_code = um.U_BC AND bc.item_D_type = 'Y' AND bc.show_tag = '0' AND bc.del_tag = '0'
                    LEFT JOIN Item_list pft ON pft.item_M_code = 'professional_title' AND pft.item_D_code = um.U_PFT AND pft.item_D_type = 'Y' AND pft.show_tag = '0' AND pft.del_tag = '0'
                    WHERE um.del_tag = '0' AND bc.item_D_name is not null
                    AND U_num IN (Select distinct group_D_code from view_User_group)
                    AND isnull(U_type,'')='' AND (U_leave_date is null OR U_leave_date >= DATEADD(MONTH, -2, GETDATE()))
                    AND um.U_BC IN (SELECT SplitValue FROM dbo.SplitStringFunction(@U_Check_BC_txt))
                    ORDER BY bc.item_sort,pft.item_sort";
                parameters.Add(new SqlParameter("@U_Check_BC_txt", U_Check_BC_txt));
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
        /// 提供可查詢業務人員之權限(當show_more_type='N'時 迴圈呼叫Feat_user_list) GetUGlist/Feat_user_list.asp&_fn.asp
        /// </summary>
        [HttpGet("GetUGlist")]
        public ActionResult<ResultClass<string>> GetUGlist()
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var User_Num = HttpContext.Session.GetString("UserID");

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"
                    Select distinct group_D_code from view_User_group Where group_M_code = @group_M_code
                    And GETDATE() between group_M_start_day and group_M_end_day";
                parameters.Add(new SqlParameter("@group_M_code", User_Num));
                #endregion
                DataTable dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
                List<UserGroup> groups = new List<UserGroup>();
                if (dtResult.Rows.Count > 0)
                {
                    foreach (DataRow row in dtResult.Rows)
                    {
                        UserGroup group = new UserGroup { U_num = row["group_D_code"].ToString() };
                        groups.Add(group);
                    }
                }
                groups.Add(new UserGroup { U_num = User_Num });
                resultClass.ResultCode = "000";
                resultClass.objResult = JsonConvert.SerializeObject(groups.OrderBy(p => p.U_num));
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
        /// 業績報表_業務列表 Feat_user_list/Feat_user_list.asp
        /// </summary>
        [HttpPost("Feat_user_list")]
        public ActionResult<ResultClass<string>> Feat_user_list_Query(Feat_user_list_req model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var User_Num = HttpContext.Session.GetString("UserID");

            try
            {
                ADOData _adoData = new ADOData();

                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"
                    select House_apply.CS_name,House_sendcase.HS_id
                    ,(select U_name FROM User_M where U_num = House_apply.plan_num AND del_tag='0') as plan_name
                    ,(select item_D_name from Item_list where item_M_code = 'appraise_company' AND item_D_type='Y' AND item_D_code = House_sendcase.appraise_company  AND show_tag='0' AND del_tag='0') as show_appraise_company
                    ,(select item_D_name from Item_list where item_M_code = 'project_title' AND item_D_type='Y' AND item_D_code = House_pre_project.project_title  AND show_tag='0' AND del_tag='0') as show_project_title
                    ,get_amount_date,get_amount,Loan_rate,interest_rate_original,interest_rate_pass,charge_M,charge_flow,charge_agent
                    ,charge_check,get_amount_final,House_apply.CS_introducer,HS_note,House_pre_project.project_title,exception_type,exception_rate
                    from House_sendcase
                    LEFT JOIN House_apply on House_apply.HA_id = House_sendcase.HA_id AND House_apply.del_tag='0'
                    LEFT JOIN House_pre_project on House_pre_project.HP_project_id = House_sendcase.HP_project_id AND House_pre_project.del_tag='0'
                    LEFT JOIN (select U_num ,U_BC FROM User_M where del_tag='0' ) User_M ON User_M.U_num = House_apply.plan_num
                    where House_sendcase.del_tag = '0' AND isnull(House_sendcase.get_amount,'')<>''
                    AND ( get_amount_date between @start_date AND @end_date )
                    AND ( House_apply.plan_num = @U_num )";
                #region 提供可查詢公司權限
                string[] str_all = new string[] { "7011" };
                SpecialClass specialClass = FuncHandler.CheckSpecial(str_all, User_Num);
                if (specialClass.special_check == "N")
                {
                    string[] str = new string[] { "7020", "7021", "7022", "7023", "7024", "7025" };
                    specialClass = FuncHandler.CheckSpecial(str, User_Num);
                    if (specialClass.special_check == "Y")
                    {
                        T_SQL += " AND User_M.U_BC IN (@Check_U_BC)";
                        parameters.Add(new SqlParameter("@Check_U_BC", specialClass.BC_Strings));
                    }
                }
                #endregion
                T_SQL += " order by House_sendcase.HS_id desc";
                parameters.Add(new SqlParameter("@start_date", model.start_date));
                parameters.Add(new SqlParameter("@end_date", model.end_date));
                //判斷組員還是組長
                if (!string.IsNullOrEmpty(model.U_num))
                {
                    parameters.Add(new SqlParameter("@U_num", model.U_num));
                }
                else
                {
                    parameters.Add(new SqlParameter("@U_num", User_Num));
                }
                #endregion
                DataTable dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
                if (dtResult.Rows.Count > 0)
                {
                    var Feat_user_list = dtResult.AsEnumerable().Select(row => new Feat_user_list_res
                    {
                        CS_name = row.Field<string>("CS_name"),
                        HS_id = row.Field<decimal>("HS_id"),
                        plan_name = row.Field<string>("plan_name"),
                        show_appraise_company = row.Field<string>("show_appraise_company"),
                        show_project_title = row.Field<string>("show_project_title"),
                        get_amount_date = row.Field<DateTime>("get_amount_date"),
                        get_amount = row.Field<string>("get_amount"),
                        Loan_rate = row.Field<string>("Loan_rate"),
                        interest_rate_original = row.Field<string>("interest_rate_original"),
                        interest_rate_pass = row.Field<string>("interest_rate_pass"),
                        charge_M = row.Field<string>("charge_M"),
                        charge_flow = row.Field<string>("charge_flow"),
                        charge_agent = row.Field<string>("charge_agent"),
                        charge_check = row.Field<string>("charge_check"),
                        get_amount_final = row.Field<string>("get_amount_final"),
                        CS_introducer = row.Field<string>("CS_introducer"),
                        HS_note = row.Field<string>("HS_note"),
                        project_title = row.Field<string>("project_title"),
                        exception_type = row.Field<string>("exception_type"),
                        exception_rate = row.Field<string>("exception_rate"),
                        FR_D_discount = "10",
                        Exceptions = "無"
                    }).ToList();

                    foreach (var item in Feat_user_list)
                    {
                        if (!string.IsNullOrEmpty(item.interest_rate_pass))
                        {
                            #region SQL
                            var parameters_ru = new List<SqlParameter>();
                            var T_SQL_Ru = @"
                                select FR_D_ratio_A,FR_D_ratio_B,FR_D_rate,FR_D_discount*10 AS show_FR_D_discount,FR_D_replace,FR_D_discount from Feat_rule
                                Where show_tag='0' AND del_tag='0' AND FR_D_type='Y'
                                AND FR_M_code = @FR_M_code AND FR_D_rate=@FR_D_rate
                                AND ( FR_D_ratio_A <=@Loan_rate AND FR_D_ratio_B >=@Loan_rate )";
                            parameters_ru.Add(new SqlParameter("@FR_M_code", item.project_title));
                            parameters_ru.Add(new SqlParameter("@FR_D_rate", item.interest_rate_pass));
                            parameters_ru.Add(new SqlParameter("@Loan_rate", item.Loan_rate));
                            #endregion
                            var Result_ru = _adoData.ExecuteQuery(T_SQL_Ru, parameters_ru);
                            if (Result_ru.Rows.Count > 0)
                            {
                                item.Feat_rule_Detail = Result_ru.AsEnumerable().Select(row => new Feat_rule_Detail
                                {
                                    FR_D_ratio_A = row.Field<decimal>("FR_D_ratio_A"),
                                    FR_D_ratio_B = row.Field<decimal>("FR_D_ratio_B"),
                                    FR_D_rate = row.Field<string>("FR_D_rate"),
                                    show_FR_D_discount = row.Field<int>("show_FR_D_discount"),
                                    FR_D_replace = row.Field<string>("FR_D_replace")
                                }).FirstOrDefault();

                                item.FR_D_discount = (item.Feat_rule_Detail.show_FR_D_discount / 10).ToString();

                                if (item.Feat_rule_Detail != null)
                                {
                                    if (string.IsNullOrEmpty(item.charge_M))
                                        item.charge_M = "0";
                                    if (Convert.ToDecimal(item.charge_M) * 100 > Convert.ToDecimal(item.Feat_rule_Detail.FR_D_replace) * 100)
                                        item.FR_D_discount = "10";
                                    if (item.exception_type == "Y")
                                    {
                                        item.FR_D_discount = "10";
                                        item.Exceptions = (Convert.ToDecimal(item.exception_rate) * 10).ToString();
                                    }
                                }
                            }
                        }
                        item.Performance_Discount = (Convert.ToDecimal(item.FR_D_discount) * 10).ToString();
                        item.Performance_Amount = ((Convert.ToDecimal(item.get_amount) * Convert.ToDecimal(item.FR_D_discount)) / 10).ToString();
                    }

                    resultClass.objResult = JsonConvert.SerializeObject(Feat_user_list);
                }
                else
                {
                    resultClass.ResultCode = "400";
                    resultClass.ResultMsg = "查無資料";
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

        //_待定義需求?
        /// <summary>
        /// 業績報表_業務合計計算 Feat_user_Totail_Query/Feat_user_list.asp
        /// </summary>
        /// <param name="TotalCheckAmount">總業績</param>
        /// <param name="TotalGetAmount">業績金額</param>
        /// <returns></returns>
        [HttpPost("Feat_user_Totail_Query")]
        public ActionResult<ResultClass<string>> Feat_user_Totail_Query(int TotalCheckAmount, int TotalGetAmount)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();

                Feat_rule_Totail FeatTotailModel = new Feat_rule_Totail();
                FeatTotailModel.show_cash = 0;


                var T_SQL = "select * from Feat_range where del_tag = '0' AND range_D_type = 'Y' AND range_M_code = 'FRANGEMB01'  order by range_sort,range_id";
                DataTable dtResult = _adoData.ExecuteSQuery(T_SQL);
                var FeatRangeList = dtResult.AsEnumerable().Select(row => new Feat_range
                {
                    range_id = row.Field<decimal>("range_id"),
                    range_cknum = row.Field<string>("range_cknum"),
                    range_M_type = row.Field<string>("range_M_type"),
                    range_M_code = row.Field<string>("range_M_code"),
                    range_M_name = row.Field<string>("range_M_name"),
                    range_D_type = row.Field<string>("range_D_type"),
                    range_D_code = row.Field<string>("range_D_code"),
                    range_D_name = row.Field<string>("range_D_name"),
                    range_D_ratio_A = row.Field<decimal>("range_D_ratio_A"),
                    range_D_ratio_B = row.Field<decimal>("range_D_ratio_B"),
                    range_D_rate = row.Field<string>("range_D_rate"),
                    range_D_reward = row.Field<string>("range_D_reward"),
                    range_D_base = row.Field<string>("range_D_base"),
                    range_sort = row.Field<int>("range_sort"),
                    show_tag = row.Field<string>("show_tag")
                }).ToList();

                foreach (var item in FeatRangeList)
                {
                    var check_rate = (TotalCheckAmount / Convert.ToDecimal(item.range_D_base)) * 100;
                    if (check_rate >= item.range_D_ratio_A && check_rate <= item.range_D_ratio_B)
                    {
                        FeatTotailModel.total_check_amount = TotalCheckAmount;
                        FeatTotailModel.range_D_base = item.range_D_base;
                        FeatTotailModel.check_rate = check_rate.ToString();
                        FeatTotailModel.range_D_ratio_A = item.range_D_ratio_A.ToString();
                        FeatTotailModel.range_D_ratio_B = item.range_D_ratio_B.ToString();
                        FeatTotailModel.range_D_rate = item.range_D_rate;
                        FeatTotailModel.range_D_reward = item.range_D_reward;
                        FeatTotailModel.total_get_amount = TotalGetAmount;
                        FeatTotailModel.show_cash = (TotalGetAmount * 10000) * (Convert.ToDecimal(item.range_D_rate) / 100) * (Convert.ToDecimal(item.range_D_reward) / 100);
                        break;
                    }
                }

                resultClass.objResult = JsonConvert.SerializeObject(FeatTotailModel);
                return Ok(resultClass);
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass);
            }
        }
        #endregion

        #region 業績報表_組長
        /// <summary>
        /// 取得show_more_type判定是否為7014權限
        /// </summary>
        [HttpGet("GetShowMoreTypeLeader")]
        public ActionResult<ResultClass<string>> GetShowMoreTypeLeader()
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var User_Num = HttpContext.Session.GetString("UserID");

            try
            {
                string[] str = new string[] { "7014" };
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
        /// 提供組長名單資料(當show_more_type='Y') GetTeamLeadersList/select_leader_more.asp
        /// </summary>
        [HttpGet("GetTeamLeadersList")]
        public ActionResult<ResultClass<string>> GetTeamLeadersList()
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var User_Num = HttpContext.Session.GetString("UserID");
            var U_Check_BC_txt = "xxx";

            try
            {
                ADOData _adoData = new ADOData();

                #region SQL
                var parameters_sc = new List<SqlParameter>();
                var T_SQL_SC = "Select U_BC,U_Check_BC From User_M Where del_tag = '0' AND U_num=@U_num";
                parameters_sc.Add(new SqlParameter("@U_num", User_Num));
                #endregion
                DataTable dtResult_sc = _adoData.ExecuteQuery(T_SQL_SC, parameters_sc);
                DataRow row = dtResult_sc.Rows[0];
                if (!string.IsNullOrEmpty(row["U_Check_BC"].ToString()))
                {
                    U_Check_BC_txt = U_Check_BC_txt + row["U_Check_BC"].ToString().Replace('#', ',');
                }
                else
                {
                    U_Check_BC_txt = U_Check_BC_txt + "," + row["U_BC"].ToString();
                }

                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"
                    SELECT bc.item_D_name AS U_BC_name,um.U_num,um.U_name,pft.item_D_name AS U_PFT_name FROM User_M um
                    LEFT JOIN Item_list bc ON bc.item_M_code = 'branch_company'  AND bc.item_D_code = um.U_BC AND bc.item_D_type = 'Y' AND bc.show_tag = '0' AND bc.del_tag = '0'
                    LEFT JOIN Item_list pft ON pft.item_M_code = 'professional_title' AND pft.item_D_code = um.U_PFT AND pft.item_D_type = 'Y' AND pft.show_tag = '0' AND pft.del_tag = '0'
                    WHERE um.del_tag = '0' AND bc.item_D_name is not null
                    AND U_num IN (select group_M_code from User_group where del_tag='0' AND group_M_type='Y')
                    AND isnull(U_type,'')='' AND (U_leave_date is null OR U_leave_date >= DATEADD(MONTH, -2, GETDATE()))
                    AND um.U_BC IN (SELECT SplitValue FROM dbo.SplitStringFunction(@U_Check_BC_txt))
                    ORDER BY bc.item_sort,pft.item_sort";
                parameters.Add(new SqlParameter("@U_Check_BC_txt", U_Check_BC_txt));
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
        //組資料(呈現整組資料含組長 另外再呈現組長的資料)_未完成
        [HttpPost("Feat_leader_list_Query")]
        public ActionResult<ResultClass<string>> Feat_leader_list_Query(Feat_leader_list_req model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var User_Num = HttpContext.Session.GetString("UserID");

            #region SQL
            var parameters = new List<SqlParameter>();
            var T_SQL = @"
                select House_apply.CS_name,House_sendcase.HS_id
                ,(select U_name FROM User_M where U_num = House_apply.plan_num AND del_tag='0') as plan_name
                ,(select item_D_name from Item_list where item_M_code = 'appraise_company' AND item_D_type='Y' AND item_D_code = House_sendcase.appraise_company  AND show_tag='0' AND del_tag='0') as show_appraise_company
                ,(select item_D_name from Item_list where item_M_code = 'project_title' AND item_D_type='Y' AND item_D_code = House_pre_project.project_title  AND show_tag='0' AND del_tag='0') as show_project_title
                ,get_amount_date,get_amount,Loan_rate,interest_rate_original,interest_rate_pass,charge_M,charge_flow,charge_agent
                ,charge_check,get_amount_final,House_apply.CS_introducer,HS_note,House_pre_project.project_title,exception_type,exception_rate
                from House_sendcase
                LEFT JOIN House_apply on House_apply.HA_id = House_sendcase.HA_id AND House_apply.del_tag='0'
                LEFT JOIN House_pre_project on House_pre_project.HP_project_id = House_sendcase.HP_project_id AND House_pre_project.del_tag='0'
                LEFT JOIN (select U_num ,U_BC FROM User_M where del_tag='0' ) User_M ON User_M.U_num = House_apply.plan_num
                where House_sendcase.del_tag = '0' AND isnull(House_sendcase.get_amount,'')<>''
                AND ( get_amount_date between @start_date AND @end_date )
                AND House_apply.plan_num IN (Select distinct group_D_code from view_User_group Where group_M_code = @group_M_code and del_tag='0')
                order by plan_num,HS_id desc";
            parameters.Add(new SqlParameter("@start_date", model.start_date));
            parameters.Add(new SqlParameter("@end_date", model.end_date));
            //判斷是一般組長還是業績組長
            if (!string.IsNullOrEmpty(model.leaders))
            {
                parameters.Add(new SqlParameter("@group_M_code", model.leaders));
            }
            else
            {
                parameters.Add(new SqlParameter("@group_M_code", User_Num));
            }

            #endregion

            return Ok(resultClass);
        }
        //組獎金_待定義需求?
        #endregion

        #region 業績報表_日報表_應該已經不用了 相關資訊都可以在202210版看到

        #endregion

        #region 業績報表_日報表(202210版)
        /// <summary>
        /// 業績目標設定_提供組長名單 GetMonthQuotaLeaderList
        /// </summary>
        [HttpGet("GetMonthQuotaLeaderList")]
        public ActionResult<ResultClass<string>> GetMonthQuotaLeaderList()
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var User_Num = HttpContext.Session.GetString("UserID");
            var U_Check_BC_txt = "xxx";

            try
            {
                ADOData _adoData = new ADOData();

                #region SQL
                var parameters_sc = new List<SqlParameter>();
                var T_SQL_SC = "Select U_BC,U_Check_BC From User_M Where del_tag = '0' AND U_num=@U_num";
                parameters_sc.Add(new SqlParameter("@U_num", User_Num));
                #endregion
                DataTable dtResult_sc = _adoData.ExecuteQuery(T_SQL_SC, parameters_sc);
                DataRow row = dtResult_sc.Rows[0];
                if (!string.IsNullOrEmpty(row["U_Check_BC"].ToString()))
                {
                    U_Check_BC_txt = U_Check_BC_txt + row["U_Check_BC"].ToString().Replace('#', ',');
                }
                else
                {
                    U_Check_BC_txt = U_Check_BC_txt + "," + row["U_BC"].ToString();
                }

                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"
                    SELECT bc.item_D_name AS U_BC_name,um.U_num,um.U_name,pft.item_D_name AS U_PFT_nam, um.U_BC FROM User_M um
                    LEFT JOIN Item_list bc ON bc.item_M_code = 'branch_company'  AND bc.item_D_code = um.U_BC AND bc.item_D_type = 'Y' AND bc.show_tag = '0' AND bc.del_tag = '0'
                    LEFT JOIN Item_list pft ON pft.item_M_code = 'professional_title' AND pft.item_D_code = um.U_PFT AND pft.item_D_type = 'Y' AND pft.show_tag = '0' AND pft.del_tag = '0'
                    WHERE um.del_tag = '0' AND bc.item_D_name is not null
                    AND U_num IN (select group_M_code from User_group where del_tag='0' AND group_M_type='Y')
                    AND isnull(U_type,'')='' AND (U_leave_date is null OR U_leave_date >= DATEADD(MONTH, -2, GETDATE()))
                    AND um.U_BC IN (SELECT SplitValue FROM dbo.SplitStringFunction(@U_Check_BC_txt))
                    ORDER BY bc.item_sort,pft.item_sort";
                parameters.Add(new SqlParameter("@U_Check_BC_txt", U_Check_BC_txt));
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
        /// 業績目標設定_查詢 GetMonthQuotaList/Month_quota_editor.asp
        /// </summary>
        [HttpGet("GetMonthQuotaList")]
        public ActionResult<ResultClass<string>> GetMonthQuotaList(string U_BC)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();

                DateTime startDate = DateTime.Now.AddMonths(-3);
                List<MonthQuota_res> monthquotaList = new List<MonthQuota_res>();
                for (int i = 0; i < 9; i++)
                {
                    DateTime currentMonth = startDate.AddMonths(i);
                    string formattedDate = currentMonth.ToString("yyyyMM");

                    #region SQL
                    var parameters = new List<SqlParameter>();
                    var T_SQL = @"
                        select group_M_id,group_M_name,group_D_name,group_D_code,sa.U_PFT_sort,sa.U_PFT_name
                        ,isnull((select target_quota from Feat_target ft where ft.del_tag='0' and ft.group_id=ug.group_id and ft.target_ym= @YYYYMM),0) target_quota
                        ,@YYYYMM AS target_ym
                        FROM view_User_group ug
                        join view_User_sales leader on leader.U_num = ug.group_M_code 
                        join view_User_sales sa on sa.U_num = ug.group_D_code
                        where getdate() between ug.group_M_start_day and ug.group_M_end_day and leader.U_BC =@U_BC
                        order by leader.U_BC,group_M_name,sa.U_PFT_sort,sa.U_PFT_name,group_D_code";
                    parameters.Add(new SqlParameter("@YYYYMM", formattedDate));
                    parameters.Add(new SqlParameter("@U_BC", U_BC));
                    #endregion
                    DataTable dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
                    var modelList = dtResult.AsEnumerable().Select(row => new MonthQuota_res
                    {
                        group_M_id = row.Field<decimal>("group_M_id"),
                        group_M_name = row.Field<string>("group_M_name"),
                        group_D_name = row.Field<string>("group_D_name"),
                        group_D_code = row.Field<string>("group_D_code"),
                        U_PFT_sort = row.Field<int>("U_PFT_sort"),
                        U_PFT_name = row.Field<string>("U_PFT_name"),
                        target_quota = Convert.ToInt32(row.Field<short>("target_quota")),
                        target_ym = row.Field<string>("target_ym")
                    }).ToList();
                    monthquotaList.AddRange(modelList);
                }

                resultClass.ResultCode = "000";
                resultClass.objResult = JsonConvert.SerializeObject(monthquotaList);
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
        /// 業績目標設定_修改/新增 UpdMonthQuota/Month_quota_editor.asp
        /// </summary>
        [HttpPost("UpdMonthQuota")]
        public ActionResult<ResultClass<string>> UpdMonthQuota(Feat_Target_Upd model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var User_Num = HttpContext.Session.GetString("UserID");
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = "select * from Feat_target where U_num=@U_num and target_ym=@target_ym ";
                parameters.Add(new SqlParameter("@U_num", model.U_num));
                parameters.Add(new SqlParameter("@target_ym", model.target_ym));
                #endregion
                var dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
                if (dtResult.Rows.Count > 0)
                {
                    //修改
                    #region SQL
                    var parameters_u = new List<SqlParameter>();
                    var T_SQL_U = @"
                        update Feat_target set target_quota=@target_quota,edit_date=getdate(),edit_num=@edit_num,edit_ip=@edit_ip
                        where U_num=@U_num and target_ym=@target_ym";
                    parameters_u.Add(new SqlParameter("@target_quota", model.target_quota));
                    parameters_u.Add(new SqlParameter("@edit_num", User_Num));
                    parameters_u.Add(new SqlParameter("@edit_ip", clientIp));
                    parameters_u.Add(new SqlParameter("@U_num", model.U_num));
                    parameters_u.Add(new SqlParameter("@target_ym", model.target_ym));
                    #endregion
                    var result_u = _adoData.ExecuteNonQuery(T_SQL_U, parameters_u);
                    if (result_u > 0)
                    {
                        resultClass.ResultCode = "000";
                        resultClass.ResultMsg = "修改成功";
                        return Ok(resultClass);
                    }
                    else
                    {
                        resultClass.ResultCode = "400";
                        resultClass.ResultMsg = "修改失敗";
                        return BadRequest(resultClass);
                    }
                }
                else
                {
                    //新增
                    #region SQL
                    var parameters_in = new List<SqlParameter>();
                    var T_SQL_IN = @"
                        Insert into Feat_target(target_ym,target_quota,group_id,U_num,del_tag,add_date,add_num,add_ip,edit_date)
                        Values (@target_ym,@target_quota,@group_id,@U_num,'0',GETDATE(),@add_num,@add_ip,GETDATE())";
                    parameters_in.Add(new SqlParameter("@target_ym", model.target_ym));
                    parameters_in.Add(new SqlParameter("@target_quota", model.target_quota));
                    parameters_in.Add(new SqlParameter("@group_id", model.group_M_id));
                    parameters_in.Add(new SqlParameter("@U_num", model.U_num));
                    parameters_in.Add(new SqlParameter("@add_num", User_Num));
                    parameters_in.Add(new SqlParameter("@add_ip", clientIp));
                    #endregion
                    var result_in = _adoData.ExecuteNonQuery(T_SQL_IN, parameters_in);
                    if (result_in > 0)
                    {
                        resultClass.ResultCode = "000";
                        resultClass.ResultMsg = "新增成功";
                        return Ok(resultClass);
                    }
                    else
                    {
                        resultClass.ResultCode = "400";
                        resultClass.ResultMsg = "新增失敗";
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
        /// <summary>
        /// 提供可查詢的分公司資料
        /// </summary>
        [HttpGet("GetUserCheckBCList")]
        public ActionResult<ResultClass<string>> GetUserCheckBCList()
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var User_Num = HttpContext.Session.GetString("UserID");

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL_抓取可查詢的分公司資料
                var parameters = new List<SqlParameter>();
                var T_SQL = @"
                    select item_D_code,item_D_name from Item_list where item_M_code='branch_company' and item_D_type='Y'
                    and item_D_code in (select SplitValue from dbo.SplitStringFunction　((
                    select 'zz'+REPLACE(U_Check_BC,'#',',') from User_M where U_num=@U_num))) order by item_sort";
                parameters.Add(new SqlParameter("@U_num", User_Num));
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
        /// 日報表_顯示個人(2020版)_查詢 Feat_Daily_Person_List_Query/Feat_daily_report_show_user.asp
        /// </summary>
        [HttpPost("Feat_Daily_Person_List_Query")]
        public ActionResult<ResultClass<string>> Feat_Daily_Person_List_Query(FeatDailyPerson_req model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"
                    with A(leader_name,plan_name,group_id,group_M_id,group_M_title,group_M_start_day,group_M_end_day,day_incase_num,month_incase_num
                    ,day_get_amount_num,day_get_amount,month_pass_num,month_get_amount_num,month_get_amount,month_pass_amount,month_pre_amount ) as (SELECT isnull(ug.group_M_name,'未分組') leader_name
                    ,f.plan_name,ug.group_id,ug.group_M_id,ug.group_M_title,ug.group_M_start_day,ug.group_M_end_day";
                //日進件數 
                T_SQL += " ,(case when convert(varchar,Send_amount_date,112) = @reportDate_n  and Send_amount_date between ug.group_M_start_day and ug.group_M_end_day and Send_amount_date between ug.group_D_start_day and ug.group_D_end_day then 1 else 0 end ) as day_incase_num";
                //月進件數
                T_SQL += " ,(case when left(convert(varchar,Send_amount_date,112),6) = LEFT(@reportDate_n,6) AND convert(varchar,Send_amount_date,112) <= @reportDate_n  and Send_amount_date between ug.group_M_start_day and ug.group_M_end_day and Send_amount_date between ug.group_D_start_day and ug.group_D_end_day then 1 else 0 end ) as month_incase_num";
                //日撥款數 
                T_SQL += " ,(case when convert(varchar,get_amount_date,112) = @reportDate_n  and get_amount_date between ug.group_M_start_day and ug.group_M_end_day and get_amount_date between ug.group_D_start_day and ug.group_D_end_day then 1 else 0 end ) as day_get_amount_num";
                //日撥款額 
                T_SQL += " ,(case when convert(varchar,get_amount_date,112) = @reportDate_n  and get_amount_date between ug.group_M_start_day and ug.group_M_end_day and get_amount_date between ug.group_D_start_day and ug.group_D_end_day then get_amount else 0 end ) as day_get_amount";
                //月核准數
                T_SQL += " ,(case when left(convert(varchar,Send_result_date,112),6) = LEFT(@reportDate_n,6) AND convert(varchar,Send_result_date,112) <= @reportDate_n  and Send_result_date between ug.group_M_start_day and ug.group_M_end_day and Send_result_date between ug.group_D_start_day and ug.group_D_end_day AND Send_result_type = 'SRT002'  then 1 else 0 end ) as month_pass_num";
                //月撥款數
                T_SQL += " ,(case when left(convert(varchar,get_amount_date,112),6) = LEFT(@reportDate_n,6) AND convert(varchar,get_amount_date,112) <= @reportDate_n  and get_amount_date between ug.group_M_start_day and ug.group_M_end_day and get_amount_date between ug.group_D_start_day and ug.group_D_end_day then 1 else 0 end ) as month_get_amount_num";
                //月撥款額
                T_SQL += " ,(case when left(convert(varchar,get_amount_date,112),6) = LEFT(@reportDate_n,6) AND convert(varchar,get_amount_date,112) <= @reportDate_n  and get_amount_date between ug.group_M_start_day and ug.group_M_end_day and get_amount_date between ug.group_D_start_day and ug.group_D_end_day then get_amount else 0 end ) as month_get_amount";
                //已核未撥
                T_SQL += " ,(case when left(convert(varchar,Send_result_date,112),6) in (LEFT(@reportDate_n,6),@reportDate_b) AND convert(varchar,Send_result_date,112) <= @reportDate_n  and Send_result_date between ug.group_M_start_day and ug.group_M_end_day and Send_result_date between ug.group_D_start_day and ug.group_D_end_day AND Send_result_type = 'SRT002' AND isnull(check_amount_type,'') NOT IN('CKAT003') AND isnull(get_amount_type,'') NOT IN('GTAT002','GTAT003')  then pass_amount else 0 end ) as month_pass_amount";
                //預核額度
                T_SQL += " ,(case when left(convert(varchar,Send_result_date,112),6) in (LEFT(@reportDate_n,6),@reportDate_b) AND convert(varchar,Send_result_date,112) <= @reportDate_n  and Send_result_date between ug.group_M_start_day and ug.group_M_end_day and Send_result_date between ug.group_D_start_day and ug.group_D_end_day AND Send_result_type = 'SRT005' then pass_amount else 0 end ) as month_pre_amount";
                T_SQL += " FROM viewFeats f LEFT JOIN view_User_group ug ON ug.group_D_code = f.plan_num AND((Send_amount_date between ug.group_M_start_day and ug.group_M_end_day and Send_amount_date between ug.group_D_start_day and ug.group_D_end_day) ";
                T_SQL += " OR(Send_result_date between ug.group_M_start_day and ug.group_M_end_day and Send_result_date between ug.group_D_start_day and ug.group_D_end_day) OR(get_amount_date between ug.group_M_start_day and ug.group_M_end_day and get_amount_date between ug.group_D_start_day and ug.group_D_end_day) )  ";
                T_SQL += " where 1=1  AND(left(convert(varchar,Send_amount_date,112),6) = LEFT(@reportDate_n,6) OR left(convert(varchar,Send_result_date,112),6) in (LEFT(@reportDate_n,6),@reportDate_b) OR left(convert(varchar,get_amount_date,112),6) = LEFT(@reportDate_n,6)  )  and CONVERT(DATE,@reportDate_n,112) between ug.group_M_start_day and ug.group_M_end_day ";
                T_SQL += " AND f.fund_company IN (SELECT SplitValue FROM dbo.SplitStringFunction(@company)) AND U_BC = @U_BC )";
                T_SQL += " select @U_BC U_BC,leader_name,plan_name,group_M_id,group_M_title,sum(day_incase_num) as day_incase_num,sum(month_incase_num) as month_incase_num,sum(day_incase_num) as day_incase_num,sum(day_get_amount_num) as day_get_amount_num,sum(day_get_amount) as day_get_amount,sum(month_pass_num) as month_pass_num,sum(month_get_amount_num) as month_get_amount_num,sum(month_get_amount) as month_get_amount,sum(month_pass_amount) as month_pass_amount,sum(month_pre_amount) as month_pre_amount  FROM A  group by leader_name,plan_name,group_M_id,group_M_title  union";
                //顯示沒業績的組員
                T_SQL += " select @U_BC U_BC,leader_name,plan_name,group_M_id,group_M_title,0 as day_incase_num,0 as month_incase_num,0 as day_incase_num,0 as day_get_amount_num,0 as day_get_amount,0 as month_pass_num,0 as month_get_amount_num,0 as month_get_amount,0 as month_pass_amount,0 as month_pre_amount  from (select group_M_name leader_name,group_D_name plan_name,group_M_id,group_M_title  from view_User_group ug  where group_M_id in(select distinct group_M_id FROM A where @reportDate_n between A.group_M_start_day and A.group_M_end_day)  and group_id not in(select distinct group_id FROM A)   group by group_M_name,group_D_name,group_M_id,group_M_title ) B order by leader_name,plan_name";
                parameters.Add(new SqlParameter("@reportDate_n", model.reportDate_n));
                parameters.Add(new SqlParameter("@company", model.company));
                parameters.Add(new SqlParameter("@U_BC", model.U_BC));
                string reportDate_b = DateTime.ParseExact(model.reportDate_n, "yyyyMMdd", null).AddMonths(-1).ToString("yyyyMM");
                parameters.Add(new SqlParameter("@reportDate_b", reportDate_b));
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
        /// 日報表_(202210版)_查詢 Feat_Daily_Report_V2022_Query/Feat_daily_report_v2021_202210.asp
        /// </summary>
        [HttpPost("Feat_Daily_Report_V2022_Query")]
        public ActionResult<ResultClass<string>> Feat_Daily_Report_V2022_Query(FeatDailyReport_req model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"
                   with A ( leader_name,plan_name,plan_num,group_id,group_M_id,group_M_title ,U_PFT_sort,U_PFT_name ,day_incase_num_FDCOM001,month_incase_num_FDCOM001 
                   ,day_incase_num_FDCOM003,month_incase_num_FDCOM003 ,day_incase_num_FDCOM004,month_incase_num_FDCOM004 ,day_incase_num_FDCOM005,month_incase_num_FDCOM005 
                   ,day_get_amount_num,day_get_amount,month_pass_num,month_get_amount_num ,month_get_amount_FDCOM001,month_get_amount_FDCOM003,month_get_amount_FDCOM004,month_get_amount_FDCOM005,month_pass_amount_FDCOM001,month_pre_amount_FDCOM001,month_pass_amount_FDCOM003,month_pass_amount_FDCOM004,month_pass_amount_FDCOM005,advance_payment_AE ) as ( SELECT isnull(ug.group_M_name,'未分組') leader_name
                   ,ug.group_D_name,ug.group_D_code, ug.group_id, ug.group_M_id,ug.group_M_title ,sa.U_PFT_sort,sa.U_PFT_name";
                //新鑫 日進件數 
                T_SQL += " ,sum(case when 'FDCOM001'=fund_company and convert(varchar, Send_amount_date, 112) = @reportDate_n then 1 else 0 end ) as day_incase_num_FDCOM001";
                //新鑫 月進件數
                T_SQL += " ,sum(case when 'FDCOM001'=fund_company and left(convert(varchar, Send_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, Send_amount_date, 112) <= @reportDate_n then 1 else 0 end ) as month_incase_num_FDCOM001";
                //國&#23791; 日進件數
                T_SQL += " ,sum(case when 'FDCOM003'=fund_company and convert(varchar, Send_amount_date, 112) = @reportDate_n then 1 else 0 end ) as day_incase_num_FDCOM003";
                //國&#23791; 月進件數
                T_SQL += " ,sum(case when 'FDCOM003'=fund_company and left(convert(varchar, Send_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, Send_amount_date, 112) <= @reportDate_n then 1 else 0 end ) as month_incase_num_FDCOM003";
                //和潤 日進件數
                T_SQL += " ,sum(case when 'FDCOM004'=fund_company and convert(varchar, Send_amount_date, 112) = @reportDate_n then 1 else 0 end ) as day_incase_num_FDCOM004";
                //和潤 月進件數
                T_SQL += " ,sum(case when 'FDCOM004'=fund_company and left(convert(varchar, Send_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, Send_amount_date, 112) <= @reportDate_n then 1 else 0 end ) as month_incase_num_FDCOM004";
                //福斯 日進件數
                T_SQL += " ,sum(case when 'FDCOM005'=fund_company and convert(varchar, Send_amount_date, 112) = @reportDate_n then 1 else 0 end ) as day_incase_num_FDCOM005";
                //福斯 月進件數
                T_SQL += " ,sum(case when 'FDCOM005'=fund_company and left(convert(varchar, Send_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, Send_amount_date, 112) <= @reportDate_n then 1 else 0 end ) as month_incase_num_FDCOM005";
                //日撥款數
                T_SQL += " ,sum(case when convert(varchar, get_amount_date, 112) = @reportDate_n  then 1 else 0 end ) as day_get_amount_num";
                //日撥款額
                T_SQL += " ,sum(case when convert(varchar, get_amount_date, 112) = @reportDate_n  then get_amount else 0 end ) as day_get_amount";
                //月核准數 
                T_SQL += " ,sum(case when left(convert(varchar, Send_result_date, 112),6)=LEFT(@reportDate_n,6) AND convert(varchar, Send_result_date, 112) <= @reportDate_n  AND Send_result_type in ('SRT002','SRT005') then 1 else 0 end ) as month_pass_num";
                //月撥款數
                T_SQL += " ,sum(case when left(convert(varchar, get_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, get_amount_date, 112) <= @reportDate_n  then 1 else 0 end ) as month_get_amount_num";
                //新鑫 月撥款額
                T_SQL += " ,sum(case when 'FDCOM001'=fund_company and left(convert(varchar, get_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, get_amount_date, 112) <= @reportDate_n  then get_amount else 0 end ) as month_get_amount_FDCOM001";
                //國&#23791; 月撥款額
                T_SQL += " ,sum(case when 'FDCOM003'=fund_company and left(convert(varchar, get_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, get_amount_date, 112) <= @reportDate_n  then get_amount else 0 end ) as month_get_amount_FDCOM003";
                //和潤 月撥款額
                T_SQL += " ,sum(case when 'FDCOM004'=fund_company and left(convert(varchar, get_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, get_amount_date, 112) <= @reportDate_n  then get_amount else 0 end ) as month_get_amount_FDCOM004";
                //福斯 月撥款額
                T_SQL += " ,sum(case when 'FDCOM005'=fund_company and left(convert(varchar, get_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, get_amount_date, 112) <= @reportDate_n  then get_amount else 0 end ) as month_get_amount_FDCOM005";
                //新鑫 已核未撥
                T_SQL += " ,sum(case when 'FDCOM001'=fund_company and left(convert(varchar, Send_result_date, 112),6) in (LEFT(@reportDate_n,6),@reportDate_b) AND convert(varchar, Send_result_date, 112) <= @reportDate_n  AND Send_result_type = 'SRT002' AND isnull(check_amount_type,'') NOT IN ('CKAT003') AND isnull(get_amount_type,'') NOT IN ('GTAT002','GTAT003') then pass_amount else 0 end ) as month_pass_amount_FDCOM001";
                //新鑫 預核額度
                T_SQL += " ,sum(case when 'FDCOM001'=fund_company and left(convert(varchar, Send_result_date, 112),6) in (LEFT(@reportDate_n,6),@reportDate_b) AND convert(varchar, Send_result_date, 112) <= @reportDate_n  AND Send_result_type = 'SRT005' then pass_amount else 0 end ) as month_pre_amount_FDCOM001";
                //國&#23791; 已核未撥
                T_SQL += " ,sum(case when 'FDCOM003'=fund_company and left(convert(varchar, Send_result_date, 112),6) in (LEFT(@reportDate_n,6),@reportDate_b) AND convert(varchar, Send_result_date, 112) <= @reportDate_n  AND Send_result_type = 'SRT002' AND isnull(check_amount_type,'') NOT IN ('CKAT003') AND isnull(get_amount_type,'') NOT IN ('GTAT002','GTAT003') then pass_amount else 0 end ) as month_pass_amount_FDCOM003";
                //和潤 已核未撥
                T_SQL += " ,sum(case when 'FDCOM004'=fund_company and left(convert(varchar, Send_result_date, 112),6) in (LEFT(@reportDate_n,6),@reportDate_b) AND convert(varchar, Send_result_date, 112) <= @reportDate_n  AND Send_result_type = 'SRT002' AND isnull(check_amount_type,'') NOT IN ('CKAT003') AND isnull(get_amount_type,'') NOT IN ('GTAT002','GTAT003') then pass_amount else 0 end ) as month_pass_amount_FDCOM004";
                //福斯 已核未撥
                T_SQL += " ,sum(case when 'FDCOM005'=fund_company and left(convert(varchar, Send_result_date, 112),6) in (LEFT(@reportDate_n,6),@reportDate_b) AND convert(varchar, Send_result_date, 112) <= @reportDate_n  AND Send_result_type = 'SRT002' AND isnull(check_amount_type,'') NOT IN ('CKAT003') AND isnull(get_amount_type,'') NOT IN ('GTAT002','GTAT003') then pass_amount else 0 end ) as month_pass_amount_FDCOM005";
                //國&#23791; 代墊款(萬) 
                T_SQL += " ,sum(case when 'FDCOM003'=fund_company and left(convert(varchar, get_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, get_amount_date, 112) <= @reportDate_n  then advance_payment_AE else 0 end ) as advance_payment_AE";
                T_SQL += " FROM view_User_group ug join view_User_sales leader on leader.U_num = ug.group_M_code AND leader.U_BC = @U_BC ";
                T_SQL += " join view_User_sales sa on sa.U_num = ug.group_D_code ";
                T_SQL += " left join viewFeats f on ug.group_D_code = f.plan_num AND ( left(convert(varchar, f.Send_amount_date, 112),6) = LEFT(@reportDate_n,6) OR left(convert(varchar, f.Send_result_date, 112),6) in (LEFT(@reportDate_n,6),@reportDate_b) OR left(convert(varchar, f.get_amount_date, 112),6) = LEFT(@reportDate_n,6) ) ";
                T_SQL += " where 1=1 and @reportDate_n between ug.group_M_start_day and ug.group_M_end_day and @reportDate_n between ug.group_D_start_day and ug.group_D_end_day  group by isnull(ug.group_M_name,'未分組'),ug.group_D_name,ug.group_D_code, ug.group_id, ug.group_M_id,ug.group_M_title,sa.U_PFT_sort,sa.U_PFT_name)";
                T_SQL += " select @U_BC U_BC,a.* ,isnull(ft.target_quota,0) target_quota FROM A  ";
                T_SQL += " left join Feat_target ft on ft.del_tag='0'  and ft.U_num=A.plan_num  and ft.group_id=A.group_id and ft.target_ym=LEFT(@reportDate_n,6) ";
                T_SQL += " order by A.leader_name,A.U_PFT_sort,A.U_PFT_name,A.plan_num";
                parameters.Add(new SqlParameter("@reportDate_n", model.reportDate_n));
                parameters.Add(new SqlParameter("@U_BC", model.U_BC));
                string reportDate_b = DateTime.ParseExact(model.reportDate_n, "yyyyMMdd", null).AddMonths(-1).ToString("yyyyMM");
                parameters.Add(new SqlParameter("@reportDate_b", reportDate_b));
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
        /// 日報表_(202210版)_下載 Feat_Daily_Report_V2022_Excel/Feat_daily_report_v2021_202210.asp
        /// </summary>
        [HttpPost("Feat_Daily_Report_V2022_Excel")]
        public IActionResult Feat_Daily_Report_V2022_Excel(string reportDate_n)
        {
            var User_Num = HttpContext.Session.GetString("UserID");

            try
            {
                FeatDailyReport_excel_Total modelFooter = new FeatDailyReport_excel_Total();
                ADOData _adoData = new ADOData();
                #region SQL_抓取可查詢的分公司資料
                var parameters_comp = new List<SqlParameter>();
                var T_SQL_COMP = @"
                    select item_D_code,item_D_name from Item_list where item_M_code='branch_company' and item_D_type='Y'
                    and item_D_code in (select SplitValue from dbo.SplitStringFunction　((
                    select 'zz'+REPLACE(U_Check_BC,'#',',') from User_M where U_num=@U_num))) order by item_sort";
                parameters_comp.Add(new SqlParameter("@U_num", User_Num));
                #endregion
                var dtResultComp = _adoData.ExecuteQuery(T_SQL_COMP, parameters_comp);
                if (dtResultComp.Rows.Count > 0)
                {
                    byte[] fileBytes = null;

                    for (int i = 0; i < dtResultComp.Rows.Count; i++)
                    {
                        string itemDCode = dtResultComp.Rows[i]["item_D_code"].ToString();
                        string itemDName = dtResultComp.Rows[i]["item_D_name"].ToString();
                        #region SQL
                        var parameters = new List<SqlParameter>();
                        var T_SQL = @"
                            with A ( leader_name,plan_name,plan_num,group_id,group_M_id,group_M_title ,U_PFT_sort,U_PFT_name ,day_incase_num_FDCOM001,month_incase_num_FDCOM001
                            ,day_incase_num_FDCOM003,month_incase_num_FDCOM003 ,day_incase_num_FDCOM004,month_incase_num_FDCOM004 ,day_incase_num_FDCOM005,month_incase_num_FDCOM005
                            ,day_get_amount_num,day_get_amount,month_pass_num,month_get_amount_num ,month_get_amount_FDCOM001,month_get_amount_FDCOM003,month_get_amount_FDCOM004,month_get_amount_FDCOM005,month_pass_amount_FDCOM001,month_pre_amount_FDCOM001,month_pass_amount_FDCOM003,month_pass_amount_FDCOM004,month_pass_amount_FDCOM005,advance_payment_AE ) as ( SELECT isnull(ug.group_M_name,'未分組') leader_name
                            ,ug.group_D_name,ug.group_D_code, ug.group_id, ug.group_M_id,ug.group_M_title ,sa.U_PFT_sort,sa.U_PFT_name";
                        //新鑫 日進件數 
                        T_SQL += " ,sum(case when 'FDCOM001'=fund_company and convert(varchar, Send_amount_date, 112) = @reportDate_n then 1 else 0 end ) as day_incase_num_FDCOM001";
                        //新鑫 月進件數
                        T_SQL += " ,sum(case when 'FDCOM001'=fund_company and left(convert(varchar, Send_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, Send_amount_date, 112) <= @reportDate_n then 1 else 0 end ) as month_incase_num_FDCOM001";
                        //國&#23791; 日進件數
                        T_SQL += " ,sum(case when 'FDCOM003'=fund_company and convert(varchar, Send_amount_date, 112) = @reportDate_n then 1 else 0 end ) as day_incase_num_FDCOM003";
                        //國&#23791; 月進件數
                        T_SQL += " ,sum(case when 'FDCOM003'=fund_company and left(convert(varchar, Send_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, Send_amount_date, 112) <= @reportDate_n then 1 else 0 end ) as month_incase_num_FDCOM003";
                        //和潤 日進件數
                        T_SQL += " ,sum(case when 'FDCOM004'=fund_company and convert(varchar, Send_amount_date, 112) = @reportDate_n then 1 else 0 end ) as day_incase_num_FDCOM004";
                        //和潤 月進件數
                        T_SQL += " ,sum(case when 'FDCOM004'=fund_company and left(convert(varchar, Send_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, Send_amount_date, 112) <= @reportDate_n then 1 else 0 end ) as month_incase_num_FDCOM004";
                        //福斯 日進件數
                        T_SQL += " ,sum(case when 'FDCOM005'=fund_company and convert(varchar, Send_amount_date, 112) = @reportDate_n then 1 else 0 end ) as day_incase_num_FDCOM005";
                        //福斯 月進件數
                        T_SQL += " ,sum(case when 'FDCOM005'=fund_company and left(convert(varchar, Send_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, Send_amount_date, 112) <= @reportDate_n then 1 else 0 end ) as month_incase_num_FDCOM005";
                        //日撥款數
                        T_SQL += " ,sum(case when convert(varchar, get_amount_date, 112) = @reportDate_n  then 1 else 0 end ) as day_get_amount_num";
                        //日撥款額
                        T_SQL += " ,sum(case when convert(varchar, get_amount_date, 112) = @reportDate_n  then get_amount else 0 end ) as day_get_amount";
                        //月核准數 
                        T_SQL += " ,sum(case when left(convert(varchar, Send_result_date, 112),6)=LEFT(@reportDate_n,6) AND convert(varchar, Send_result_date, 112) <= @reportDate_n  AND Send_result_type in ('SRT002','SRT005') then 1 else 0 end ) as month_pass_num";
                        //月撥款數
                        T_SQL += " ,sum(case when left(convert(varchar, get_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, get_amount_date, 112) <= @reportDate_n  then 1 else 0 end ) as month_get_amount_num";
                        //新鑫 月撥款額
                        T_SQL += " ,sum(case when 'FDCOM001'=fund_company and left(convert(varchar, get_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, get_amount_date, 112) <= @reportDate_n  then get_amount else 0 end ) as month_get_amount_FDCOM001";
                        //國&#23791; 月撥款額
                        T_SQL += " ,sum(case when 'FDCOM003'=fund_company and left(convert(varchar, get_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, get_amount_date, 112) <= @reportDate_n  then get_amount else 0 end ) as month_get_amount_FDCOM003";
                        //和潤 月撥款額
                        T_SQL += " ,sum(case when 'FDCOM004'=fund_company and left(convert(varchar, get_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, get_amount_date, 112) <= @reportDate_n  then get_amount else 0 end ) as month_get_amount_FDCOM004";
                        //福斯 月撥款額
                        T_SQL += " ,sum(case when 'FDCOM005'=fund_company and left(convert(varchar, get_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, get_amount_date, 112) <= @reportDate_n  then get_amount else 0 end ) as month_get_amount_FDCOM005";
                        //新鑫 已核未撥
                        T_SQL += " ,sum(case when 'FDCOM001'=fund_company and left(convert(varchar, Send_result_date, 112),6) in (LEFT(@reportDate_n,6),@reportDate_b) AND convert(varchar, Send_result_date, 112) <= @reportDate_n  AND Send_result_type = 'SRT002' AND isnull(check_amount_type,'') NOT IN ('CKAT003') AND isnull(get_amount_type,'') NOT IN ('GTAT002','GTAT003') then pass_amount else 0 end ) as month_pass_amount_FDCOM001";
                        //新鑫 預核額度
                        T_SQL += " ,sum(case when 'FDCOM001'=fund_company and left(convert(varchar, Send_result_date, 112),6) in (LEFT(@reportDate_n,6),@reportDate_b) AND convert(varchar, Send_result_date, 112) <= @reportDate_n  AND Send_result_type = 'SRT005' then pass_amount else 0 end ) as month_pre_amount_FDCOM001";
                        //國&#23791; 已核未撥
                        T_SQL += " ,sum(case when 'FDCOM003'=fund_company and left(convert(varchar, Send_result_date, 112),6) in (LEFT(@reportDate_n,6),@reportDate_b) AND convert(varchar, Send_result_date, 112) <= @reportDate_n  AND Send_result_type = 'SRT002' AND isnull(check_amount_type,'') NOT IN ('CKAT003') AND isnull(get_amount_type,'') NOT IN ('GTAT002','GTAT003') then pass_amount else 0 end ) as month_pass_amount_FDCOM003";
                        //和潤 已核未撥
                        T_SQL += " ,sum(case when 'FDCOM004'=fund_company and left(convert(varchar, Send_result_date, 112),6) in (LEFT(@reportDate_n,6),@reportDate_b) AND convert(varchar, Send_result_date, 112) <= @reportDate_n  AND Send_result_type = 'SRT002' AND isnull(check_amount_type,'') NOT IN ('CKAT003') AND isnull(get_amount_type,'') NOT IN ('GTAT002','GTAT003') then pass_amount else 0 end ) as month_pass_amount_FDCOM004";
                        //福斯 已核未撥
                        T_SQL += " ,sum(case when 'FDCOM005'=fund_company and left(convert(varchar, Send_result_date, 112),6) in (LEFT(@reportDate_n,6),@reportDate_b) AND convert(varchar, Send_result_date, 112) <= @reportDate_n  AND Send_result_type = 'SRT002' AND isnull(check_amount_type,'') NOT IN ('CKAT003') AND isnull(get_amount_type,'') NOT IN ('GTAT002','GTAT003') then pass_amount else 0 end ) as month_pass_amount_FDCOM005";
                        //國&#23791; 代墊款(萬) 
                        T_SQL += " ,sum(case when 'FDCOM003'=fund_company and left(convert(varchar, get_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, get_amount_date, 112) <= @reportDate_n  then advance_payment_AE else 0 end ) as advance_payment_AE";
                        T_SQL += " FROM view_User_group ug join view_User_sales leader on leader.U_num = ug.group_M_code AND leader.U_BC = @U_BC ";
                        T_SQL += " join view_User_sales sa on sa.U_num = ug.group_D_code ";
                        T_SQL += " left join viewFeats f on ug.group_D_code = f.plan_num AND ( left(convert(varchar, f.Send_amount_date, 112),6) = LEFT(@reportDate_n,6) OR left(convert(varchar, f.Send_result_date, 112),6) in (LEFT(@reportDate_n,6),@reportDate_b) OR left(convert(varchar, f.get_amount_date, 112),6) = LEFT(@reportDate_n,6) ) ";
                        T_SQL += " where 1=1 and @reportDate_n between ug.group_M_start_day and ug.group_M_end_day and @reportDate_n between ug.group_D_start_day and ug.group_D_end_day  group by isnull(ug.group_M_name,'未分組'),ug.group_D_name,ug.group_D_code, ug.group_id, ug.group_M_id,ug.group_M_title,sa.U_PFT_sort,sa.U_PFT_name)";
                        T_SQL += " select @U_BC U_BC,a.* ,isnull(ft.target_quota,0) target_quota FROM A  ";
                        T_SQL += " left join Feat_target ft on ft.del_tag='0'  and ft.U_num=A.plan_num  and ft.group_id=A.group_id and ft.target_ym=LEFT(@reportDate_n,6) ";
                        T_SQL += " order by A.leader_name,A.U_PFT_sort,A.U_PFT_name,A.plan_num";
                        parameters.Add(new SqlParameter("@reportDate_n", reportDate_n));
                        parameters.Add(new SqlParameter("@U_BC", itemDCode));
                        string reportDate_b = DateTime.ParseExact(reportDate_n, "yyyyMMdd", null).AddMonths(-1).ToString("yyyyMM");
                        parameters.Add(new SqlParameter("@reportDate_b", reportDate_b));
                        #endregion
                        var dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
                        if (dtResult.Rows.Count > 0)
                        {
                            var excelList = dtResult.AsEnumerable().Select(row => new FeatDailyReport_excel
                            {
                                plan_name = row.Field<string>("plan_name"),
                                U_PFT_name = row.Field<string>("U_PFT_name"),
                                day_incase_num_FDCOM001 = row.Field<int>("day_incase_num_FDCOM001"),
                                month_incase_num_FDCOM001 = row.Field<int>("month_incase_num_FDCOM001"),
                                day_incase_num_FDCOM003 = row.Field<int>("day_incase_num_FDCOM003"),
                                month_incase_num_FDCOM003 = row.Field<int>("month_incase_num_FDCOM003"),
                                day_get_amount_num = row.Field<int>("day_get_amount_num"),
                                day_get_amount = row.Field<int>("day_get_amount"),
                                month_pass_num = row.Field<int>("month_pass_num"),
                                month_get_amount_num = row.Field<int>("month_get_amount_num"),
                                month_get_amount_FDCOM001 = row.Field<int>("month_get_amount_FDCOM001"),
                                month_get_amount_FDCOM003 = row.Field<int>("month_get_amount_FDCOM003"),
                                month_pass_amount_FDCOM001 = row.Field<int>("month_pass_amount_FDCOM001"),
                                month_pre_amount_FDCOM001 = row.Field<int>("month_pre_amount_FDCOM001"),
                                month_pass_amount_FDCOM003 = row.Field<int>("month_pass_amount_FDCOM003"),
                                advance_payment_AE = row.Field<decimal>("advance_payment_AE"),
                                target_quota = Convert.ToInt32(row.Field<short>("target_quota"))
                            }).ToList();
                            var Excel_Headers = new Dictionary<string, string>
                            {
                                { "U_PFT_name", "職位" },
                                { "plan_name", "業務" },
                                { "day_incase_num_FDCOM001", "新鑫日進件數" },
                                { "month_incase_num_FDCOM001", "新鑫累積進件數" },
                                { "day_incase_num_FDCOM003", "國峯日進件數" },
                                { "month_incase_num_FDCOM003", "國峯累積進件數" },
                                { "day_get_amount_num", "日撥件數" },
                                { "day_get_amount", "日撥金額" },
                                { "month_pass_num", "核准件數" },
                                { "month_get_amount_num", "累積撥款件數" },
                                { "month_get_amount_FDCOM001", "新鑫-撥款金額(萬)" },
                                { "month_get_amount_FDCOM003", "國峯-撥款金額(萬)" },
                                { "month_pass_amount_FDCOM001", "新鑫已核未撥" },
                                { "month_pre_amount_FDCOM001", "新鑫預核額度" },
                                { "month_pass_amount_FDCOM003", "國峯已核未撥" },
                                { "advance_payment_AE", "國峯-代墊款(萬)" },
                                { "target_quota", "目標(萬)" },
                                { "AchievementRate", "達成率" }
                            };
                            if (i == 0)
                            {
                                fileBytes = FuncHandler.FeatDailyToExcel(excelList, Excel_Headers, itemDName, reportDate_n);
                            }
                            else
                            {
                                fileBytes = FuncHandler.FeatDailyToExcelAgain(fileBytes, excelList, Excel_Headers, itemDName, reportDate_n);
                            }
                            #region 總計 FeatDailyToExcelFooter
                            var total_D_FDCOM001 = excelList.Sum(item => item.day_incase_num_FDCOM001);
                            modelFooter.day_incase_num_FDCOM001_total += total_D_FDCOM001;
                            var total_M_FDCOM001 = excelList.Sum(item => item.month_incase_num_FDCOM001);
                            modelFooter.month_incase_num_FDCOM001_total += total_M_FDCOM001;
                            var total_D_FDCOM003 = excelList.Sum(item => item.day_incase_num_FDCOM003);
                            modelFooter.day_incase_num_FDCOM003_total += total_D_FDCOM003;
                            var total_M_FDCOM003 = excelList.Sum(item => item.month_incase_num_FDCOM003);
                            modelFooter.month_incase_num_FDCOM003_total += total_M_FDCOM003;
                            var total_D_AmoutNum = excelList.Sum(item => item.day_get_amount_num);
                            modelFooter.day_get_amount_num_total += total_D_AmoutNum;
                            var total_S_Amout = excelList.Sum(item => item.day_get_amount);
                            modelFooter.day_get_amount_total += total_S_Amout;
                            var totol_M_PassNum = excelList.Sum(item => item.month_pass_num);
                            modelFooter.month_pass_num_total += totol_M_PassNum;
                            var total_M_AmoutNum = excelList.Sum(item => item.month_get_amount_num);
                            modelFooter.month_get_amount_num_total += total_M_AmoutNum;
                            var total_M_Amout_FDCOM001 = excelList.Sum(item => item.month_get_amount_FDCOM001);
                            modelFooter.month_get_amount_FDCOM001_total += total_M_Amout_FDCOM001;
                            var total_M_Amout_FDCOM003 = excelList.Sum(item => item.month_get_amount_FDCOM003);
                            modelFooter.month_get_amount_FDCOM003_total += total_M_Amout_FDCOM003;
                            var total_M_PassAmout_FDCOM001 = excelList.Sum(item => item.month_pass_amount_FDCOM001);
                            modelFooter.month_pass_amount_FDCOM001_total += total_M_PassAmout_FDCOM001;
                            var total_M_PreAmout_FDCOM001 = excelList.Sum(item => item.month_pre_amount_FDCOM001);
                            modelFooter.month_pre_amount_FDCOM001_total += total_M_PreAmout_FDCOM001;
                            var total_M_PassAmout_FDCOM003 = excelList.Sum(item => item.month_pass_amount_FDCOM003);
                            modelFooter.month_pass_amount_FDCOM003_total += total_M_PassAmout_FDCOM003;
                            var total_Pay = excelList.Sum(item => item.advance_payment_AE);
                            modelFooter.advance_payment_AE_total += total_Pay;
                            var total_Quota = excelList.Sum(item => item.target_quota);
                            modelFooter.target_quota_total += total_Quota;
                            #endregion
                        }
                    }
                    //全區總計 FeatDailyToExcelFooter
                    var Excel_Footer = new Dictionary<string, string>
                    {
                        { "U_PFT_name", "職位" },
                        { "plan_name", "業務" },
                        { "day_incase_num_FDCOM001_total", "新鑫日進件數" },
                        { "month_incase_num_FDCOM001_total", "新鑫累積進件數" },
                        { "day_incase_num_FDCOM003_total", "國峯日進件數" },
                        { "month_incase_num_FDCOM003_total", "國峯累積進件數" },
                        { "day_get_amount_num_total", "日撥件數" },
                        { "day_get_amount_total", "日撥金額" },
                        { "month_pass_num_total", "核准件數" },
                        { "month_get_amount_num_total", "累積撥款件數" },
                        { "month_get_amount_FDCOM001_total", "新鑫-撥款金額(萬)" },
                        { "month_get_amount_FDCOM003_total", "國峯-撥款金額(萬)" },
                        { "month_pass_amount_FDCOM001_total", "新鑫已核未撥" },
                        { "month_pre_amount_FDCOM001_total", "新鑫預核額度" },
                        { "month_pass_amount_FDCOM003_total", "國峯已核未撥" },
                        { "advance_payment_AE_total", "國峯-代墊款(萬)" },
                        { "target_quota_total", "目標(萬)" },
                        { "AchievementRate", "達成率" }
                    };
                    fileBytes = FuncHandler.FeatDailyToExcelFooter(fileBytes, modelFooter, Excel_Footer, reportDate_n);
                    var fileName = "日報表" + DateTime.Now.ToString("yyyyMMddHHmm") + ".xlsx";
                    return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
                }
                else
                {
                    return NotFound(); // 檔案不存在時返回 404
                }
            }
            catch (Exception ex)
            {
                ResultClass<string> resultClass = new ResultClass<string>();
                resultClass.ResultCode = "500";
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass);
            }
        }
        #endregion

        #region 業績報表_日報表(202106合計版)
        /// <summary>
        /// 提供所有的分公司資料
        /// </summary>
        [HttpGet("GetUserCheckBCListAll")]
        public ActionResult<ResultClass<string>> GetUserCheckBCListAll()
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL_抓取可查詢的分公司資料
                var parameters = new List<SqlParameter>();
                var T_SQL = @"
                    select item_D_code,item_D_name,COUNT(*) from view_User_group ug
                    join view_User_sales leader on leader.U_num = ug.group_M_code
                    left join Item_list on Item_list.item_D_code=leader.U_BC and item_M_code='branch_company' and item_D_type='Y'
                    group by item_D_code,item_D_name,item_sort order by item_sort";
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
        /// 業績報表_日報表(202106合計版)_查詢 Feat_daily_report_v2021_Query/Feat_daily_report_v2021.asp
        /// </summary>
        [HttpPost("Feat_Daily_Report_v2021_Query")]
        public ActionResult<ResultClass<string>> Feat_Daily_Report_v2021_Query(FeatDailyReport_req model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"
                    with A ( leader_name,plan_name,plan_num,group_id,group_M_id,group_M_title ,U_PFT_sort,U_PFT_name ,day_incase_num_FDCOM001,month_incase_num_FDCOM001 
                    ,day_incase_num_FDCOM003,month_incase_num_FDCOM003 ,day_incase_num_FDCOM004,month_incase_num_FDCOM004 ,day_incase_num_FDCOM005,month_incase_num_FDCOM005 
                    ,day_get_amount_num,day_get_amount,month_pass_num,month_get_amount_num ,month_get_amount_FDCOM001,month_get_amount_FDCOM003,month_get_amount_FDCOM004,month_get_amount_FDCOM005,month_pass_amount_FDCOM001,month_pre_amount_FDCOM001,month_pass_amount_FDCOM003,month_pass_amount_FDCOM004,month_pass_amount_FDCOM005,advance_payment_AE ) as ( SELECT isnull(ug.group_M_name,'未分組') leader_name
                    ,ug.group_D_name,ug.group_D_code, ug.group_id, ug.group_M_id,ug.group_M_title ,sa.U_PFT_sort,sa.U_PFT_name";
                //新鑫 日進件數 
                T_SQL += " ,sum(case when 'FDCOM001'=fund_company and convert(varchar, Send_amount_date, 112) = @reportDate_n then 1 else 0 end ) as day_incase_num_FDCOM001";
                //新鑫 月進件數
                T_SQL += " ,sum(case when 'FDCOM001'=fund_company and left(convert(varchar, Send_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, Send_amount_date, 112) <= @reportDate_n then 1 else 0 end ) as month_incase_num_FDCOM001";
                //國&#23791; 日進件數
                T_SQL += " ,sum(case when 'FDCOM003'=fund_company and convert(varchar, Send_amount_date, 112) = @reportDate_n then 1 else 0 end ) as day_incase_num_FDCOM003";
                //國&#23791; 月進件數
                T_SQL += " ,sum(case when 'FDCOM003'=fund_company and left(convert(varchar, Send_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, Send_amount_date, 112) <= @reportDate_n then 1 else 0 end ) as month_incase_num_FDCOM003";
                //和潤 日進件數
                T_SQL += " ,sum(case when 'FDCOM004'=fund_company and convert(varchar, Send_amount_date, 112) = @reportDate_n then 1 else 0 end ) as day_incase_num_FDCOM004";
                //和潤 月進件數
                T_SQL += " ,sum(case when 'FDCOM004'=fund_company and left(convert(varchar, Send_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, Send_amount_date, 112) <= @reportDate_n then 1 else 0 end ) as month_incase_num_FDCOM004";
                //福斯 日進件數
                T_SQL += " ,sum(case when 'FDCOM005'=fund_company and convert(varchar, Send_amount_date, 112) = @reportDate_n then 1 else 0 end ) as day_incase_num_FDCOM005";
                //福斯 月進件數
                T_SQL += " ,sum(case when 'FDCOM005'=fund_company and left(convert(varchar, Send_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, Send_amount_date, 112) <= @reportDate_n then 1 else 0 end ) as month_incase_num_FDCOM005";
                //日撥款數
                T_SQL += " ,sum(case when convert(varchar, get_amount_date, 112) = @reportDate_n  then 1 else 0 end ) as day_get_amount_num";
                //日撥款額
                T_SQL += " ,sum(case when convert(varchar, get_amount_date, 112) = @reportDate_n  then get_amount else 0 end ) as day_get_amount";
                //月核准數 
                T_SQL += " ,sum(case when left(convert(varchar, Send_result_date, 112),6)=LEFT(@reportDate_n,6) AND convert(varchar, Send_result_date, 112) <= @reportDate_n  AND Send_result_type in ('SRT002','SRT005') then 1 else 0 end ) as month_pass_num";
                //月撥款數
                T_SQL += " ,sum(case when left(convert(varchar, get_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, get_amount_date, 112) <= @reportDate_n  then 1 else 0 end ) as month_get_amount_num";
                //新鑫 月撥款額
                T_SQL += " ,sum(case when 'FDCOM001'=fund_company and left(convert(varchar, get_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, get_amount_date, 112) <= @reportDate_n  then get_amount else 0 end ) as month_get_amount_FDCOM001";
                //國&#23791; 月撥款額
                T_SQL += " ,sum(case when 'FDCOM003'=fund_company and left(convert(varchar, get_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, get_amount_date, 112) <= @reportDate_n  then get_amount else 0 end ) as month_get_amount_FDCOM003";
                //和潤 月撥款額
                T_SQL += " ,sum(case when 'FDCOM004'=fund_company and left(convert(varchar, get_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, get_amount_date, 112) <= @reportDate_n  then get_amount else 0 end ) as month_get_amount_FDCOM004";
                //福斯 月撥款額
                T_SQL += " ,sum(case when 'FDCOM005'=fund_company and left(convert(varchar, get_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, get_amount_date, 112) <= @reportDate_n  then get_amount else 0 end ) as month_get_amount_FDCOM005";
                //新鑫 已核未撥
                T_SQL += " ,sum(case when 'FDCOM001'=fund_company and left(convert(varchar, Send_result_date, 112),6) in (LEFT(@reportDate_n,6),@reportDate_b) AND convert(varchar, Send_result_date, 112) <= @reportDate_n  AND Send_result_type = 'SRT002' AND isnull(check_amount_type,'') NOT IN ('CKAT003') AND isnull(get_amount_type,'') NOT IN ('GTAT002','GTAT003') then pass_amount else 0 end ) as month_pass_amount_FDCOM001";
                //新鑫 預核額度
                T_SQL += " ,sum(case when 'FDCOM001'=fund_company and left(convert(varchar, Send_result_date, 112),6) in (LEFT(@reportDate_n,6),@reportDate_b) AND convert(varchar, Send_result_date, 112) <= @reportDate_n  AND Send_result_type = 'SRT005' then pass_amount else 0 end ) as month_pre_amount_FDCOM001";
                //國&#23791; 已核未撥
                T_SQL += " ,sum(case when 'FDCOM003'=fund_company and left(convert(varchar, Send_result_date, 112),6) in (LEFT(@reportDate_n,6),@reportDate_b) AND convert(varchar, Send_result_date, 112) <= @reportDate_n  AND Send_result_type = 'SRT002' AND isnull(check_amount_type,'') NOT IN ('CKAT003') AND isnull(get_amount_type,'') NOT IN ('GTAT002','GTAT003') then pass_amount else 0 end ) as month_pass_amount_FDCOM003";
                //和潤 已核未撥
                T_SQL += " ,sum(case when 'FDCOM004'=fund_company and left(convert(varchar, Send_result_date, 112),6) in (LEFT(@reportDate_n,6),@reportDate_b) AND convert(varchar, Send_result_date, 112) <= @reportDate_n  AND Send_result_type = 'SRT002' AND isnull(check_amount_type,'') NOT IN ('CKAT003') AND isnull(get_amount_type,'') NOT IN ('GTAT002','GTAT003') then pass_amount else 0 end ) as month_pass_amount_FDCOM004";
                //福斯 已核未撥
                T_SQL += " ,sum(case when 'FDCOM005'=fund_company and left(convert(varchar, Send_result_date, 112),6) in (LEFT(@reportDate_n,6),@reportDate_b) AND convert(varchar, Send_result_date, 112) <= @reportDate_n  AND Send_result_type = 'SRT002' AND isnull(check_amount_type,'') NOT IN ('CKAT003') AND isnull(get_amount_type,'') NOT IN ('GTAT002','GTAT003') then pass_amount else 0 end ) as month_pass_amount_FDCOM005";
                //國&#23791; 代墊款(萬) 
                T_SQL += " ,sum(case when 'FDCOM003'=fund_company and left(convert(varchar, get_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, get_amount_date, 112) <= @reportDate_n  then advance_payment_AE else 0 end ) as advance_payment_AE";
                T_SQL += " FROM view_User_group ug join view_User_sales leader on leader.U_num = ug.group_M_code AND leader.U_BC = @U_BC ";
                T_SQL += " join view_User_sales sa on sa.U_num = ug.group_D_code ";
                T_SQL += " left join viewFeats f on ug.group_D_code = f.plan_num AND ( left(convert(varchar, f.Send_amount_date, 112),6) = LEFT(@reportDate_n,6) OR left(convert(varchar, f.Send_result_date, 112),6) in (LEFT(@reportDate_n,6),@reportDate_b) OR left(convert(varchar, f.get_amount_date, 112),6) = LEFT(@reportDate_n,6) ) ";
                T_SQL += " where 1=1 and @reportDate_n between ug.group_M_start_day and ug.group_M_end_day and @reportDate_n between ug.group_D_start_day and ug.group_D_end_day  group by isnull(ug.group_M_name,'未分組'),ug.group_D_name,ug.group_D_code, ug.group_id, ug.group_M_id,ug.group_M_title,sa.U_PFT_sort,sa.U_PFT_name)";
                T_SQL += " select SUM(a.day_incase_num_FDCOM001) as day_incase_num_FDCOM001_total,SUM(a.month_incase_num_FDCOM001) as month_incase_num_FDCOM001_total";
                T_SQL += " ,SUM(a.day_incase_num_FDCOM003) as day_incase_num_FDCOM003_total,SUM(a.month_incase_num_FDCOM003) as month_incase_num_FDCOM003_total";
                T_SQL += " ,SUM(a.day_get_amount_num) as day_get_amount_num_total,SUM(a.day_get_amount) as day_get_amount_total";
                T_SQL += " ,SUM(a.month_pass_num) as month_pass_num_total,SUM(a.month_get_amount_num) as month_get_amount_num_total";
                T_SQL += " ,SUM(a.month_get_amount_FDCOM001) as month_get_amount_FDCOM001_total,SUM(a.month_get_amount_FDCOM003) as month_get_amount_FDCOM003_total";
                T_SQL += " ,SUM(a.month_pass_amount_FDCOM001) as month_pass_amount_FDCOM001_total,SUM(a.month_pre_amount_FDCOM001) as month_pre_amount_FDCOM001_total";
                T_SQL += " ,SUM(a.month_pass_amount_FDCOM003) as month_pass_amount_FDCOM003_total,SUM(a.advance_payment_AE) as advance_payment_AE_total";
                T_SQL += " ,SUM(ft.target_quota) as target_quota_total";
                T_SQL += " ,CONVERT(DECIMAL(10, 2),CAST((SUM(a.month_get_amount_FDCOM001) + SUM(a.month_get_amount_FDCOM003)) AS FLOAT) / SUM(ft.target_quota) * 100) as percentage";
                T_SQL += " FROM A  left join Feat_target ft on ft.del_tag='0'  and ft.U_num=A.plan_num  and ft.group_id=A.group_id  and ft.target_ym=LEFT(@reportDate_n,6)";
                parameters.Add(new SqlParameter("@reportDate_n", model.reportDate_n));
                parameters.Add(new SqlParameter("@U_BC", model.U_BC));
                string reportDate_b = DateTime.ParseExact(model.reportDate_n, "yyyyMMdd", null).AddMonths(-1).ToString("yyyyMM");
                parameters.Add(new SqlParameter("@reportDate_b", reportDate_b));
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
        /// 日報表_(202106版)_下載 Feat_daily_report_v2021_Excel/Feat_daily_report_v2021.asp
        /// </summary>
        [HttpPost("Feat_Daily_Report_v2021_Excel")]
        public IActionResult Feat_Daily_Report_v2021_Excel(string reportDate_n)
        {
            try
            {
                FeatDailyReport_excel_Total modelFooter = new FeatDailyReport_excel_Total();
                ADOData _adoData = new ADOData();
                #region SQL_抓取所有分公司資料
                var parameters_comp = new List<SqlParameter>();
                var T_SQL_COMP = @"
                    select item_D_code,item_D_name,COUNT(*) as peoplecount from view_User_group ug
                    join view_User_sales leader on leader.U_num = ug.group_M_code
                    left join Item_list on Item_list.item_D_code=leader.U_BC and item_M_code='branch_company' and item_D_type='Y'
                    group by item_D_code,item_D_name,item_sort order by item_sort";
                #endregion
                var dtResultComp = _adoData.ExecuteQuery(T_SQL_COMP, parameters_comp);
                if (dtResultComp.Rows.Count > 0)
                {
                    byte[] fileBytes = null;
                    for (int i = 0; i < dtResultComp.Rows.Count; i++)
                    {
                        string itemDCode = dtResultComp.Rows[i]["item_D_code"].ToString();
                        string itemDName = dtResultComp.Rows[i]["item_D_name"].ToString();
                        int itemCount = (int)dtResultComp.Rows[i]["peoplecount"];

                        #region SQL
                        var parameters = new List<SqlParameter>();
                        var T_SQL = @"
                            with A ( leader_name,plan_name,plan_num,group_id,group_M_id,group_M_title ,U_PFT_sort,U_PFT_name ,day_incase_num_FDCOM001,month_incase_num_FDCOM001
                            ,day_incase_num_FDCOM003,month_incase_num_FDCOM003 ,day_incase_num_FDCOM004,month_incase_num_FDCOM004 ,day_incase_num_FDCOM005,month_incase_num_FDCOM005
                            ,day_get_amount_num,day_get_amount,month_pass_num,month_get_amount_num ,month_get_amount_FDCOM001,month_get_amount_FDCOM003,month_get_amount_FDCOM004,month_get_amount_FDCOM005,month_pass_amount_FDCOM001,month_pre_amount_FDCOM001,month_pass_amount_FDCOM003,month_pass_amount_FDCOM004,month_pass_amount_FDCOM005,advance_payment_AE ) as ( SELECT isnull(ug.group_M_name,'未分組') leader_name
                            ,ug.group_D_name,ug.group_D_code, ug.group_id, ug.group_M_id,ug.group_M_title ,sa.U_PFT_sort,sa.U_PFT_name";
                        //新鑫 日進件數 
                        T_SQL += " ,sum(case when 'FDCOM001'=fund_company and convert(varchar, Send_amount_date, 112) = @reportDate_n then 1 else 0 end ) as day_incase_num_FDCOM001";
                        //新鑫 月進件數
                        T_SQL += " ,sum(case when 'FDCOM001'=fund_company and left(convert(varchar, Send_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, Send_amount_date, 112) <= @reportDate_n then 1 else 0 end ) as month_incase_num_FDCOM001";
                        //國&#23791; 日進件數
                        T_SQL += " ,sum(case when 'FDCOM003'=fund_company and convert(varchar, Send_amount_date, 112) = @reportDate_n then 1 else 0 end ) as day_incase_num_FDCOM003";
                        //國&#23791; 月進件數
                        T_SQL += " ,sum(case when 'FDCOM003'=fund_company and left(convert(varchar, Send_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, Send_amount_date, 112) <= @reportDate_n then 1 else 0 end ) as month_incase_num_FDCOM003";
                        //和潤 日進件數
                        T_SQL += " ,sum(case when 'FDCOM004'=fund_company and convert(varchar, Send_amount_date, 112) = @reportDate_n then 1 else 0 end ) as day_incase_num_FDCOM004";
                        //和潤 月進件數
                        T_SQL += " ,sum(case when 'FDCOM004'=fund_company and left(convert(varchar, Send_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, Send_amount_date, 112) <= @reportDate_n then 1 else 0 end ) as month_incase_num_FDCOM004";
                        //福斯 日進件數
                        T_SQL += " ,sum(case when 'FDCOM005'=fund_company and convert(varchar, Send_amount_date, 112) = @reportDate_n then 1 else 0 end ) as day_incase_num_FDCOM005";
                        //福斯 月進件數
                        T_SQL += " ,sum(case when 'FDCOM005'=fund_company and left(convert(varchar, Send_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, Send_amount_date, 112) <= @reportDate_n then 1 else 0 end ) as month_incase_num_FDCOM005";
                        //日撥款數
                        T_SQL += " ,sum(case when convert(varchar, get_amount_date, 112) = @reportDate_n  then 1 else 0 end ) as day_get_amount_num";
                        //日撥款額
                        T_SQL += " ,sum(case when convert(varchar, get_amount_date, 112) = @reportDate_n  then get_amount else 0 end ) as day_get_amount";
                        //月核准數 
                        T_SQL += " ,sum(case when left(convert(varchar, Send_result_date, 112),6)=LEFT(@reportDate_n,6) AND convert(varchar, Send_result_date, 112) <= @reportDate_n  AND Send_result_type in ('SRT002','SRT005') then 1 else 0 end ) as month_pass_num";
                        //月撥款數
                        T_SQL += " ,sum(case when left(convert(varchar, get_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, get_amount_date, 112) <= @reportDate_n  then 1 else 0 end ) as month_get_amount_num";
                        //新鑫 月撥款額
                        T_SQL += " ,sum(case when 'FDCOM001'=fund_company and left(convert(varchar, get_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, get_amount_date, 112) <= @reportDate_n  then get_amount else 0 end ) as month_get_amount_FDCOM001";
                        //國&#23791; 月撥款額
                        T_SQL += " ,sum(case when 'FDCOM003'=fund_company and left(convert(varchar, get_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, get_amount_date, 112) <= @reportDate_n  then get_amount else 0 end ) as month_get_amount_FDCOM003";
                        //和潤 月撥款額
                        T_SQL += " ,sum(case when 'FDCOM004'=fund_company and left(convert(varchar, get_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, get_amount_date, 112) <= @reportDate_n  then get_amount else 0 end ) as month_get_amount_FDCOM004";
                        //福斯 月撥款額
                        T_SQL += " ,sum(case when 'FDCOM005'=fund_company and left(convert(varchar, get_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, get_amount_date, 112) <= @reportDate_n  then get_amount else 0 end ) as month_get_amount_FDCOM005";
                        //新鑫 已核未撥
                        T_SQL += " ,sum(case when 'FDCOM001'=fund_company and left(convert(varchar, Send_result_date, 112),6) in (LEFT(@reportDate_n,6),@reportDate_b) AND convert(varchar, Send_result_date, 112) <= @reportDate_n  AND Send_result_type = 'SRT002' AND isnull(check_amount_type,'') NOT IN ('CKAT003') AND isnull(get_amount_type,'') NOT IN ('GTAT002','GTAT003') then pass_amount else 0 end ) as month_pass_amount_FDCOM001";
                        //新鑫 預核額度
                        T_SQL += " ,sum(case when 'FDCOM001'=fund_company and left(convert(varchar, Send_result_date, 112),6) in (LEFT(@reportDate_n,6),@reportDate_b) AND convert(varchar, Send_result_date, 112) <= @reportDate_n  AND Send_result_type = 'SRT005' then pass_amount else 0 end ) as month_pre_amount_FDCOM001";
                        //國&#23791; 已核未撥
                        T_SQL += " ,sum(case when 'FDCOM003'=fund_company and left(convert(varchar, Send_result_date, 112),6) in (LEFT(@reportDate_n,6),@reportDate_b) AND convert(varchar, Send_result_date, 112) <= @reportDate_n  AND Send_result_type = 'SRT002' AND isnull(check_amount_type,'') NOT IN ('CKAT003') AND isnull(get_amount_type,'') NOT IN ('GTAT002','GTAT003') then pass_amount else 0 end ) as month_pass_amount_FDCOM003";
                        //和潤 已核未撥
                        T_SQL += " ,sum(case when 'FDCOM004'=fund_company and left(convert(varchar, Send_result_date, 112),6) in (LEFT(@reportDate_n,6),@reportDate_b) AND convert(varchar, Send_result_date, 112) <= @reportDate_n  AND Send_result_type = 'SRT002' AND isnull(check_amount_type,'') NOT IN ('CKAT003') AND isnull(get_amount_type,'') NOT IN ('GTAT002','GTAT003') then pass_amount else 0 end ) as month_pass_amount_FDCOM004";
                        //福斯 已核未撥
                        T_SQL += " ,sum(case when 'FDCOM005'=fund_company and left(convert(varchar, Send_result_date, 112),6) in (LEFT(@reportDate_n,6),@reportDate_b) AND convert(varchar, Send_result_date, 112) <= @reportDate_n  AND Send_result_type = 'SRT002' AND isnull(check_amount_type,'') NOT IN ('CKAT003') AND isnull(get_amount_type,'') NOT IN ('GTAT002','GTAT003') then pass_amount else 0 end ) as month_pass_amount_FDCOM005";
                        //國&#23791; 代墊款(萬) 
                        T_SQL += " ,sum(case when 'FDCOM003'=fund_company and left(convert(varchar, get_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, get_amount_date, 112) <= @reportDate_n  then advance_payment_AE else 0 end ) as advance_payment_AE";
                        T_SQL += " FROM view_User_group ug join view_User_sales leader on leader.U_num = ug.group_M_code AND leader.U_BC = @U_BC ";
                        T_SQL += " join view_User_sales sa on sa.U_num = ug.group_D_code ";
                        T_SQL += " left join viewFeats f on ug.group_D_code = f.plan_num AND ( left(convert(varchar, f.Send_amount_date, 112),6) = LEFT(@reportDate_n,6) OR left(convert(varchar, f.Send_result_date, 112),6) in (LEFT(@reportDate_n,6),@reportDate_b) OR left(convert(varchar, f.get_amount_date, 112),6) = LEFT(@reportDate_n,6) ) ";
                        T_SQL += " where 1=1 and @reportDate_n between ug.group_M_start_day and ug.group_M_end_day and @reportDate_n between ug.group_D_start_day and ug.group_D_end_day  group by isnull(ug.group_M_name,'未分組'),ug.group_D_name,ug.group_D_code, ug.group_id, ug.group_M_id,ug.group_M_title,sa.U_PFT_sort,sa.U_PFT_name)";
                        T_SQL += " select SUM(a.day_incase_num_FDCOM001) as day_incase_num_FDCOM001_total,SUM(a.month_incase_num_FDCOM001) as month_incase_num_FDCOM001_total";
                        T_SQL += " ,SUM(a.day_incase_num_FDCOM003) as day_incase_num_FDCOM003_total,SUM(a.month_incase_num_FDCOM003) as month_incase_num_FDCOM003_total";
                        T_SQL += " ,SUM(a.day_get_amount_num) as day_get_amount_num_total,SUM(a.day_get_amount) as day_get_amount_total";
                        T_SQL += " ,SUM(a.month_pass_num) as month_pass_num_total,SUM(a.month_get_amount_num) as month_get_amount_num_total";
                        T_SQL += " ,SUM(a.month_get_amount_FDCOM001) as month_get_amount_FDCOM001_total,SUM(a.month_get_amount_FDCOM003) as month_get_amount_FDCOM003_total";
                        T_SQL += " ,SUM(a.month_pass_amount_FDCOM001) as month_pass_amount_FDCOM001_total,SUM(a.month_pre_amount_FDCOM001) as month_pre_amount_FDCOM001_total";
                        T_SQL += " ,SUM(a.month_pass_amount_FDCOM003) as month_pass_amount_FDCOM003_total,SUM(a.advance_payment_AE) as advance_payment_AE_total";
                        T_SQL += " ,ISNULL(SUM(ft.target_quota),0) as target_quota_total";
                        T_SQL += " ,ISNULL(CONVERT(DECIMAL(10, 2),CAST((SUM(a.month_get_amount_FDCOM001) + SUM(a.month_get_amount_FDCOM003)) AS FLOAT) / SUM(ft.target_quota) * 100),0) as percentage";
                        T_SQL += " FROM A  left join Feat_target ft on ft.del_tag='0'  and ft.U_num=A.plan_num  and ft.group_id=A.group_id  and ft.target_ym=LEFT(@reportDate_n,6)";
                        parameters.Add(new SqlParameter("@reportDate_n", reportDate_n));
                        parameters.Add(new SqlParameter("@U_BC", itemDCode));
                        string reportDate_b = DateTime.ParseExact(reportDate_n, "yyyyMMdd", null).AddMonths(-1).ToString("yyyyMM");
                        parameters.Add(new SqlParameter("@reportDate_b", reportDate_b));
                        #endregion
                        var dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
                        if (dtResult.Rows.Count > 0)
                        {
                            var excelList = dtResult.AsEnumerable().Select(row => new FeatDailyReport_excel_Total
                            {
                                day_incase_num_FDCOM001_total = row.Field<int>("day_incase_num_FDCOM001_total"),
                                month_incase_num_FDCOM001_total = row.Field<int>("month_incase_num_FDCOM001_total"),
                                day_incase_num_FDCOM003_total = row.Field<int>("day_incase_num_FDCOM003_total"),
                                month_incase_num_FDCOM003_total = row.Field<int>("month_incase_num_FDCOM003_total"),
                                day_get_amount_num_total = row.Field<int>("day_get_amount_num_total"),
                                day_get_amount_total = row.Field<int>("day_get_amount_total"),
                                month_pass_num_total = row.Field<int>("month_pass_num_total"),
                                month_get_amount_num_total = row.Field<int>("month_get_amount_num_total"),
                                month_get_amount_FDCOM001_total = row.Field<int>("month_get_amount_FDCOM001_total"),
                                month_get_amount_FDCOM003_total = row.Field<int>("month_get_amount_FDCOM003_total"),
                                month_pass_amount_FDCOM001_total = row.Field<int>("month_pass_amount_FDCOM001_total"),
                                month_pre_amount_FDCOM001_total = row.Field<int>("month_pre_amount_FDCOM001_total"),
                                month_pass_amount_FDCOM003_total = row.Field<int>("month_pass_amount_FDCOM003_total"),
                                advance_payment_AE_total = row.Field<decimal>("advance_payment_AE_total"),
                                target_quota_total = row.Field<int>("target_quota_total")
                            }).ToList();
                            var Excel_Headers = new Dictionary<string, string>
                            {
                                { "U_PFT_name", "職位" },
                                { "plan_name", "業務" },
                                { "day_incase_num_FDCOM001_total", "新鑫日進件數" },
                                { "month_incase_num_FDCOM001_total", "新鑫累積進件數" },
                                { "day_incase_num_FDCOM003_total", "國峯日進件數" },
                                { "month_incase_num_FDCOM003_total", "國峯累積進件數" },
                                { "day_get_amount_num_total", "日撥件數" },
                                { "day_get_amount_total", "日撥金額" },
                                { "month_pass_num_total", "核准件數" },
                                { "month_get_amount_num_total", "累積撥款件數" },
                                { "month_get_amount_FDCOM001_total", "新鑫-撥款金額(萬)" },
                                { "month_get_amount_FDCOM003_total", "國峯-撥款金額(萬)" },
                                { "month_pass_amount_FDCOM001_total", "新鑫已核未撥" },
                                { "month_pre_amount_FDCOM001_total", "新鑫預核額度" },
                                { "month_pass_amount_FDCOM003_total", "國峯已核未撥" },
                                { "advance_payment_AE_total", "國峯-代墊款(萬)" },
                                { "target_quota_total", "目標(萬)" },
                                { "AchievementRate", "達成率" }
                            };
                            if (i == 0)
                            {
                                fileBytes = FuncHandler.FeatDailyToExcel(excelList, Excel_Headers, itemDName, reportDate_n, itemCount);
                            }
                            else
                            {
                                fileBytes = FuncHandler.FeatDailyToExcelAgain(fileBytes, excelList, Excel_Headers, itemDName, reportDate_n, itemCount);
                            }
                            #region 總計 FeatDailyToExcelFooter
                            var total_D_FDCOM001 = excelList.Sum(item => item.day_incase_num_FDCOM001_total);
                            modelFooter.day_incase_num_FDCOM001_total += total_D_FDCOM001;
                            var total_M_FDCOM001 = excelList.Sum(item => item.month_incase_num_FDCOM001_total);
                            modelFooter.month_incase_num_FDCOM001_total += total_M_FDCOM001;
                            var total_D_FDCOM003 = excelList.Sum(item => item.day_incase_num_FDCOM003_total);
                            modelFooter.day_incase_num_FDCOM003_total += total_D_FDCOM003;
                            var total_M_FDCOM003 = excelList.Sum(item => item.month_incase_num_FDCOM003_total);
                            modelFooter.month_incase_num_FDCOM003_total += total_M_FDCOM003;
                            var total_D_AmoutNum = excelList.Sum(item => item.day_get_amount_num_total);
                            modelFooter.day_get_amount_num_total += total_D_AmoutNum;
                            var total_S_Amout = excelList.Sum(item => item.day_get_amount_total);
                            modelFooter.day_get_amount_total += total_S_Amout;
                            var totol_M_PassNum = excelList.Sum(item => item.month_pass_num_total);
                            modelFooter.month_pass_num_total += totol_M_PassNum;
                            var total_M_AmoutNum = excelList.Sum(item => item.month_get_amount_num_total);
                            modelFooter.month_get_amount_num_total += total_M_AmoutNum;
                            var total_M_Amout_FDCOM001 = excelList.Sum(item => item.month_get_amount_FDCOM001_total);
                            modelFooter.month_get_amount_FDCOM001_total += total_M_Amout_FDCOM001;
                            var total_M_Amout_FDCOM003 = excelList.Sum(item => item.month_get_amount_FDCOM003_total);
                            modelFooter.month_get_amount_FDCOM003_total += total_M_Amout_FDCOM003;
                            var total_M_PassAmout_FDCOM001 = excelList.Sum(item => item.month_pass_amount_FDCOM001_total);
                            modelFooter.month_pass_amount_FDCOM001_total += total_M_PassAmout_FDCOM001;
                            var total_M_PreAmout_FDCOM001 = excelList.Sum(item => item.month_pre_amount_FDCOM001_total);
                            modelFooter.month_pre_amount_FDCOM001_total += total_M_PreAmout_FDCOM001;
                            var total_M_PassAmout_FDCOM003 = excelList.Sum(item => item.month_pass_amount_FDCOM003_total);
                            modelFooter.month_pass_amount_FDCOM003_total += total_M_PassAmout_FDCOM003;
                            var total_Pay = excelList.Sum(item => item.advance_payment_AE_total);
                            modelFooter.advance_payment_AE_total += total_Pay;
                            var total_Quota = excelList.Sum(item => item.target_quota_total);
                            modelFooter.target_quota_total += total_Quota;
                            #endregion
                        }
                    }
                    //全區總計 FeatDailyToExcelFooter
                    var Excel_Footer = new Dictionary<string, string>
                    {
                        { "U_PFT_name", "職位" },
                        { "plan_name", "業務" },
                        { "day_incase_num_FDCOM001_total", "新鑫日進件數" },
                        { "month_incase_num_FDCOM001_total", "新鑫累積進件數" },
                        { "day_incase_num_FDCOM003_total", "國峯日進件數" },
                        { "month_incase_num_FDCOM003_total", "國峯累積進件數" },
                        { "day_get_amount_num_total", "日撥件數" },
                        { "day_get_amount_total", "日撥金額" },
                        { "month_pass_num_total", "核准件數" },
                        { "month_get_amount_num_total", "累積撥款件數" },
                        { "month_get_amount_FDCOM001_total", "新鑫-撥款金額(萬)" },
                        { "month_get_amount_FDCOM003_total", "國峯-撥款金額(萬)" },
                        { "month_pass_amount_FDCOM001_total", "新鑫已核未撥" },
                        { "month_pre_amount_FDCOM001_total", "新鑫預核額度" },
                        { "month_pass_amount_FDCOM003_total", "國峯已核未撥" },
                        { "advance_payment_AE_total", "國峯-代墊款(萬)" },
                        { "target_quota_total", "目標(萬)" },
                        { "AchievementRate", "達成率" }
                    };
                    fileBytes = FuncHandler.FeatDailyToExcelFooter(fileBytes, modelFooter, Excel_Footer, reportDate_n);
                    var fileName = "日報表_合計版" + DateTime.Now.ToString("yyyyMMddHHmm") + ".xlsx";
                    return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
                }
                else
                {
                    return NotFound(); // 檔案不存在時返回 404
                }
            }
            catch (Exception ex)
            {
                ResultClass<string> resultClass = new ResultClass<string>();
                resultClass.ResultCode = "500";
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass);
            }
        }
        #endregion

        #region 案件狀態表
        /// <summary>
        /// 案件狀態表_查詢 SendCaseStatus_Query/SendCaseStatus.asp
        /// </summary>
        [HttpPost("SendCaseStatus_Query")]
        public ActionResult<ResultClass<string>> SendCaseStatus_Query(SendCaseStatu_req model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = "select * from [dbo].[fun_SendCaseStatus]( @Send_Date_S,@Send_Date_E )";
                if (!string.IsNullOrEmpty(model.Company))
                {
                    T_SQL += " where fund_Company=@Company";
                    parameters.Add(new SqlParameter("@Company", model.Company));
                }
                parameters.Add(new SqlParameter("@Send_Date_S", model.Send_Date_S));
                parameters.Add(new SqlParameter("@Send_Date_E", model.Send_Date_E));
                #endregion
                DataTable dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
                if (dtResult.Rows.Count > 0)
                {
                    var ExcelList = dtResult.AsEnumerable().Select(row => new SendCaseStatu_res
                    {
                        fund_company = row.Field<string>("fund_company"),
                        Company_Name = row.Field<string>("Company_Name"),
                        totleCount = row.Field<int>("totleCount"),
                        ApprCount = row.Field<int>("ApprCount"),
                        unApprCount = row.Field<int>("unApprCount"),
                        PayCount = row.Field<int>("PayCount"),
                        unPayCount = row.Field<int>("unPayCount"),
                        WPayCount = row.Field<int>("WPayCount"),
                        GUCount = row.Field<int>("GUCount")
                    }).ToList();
                    foreach (var item in ExcelList)
                    {
                        //送審中跟婉拒的筆數
                        #region SQL
                        var parameters_q = new List<SqlParameter>();
                        var T_SQL_Q = @"
                            SELECT item_D_code,COUNT(fund_company) AS unApprCount FROM House_sendcase H
                            LEFT JOIN (SELECT item_D_code,item_D_name FROM Item_list WHERE item_M_code = 'Send_result_type' AND item_D_type = 'Y') ST
                            ON H.Send_result_type = ST.item_D_code
                            WHERE Send_amount_date >= @Send_Date_S AND Send_amount_date <= @Send_Date_E
                            AND Send_result_type <> 'SRT002' AND sendcase_handle_type = 'Y' AND fund_company = @Company
                            GROUP BY item_D_code";
                        parameters_q.Add(new SqlParameter("@Send_Date_S", model.Send_Date_S));
                        parameters_q.Add(new SqlParameter("@Send_Date_E", model.Send_Date_E));
                        parameters_q.Add(new SqlParameter("@Company", item.fund_company));
                        #endregion
                        DataTable dt = _adoData.ExecuteQuery(T_SQL_Q, parameters_q);
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            if (dt.Rows[i]["item_D_code"].ToString() == "SRT001")
                            {
                                item.Review_count = (int)dt.Rows[i]["unApprCount"];
                            }
                            else if (dt.Rows[i]["item_D_code"].ToString() == "SRT004")
                            {
                                item.Decline_count = (int)dt.Rows[i]["unApprCount"];
                            }
                        }
                        //不對保跟待對保筆數
                        #region SQL
                        var parameters_gu = new List<SqlParameter>();
                        var T_SQL_GU = @"
                            select item_D_code,count(fund_company) as GUDCount from House_sendcase H
                            left join (select item_D_code,item_D_name  from Item_list
                            where item_M_code = 'check_amount_type' AND item_D_type='Y') ST on H.check_amount_type=ST.item_D_code
                            where (Send_amount_date >= @Send_Date_S AND Send_amount_date <= @Send_Date_E ) and sendcase_handle_type='Y'
                            AND Send_result_type= 'SRT002' and HS_id not in ( select  H.HS_id from House_sendcase H
                            where (Send_amount_date >= @Send_Date_S AND Send_amount_date <= @Send_Date_E ) and sendcase_handle_type='Y'
                            AND isnull(H.Send_amount,'')<>'' AND get_amount_type in ( 'GTAT001','GTAT002','GTAT003')
                            ) AND  fund_company= @Company group by item_D_name,item_D_code  order by item_D_name";
                        parameters_gu.Add(new SqlParameter("@Send_Date_S", model.Send_Date_S));
                        parameters_gu.Add(new SqlParameter("@Send_Date_E", model.Send_Date_E));
                        parameters_gu.Add(new SqlParameter("@Company", item.fund_company));
                        #endregion
                        DataTable dt_gu = _adoData.ExecuteQuery(T_SQL_GU, parameters_gu);
                        for (int i = 0; i < dt_gu.Rows.Count; i++)
                        {
                            if (dt_gu.Rows[i]["item_D_code"].ToString() == "CKAT003")
                            {
                                item.GuaranteeNone = (int)dt_gu.Rows[i]["GUDCount"];
                            }
                            else if (dt_gu.Rows[i]["item_D_code"].ToString() == "CKAT001")
                            {
                                item.Guarantee = (int)dt_gu.Rows[i]["GUDCount"];
                            }
                        }
                    }
                    resultClass.ResultCode = "000";
                    resultClass.objResult = JsonConvert.SerializeObject(ExcelList);
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
        /// 案件狀態表_明細 SendCaseStatus_Open_Detail/SendCaseStatus_Detl.asp
        /// </summary>
        [HttpPost("SendCaseStatus_Open_Detail")]
        public ActionResult<ResultClass<string>> SendCaseStatus_Open_Detail(SendCaseStatu_req model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"
                    SELECT ISNULL(CS_PID, '') AS CS_ID,ISNULL(CS_register_address, '') AS addr,ISNULL(Loan_rate, '') AS Loan_rate
                    ,REPLACE(REPLACE(CONVERT(varchar(10), Send_amount_date, 126), '-', '/'),YEAR(Send_amount_date)
                    ,YEAR(Send_amount_date) - 1911) AS Send_amount_date,H.HS_id,H.pass_amount,H.Send_amount
                    ,House_pre_project.project_title,House_apply.CS_name,House_apply.CS_MTEL1,House_apply.CS_MTEL2
                    ,House_apply.CS_introducer,User_M.U_name AS plan_name,User_M.U_BC_name
                    ,ISNULL((SELECT item_D_name FROM Item_list WHERE item_M_code = 'Send_result_type'
                    AND item_D_type = 'Y' AND item_D_code = H.Send_result_type AND show_tag = '0'
                    AND del_tag = '0'), '--') AS show_Send_result_type
                    ,ISNULL((SELECT item_D_name FROM Item_list WHERE item_M_code = 'check_amount_type'
                    AND item_D_type = 'Y' AND item_D_code = H.check_amount_type
                    AND show_tag = '0'AND del_tag = '0'), '--') AS show_check_amount_type
                    ,ISNULL((SELECT item_D_name FROM Item_list WHERE item_M_code = 'get_amount_type'
                    AND item_D_type = 'Y' AND item_D_code = H.get_amount_type
                    AND show_tag = '0' AND del_tag = '0'), '--') AS show_get_amount_type,
                    (SELECT item_D_name FROM Item_list WHERE item_M_code = 'appraise_company' AND item_D_type = 'Y'
                    AND item_D_code = H.appraise_company AND show_tag = '0' AND del_tag = '0') AS show_appraise_company,
                    (SELECT item_D_name FROM Item_list WHERE item_M_code = 'fund_company'
                    AND item_D_type = 'Y' AND item_D_code = H.fund_company AND show_tag = '0' AND del_tag = '0') AS show_fund_company,
                    (SELECT item_D_name FROM Item_list WHERE item_M_code = 'project_title' AND item_D_type = 'Y'
                    AND item_D_code = House_pre_project.project_title AND show_tag = '0'
                    AND del_tag = '0') AS show_project_title,Users.U_name AS fin_name,
                    CONVERT(varchar, House_pre_project.fin_date, 111) AS fin_date FROM House_sendcase H
                    LEFT JOIN House_apply ON House_apply.HA_id = H.HA_id AND House_apply.del_tag = '0'
                    LEFT JOIN House_pre_project ON House_pre_project.HP_project_id = H.HP_project_id AND House_pre_project.del_tag = '0'
                    LEFT JOIN view_User_sales User_M ON User_M.U_num = House_apply.plan_num
                    LEFT JOIN view_user_sales Users ON house_pre_project.fin_user = Users.U_num
                    WHERE H.del_tag = '0' AND H.sendcase_handle_type = 'Y' AND ISNULL(H.Send_amount, '') <> ''
                    AND (Send_amount_date >= @Send_Date_S AND Send_amount_date <= @Send_Date_E)
                    AND fund_Company=@Company";
                switch (model.status)
                {
                    case "Appr":
                        T_SQL += " AND Send_result_type= 'SRT002'";
                        break;
                    case "unAppr":
                        T_SQL += " AND Send_result_type<> 'SRT002'";
                        break;
                    case "Pay":
                        T_SQL += " AND get_amount_type= 'GTAT002'";
                        break;
                    case "unPay":
                        T_SQL += " AND get_amount_type = 'GTAT003'";
                        break;
                    case "WPay":
                        T_SQL += " AND get_amount_type = 'GTAT001'";
                        break;
                    case "GU":
                        T_SQL += " AND Send_result_type= 'SRT002' and HS_id not in (";
                        T_SQL += " select  H.HS_id from House_sendcase H";
                        T_SQL += " where (Send_amount_date >= @Send_Date_S AND Send_amount_date <= @Send_Date_E";
                        T_SQL += " ) and sendcase_handle_type='Y'　AND isnull(H.Send_amount,'')<>''";
                        T_SQL += " AND get_amount_type in ( 'GTAT001','GTAT002','GTAT003'))";
                        break;
                    default:
                        break;
                }
                T_SQL += " ORDER BY Send_amount_date, H.HS_id DESC";
                parameters.Add(new SqlParameter("@Send_Date_S", model.Send_Date_S));
                parameters.Add(new SqlParameter("@Send_Date_E", model.Send_Date_E));
                parameters.Add(new SqlParameter("@Company", model.Company));
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
        /// 案件狀態表_下載 SendCaseStatus_Query_Excel/SendCaseStatus.asp
        /// </summary>
        [HttpPost("SendCaseStatus_Query_Excel")]
        public IActionResult SendCaseStatus_Query_Excel(SendCaseStatu_req model)
        {
            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = "select * from [dbo].[fun_SendCaseStatus]( @Send_Date_S,@Send_Date_E )";
                if (!string.IsNullOrEmpty(model.Company))
                {
                    T_SQL += " where fund_Company=@Company";
                    parameters.Add(new SqlParameter("@Company", model.Company));
                }
                parameters.Add(new SqlParameter("@Send_Date_S", model.Send_Date_S));
                parameters.Add(new SqlParameter("@Send_Date_E", model.Send_Date_E));
                #endregion
                DataTable dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
                if (dtResult.Rows.Count > 0)
                {
                    var ExcelList = dtResult.AsEnumerable().Select(row => new SendCaseStatu_Excel
                    {
                        Company_Name = row.Field<string>("Company_Name"),
                        totleCount = row.Field<int>("totleCount"),
                        ApprCount = row.Field<int>("ApprCount"),
                        unApprCount = row.Field<int>("unApprCount"),
                        PayCount = row.Field<int>("PayCount"),
                        unPayCount = row.Field<int>("unPayCount"),
                        WPayCount = row.Field<int>("WPayCount"),
                        GUCount = row.Field<int>("GUCount")
                    }).ToList();
                    var Excel_Headers = new Dictionary<string, string>
                    {
                        { "Company_Name", "放款公司" },
                        { "totleCount", "案件總數" },
                        { "unApprCount", "未核准" },
                        { "ApprCount", "已核准" },
                        { "PayCount", "已撥款" },
                        { "unPayCount", "不撥款" },
                        { "WPayCount", "待撥款" },
                        { "GUCount", "對保" }
                    };
                    var fileBytes = FuncHandler.ExportToExcel(ExcelList, Excel_Headers);
                    var fileName = "案件狀態表" + DateTime.Now.ToString("yyyyMMddHHmm") + ".xlsx";
                    return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
                }
                else
                {
                    return NotFound(); // 檔案不存在時返回 404
                }
            }
            catch (Exception ex)
            {
                ResultClass<string> resultClass = new ResultClass<string>();
                resultClass.ResultCode = "500";
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass);
            }
        }
        /// <summary>
        /// 案件狀態表_明細_下載 SendCaseStatus_Open_Detail_Excel/SendCaseStatus_Open_Detail_Excel
        /// </summary>
        [HttpPost("SendCaseStatus_Open_Detail_Excel")]
        public IActionResult SendCaseStatus_Open_Detail_Excel(SendCaseStatu_req model)
        {
            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"
                    SELECT ISNULL(CS_PID, '') AS CS_ID,ISNULL(CS_register_address, '') AS addr,ISNULL(Loan_rate, '') AS Loan_rate
                    ,REPLACE(REPLACE(CONVERT(varchar(10), Send_amount_date, 126), '-', '/'),YEAR(Send_amount_date)
                    ,YEAR(Send_amount_date) - 1911) AS Send_amount_date,H.HS_id,H.pass_amount,H.Send_amount
                    ,House_pre_project.project_title,House_apply.CS_name,House_apply.CS_MTEL1,House_apply.CS_MTEL2
                    ,House_apply.CS_introducer,User_M.U_name AS plan_name,User_M.U_BC_name
                    ,ISNULL((SELECT item_D_name FROM Item_list WHERE item_M_code = 'Send_result_type'
                    AND item_D_type = 'Y' AND item_D_code = H.Send_result_type AND show_tag = '0'
                    AND del_tag = '0'), '--') AS show_Send_result_type
                    ,ISNULL((SELECT item_D_name FROM Item_list WHERE item_M_code = 'check_amount_type'
                    AND item_D_type = 'Y' AND item_D_code = H.check_amount_type
                    AND show_tag = '0'AND del_tag = '0'), '--') AS show_check_amount_type
                    ,ISNULL((SELECT item_D_name FROM Item_list WHERE item_M_code = 'get_amount_type'
                    AND item_D_type = 'Y' AND item_D_code = H.get_amount_type
                    AND show_tag = '0' AND del_tag = '0'), '--') AS show_get_amount_type,
                    (SELECT item_D_name FROM Item_list WHERE item_M_code = 'appraise_company' AND item_D_type = 'Y'
                    AND item_D_code = H.appraise_company AND show_tag = '0' AND del_tag = '0') AS show_appraise_company,
                    (SELECT item_D_name FROM Item_list WHERE item_M_code = 'fund_company'
                    AND item_D_type = 'Y' AND item_D_code = H.fund_company AND show_tag = '0' AND del_tag = '0') AS show_fund_company,
                    (SELECT item_D_name FROM Item_list WHERE item_M_code = 'project_title' AND item_D_type = 'Y'
                    AND item_D_code = House_pre_project.project_title AND show_tag = '0'
                    AND del_tag = '0') AS show_project_title,Users.U_name AS fin_name,
                    CONVERT(varchar, House_pre_project.fin_date, 111) AS fin_date FROM House_sendcase H
                    LEFT JOIN House_apply ON House_apply.HA_id = H.HA_id AND House_apply.del_tag = '0'
                    LEFT JOIN House_pre_project ON House_pre_project.HP_project_id = H.HP_project_id AND House_pre_project.del_tag = '0'
                    LEFT JOIN view_User_sales User_M ON User_M.U_num = House_apply.plan_num
                    LEFT JOIN view_user_sales Users ON house_pre_project.fin_user = Users.U_num
                    WHERE H.del_tag = '0' AND H.sendcase_handle_type = 'Y' AND ISNULL(H.Send_amount, '') <> ''
                    AND (Send_amount_date >= @Send_Date_S AND Send_amount_date <= @Send_Date_E)
                    AND fund_Company=@Company";
                switch (model.status)
                {
                    case "Appr":
                        T_SQL += " AND Send_result_type= 'SRT002'";
                        break;
                    case "unAppr":
                        T_SQL += " AND Send_result_type<> 'SRT002'";
                        break;
                    case "Pay":
                        T_SQL += " AND get_amount_type= 'GTAT002'";
                        break;
                    case "unPay":
                        T_SQL += " AND get_amount_type = 'GTAT003'";
                        break;
                    case "WPay":
                        T_SQL += " AND get_amount_type = 'GTAT001'";
                        break;
                    case "GU":
                        T_SQL += " AND Send_result_type= 'SRT002' and HS_id not in (";
                        T_SQL += " select  H.HS_id from House_sendcase H";
                        T_SQL += " where (Send_amount_date >= @Send_Date_S AND Send_amount_date <= @Send_Date_E";
                        T_SQL += " ) and sendcase_handle_type='Y'　AND isnull(H.Send_amount,'')<>''";
                        T_SQL += " AND get_amount_type in ( 'GTAT001','GTAT002','GTAT003'))";
                        break;
                    default:
                        break;
                }
                T_SQL += " ORDER BY Send_amount_date, H.HS_id DESC";
                parameters.Add(new SqlParameter("@Send_Date_S", model.Send_Date_S));
                parameters.Add(new SqlParameter("@Send_Date_E", model.Send_Date_E));
                parameters.Add(new SqlParameter("@Company", model.Company));
                #endregion
                DataTable dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
                if (dtResult.Rows.Count > 0)
                {

                    List<SendCaseStatu_Det_Excel> ExcelList = new List<SendCaseStatu_Det_Excel>();
                    switch (model.status)
                    {
                        case "Appr":
                            ExcelList = dtResult.AsEnumerable().Select(row => new SendCaseStatu_Det_Excel
                            {
                                HS_id = row.Field<decimal>("HS_id"),
                                show_fund_company = row.Field<string>("show_fund_company"),
                                Send_amount_date = row.Field<string>("Send_amount_date"),
                                CS_name = row.Field<string>("CS_name"),
                                CS_ID = row.Field<string>("CS_ID"),
                                CS_MTEL1 = row.Field<string>("CS_MTEL1"),
                                show_appraise_company = row.Field<string>("show_appraise_company"),
                                show_project_title = row.Field<string>("show_project_title"),
                                ShowType = row.Field<string>("show_Send_result_type"),
                                Send_amount = row.Field<string>("Send_amount"),
                                pass_amount = row.Field<string>("pass_amount"),
                                Loan_rate = row.Field<string>("Loan_rate"),
                                addr = row.Field<string>("addr")
                            }).ToList();
                            break;
                        case "unAppr":
                            ExcelList = dtResult.AsEnumerable().Select(row => new SendCaseStatu_Det_Excel
                            {
                                HS_id = row.Field<decimal>("HS_id"),
                                show_fund_company = row.Field<string>("show_fund_company"),
                                Send_amount_date = row.Field<string>("Send_amount_date"),
                                CS_name = row.Field<string>("CS_name"),
                                CS_ID = row.Field<string>("CS_ID"),
                                CS_MTEL1 = row.Field<string>("CS_MTEL1"),
                                show_appraise_company = row.Field<string>("show_appraise_company"),
                                show_project_title = row.Field<string>("show_project_title"),
                                ShowType = row.Field<string>("show_Send_result_type"),
                                Send_amount = row.Field<string>("Send_amount"),
                                pass_amount = row.Field<string>("pass_amount"),
                                Loan_rate = row.Field<string>("Loan_rate"),
                                addr = row.Field<string>("addr")
                            }).ToList();
                            break;
                        case "Pay":
                            ExcelList = dtResult.AsEnumerable().Select(row => new SendCaseStatu_Det_Excel
                            {
                                HS_id = row.Field<decimal>("HS_id"),
                                show_fund_company = row.Field<string>("show_fund_company"),
                                Send_amount_date = row.Field<string>("Send_amount_date"),
                                CS_name = row.Field<string>("CS_name"),
                                CS_ID = row.Field<string>("CS_ID"),
                                CS_MTEL1 = row.Field<string>("CS_MTEL1"),
                                show_appraise_company = row.Field<string>("show_appraise_company"),
                                show_project_title = row.Field<string>("show_project_title"),
                                ShowType = row.Field<string>("show_get_amount_type"),
                                Send_amount = row.Field<string>("Send_amount"),
                                pass_amount = row.Field<string>("pass_amount"),
                                Loan_rate = row.Field<string>("Loan_rate"),
                                addr = row.Field<string>("addr")
                            }).ToList();
                            break;
                        case "unPay":
                            ExcelList = dtResult.AsEnumerable().Select(row => new SendCaseStatu_Det_Excel
                            {
                                HS_id = row.Field<decimal>("HS_id"),
                                show_fund_company = row.Field<string>("show_fund_company"),
                                Send_amount_date = row.Field<string>("Send_amount_date"),
                                CS_name = row.Field<string>("CS_name"),
                                CS_ID = row.Field<string>("CS_ID"),
                                CS_MTEL1 = row.Field<string>("CS_MTEL1"),
                                show_appraise_company = row.Field<string>("show_appraise_company"),
                                show_project_title = row.Field<string>("show_project_title"),
                                ShowType = row.Field<string>("show_get_amount_type"),
                                Send_amount = row.Field<string>("Send_amount"),
                                pass_amount = row.Field<string>("pass_amount"),
                                Loan_rate = row.Field<string>("Loan_rate"),
                                addr = row.Field<string>("addr")
                            }).ToList();
                            break;
                        case "WPay":
                            ExcelList = dtResult.AsEnumerable().Select(row => new SendCaseStatu_Det_Excel
                            {
                                HS_id = row.Field<decimal>("HS_id"),
                                show_fund_company = row.Field<string>("show_fund_company"),
                                Send_amount_date = row.Field<string>("Send_amount_date"),
                                CS_name = row.Field<string>("CS_name"),
                                CS_ID = row.Field<string>("CS_ID"),
                                CS_MTEL1 = row.Field<string>("CS_MTEL1"),
                                show_appraise_company = row.Field<string>("show_appraise_company"),
                                show_project_title = row.Field<string>("show_project_title"),
                                ShowType = row.Field<string>("show_get_amount_type"),
                                Send_amount = row.Field<string>("Send_amount"),
                                pass_amount = row.Field<string>("pass_amount"),
                                Loan_rate = row.Field<string>("Loan_rate"),
                                addr = row.Field<string>("addr")
                            }).ToList();
                            break;
                        case "GU":
                            ExcelList = dtResult.AsEnumerable().Select(row => new SendCaseStatu_Det_Excel
                            {
                                HS_id = row.Field<decimal>("HS_id"),
                                show_fund_company = row.Field<string>("show_fund_company"),
                                Send_amount_date = row.Field<string>("Send_amount_date"),
                                CS_name = row.Field<string>("CS_name"),
                                CS_ID = row.Field<string>("CS_ID"),
                                CS_MTEL1 = row.Field<string>("CS_MTEL1"),
                                show_appraise_company = row.Field<string>("show_appraise_company"),
                                show_project_title = row.Field<string>("show_project_title"),
                                ShowType = row.Field<string>("show_check_amount_type"),
                                Send_amount = row.Field<string>("Send_amount"),
                                pass_amount = row.Field<string>("pass_amount"),
                                Loan_rate = row.Field<string>("Loan_rate"),
                                addr = row.Field<string>("addr")
                            }).ToList();
                            break;
                        default:
                            break;
                    }
                    var Excel_Headers = new Dictionary<string, string>
                    {
                        { "HS_id", "申貸案號" },
                        { "show_fund_company", "放款公司" },
                        { "Send_amount_date", "出件日期" },
                        { "CS_name", "申請人" },
                        { "CS_ID", "申請人ID" },
                        { "CS_MTEL1", "行動一" },
                        { "show_appraise_company", "出件公司" },
                        { "show_project_title", "出件方案" },
                        { "ShowType", "狀態" },
                        { "Send_amount", "申貸金額(萬)" },
                        { "pass_amount", "核准金額(萬)" },
                        { "Loan_rate", "貸款成數" },
                        { "addr", "地址" }
                    };
                    var fileBytes = FuncHandler.ExportToExcel(ExcelList, Excel_Headers);
                    var fileName = "案件狀態明細表" + DateTime.Now.ToString("yyyyMMddHHmm") + ".xlsx";
                    return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
                }
                else
                {
                    return NotFound(); // 檔案不存在時返回 404
                }
            }
            catch (Exception ex)
            {
                ResultClass<string> resultClass = new ResultClass<string>();
                resultClass.ResultCode = "500";
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass);
            }
        }
        #endregion

        #region 年度業績還比表 / 部門年度業績還比表(多出需SESSION U_BC) / 業績表(放款公司)
        /// <summary>
        /// GetPerfByYYYY/_Ajaxhandler.asp
        /// </summary>
        /// <param name="this_YY">2024-09</param>
        /// <param name="last_YY">2023</param>
        /// <param name="fund_Comp">FDCOM003</param>
        [HttpGet("GetPerfByYYYY")]
        public ActionResult<ResultClass<string>> GetPerfByYYYY(string this_YY, string last_YY, string? fund_Comp)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"
                    select sum( CAST(House_sendcase.get_amount AS DECIMAL))amount, month( get_amount_date)  mm
                    ,Year(get_amount_date) yyyy from House_sendcase
                    LEFT JOIN House_apply on House_apply.HA_id = House_sendcase.HA_id where get_amount_type='GTAT002'
                    and House_sendcase.del_tag = '0'  AND House_apply.del_tag='0'
                    AND isnull(House_sendcase.get_amount,'') <>'' AND convert(varchar(7), (get_amount_date), 126)   between @last_YY and  @this_YY";
                if (!string.IsNullOrEmpty(fund_Comp))
                {
                    T_SQL += " and fund_company=@fund_Comp";
                    parameters.Add(new SqlParameter("@fund_Comp", fund_Comp));
                }
                T_SQL += " group by Year(get_amount_date), month( get_amount_date)";
                T_SQL += " order by Year(get_amount_date) asc, month( get_amount_date) asc";
                parameters.Add(new SqlParameter("@this_YY", this_YY));
                parameters.Add(new SqlParameter("@last_YY", last_YY + "-01"));
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
        /// GetBCPerfByYYYYMM/_Ajaxhandler.asp
        /// </summary>
        /// <param name="this_YYMM">2024-09</param>
        /// <param name="last_YY">2023</param>
        /// <param name="fund_Comp">FDCOM003</param>
        [HttpGet("GetBCPerfByYYYYMM")]
        public ActionResult<ResultClass<string>> GetBCPerfByYYYYMM(string this_YYMM, string last_YY, string? fund_company)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"
                    SELECT sum(CAST(H.get_amount AS DECIMAL))amount,Year(get_amount_date) yyyy,month(get_amount_date) mm,U_BC,BC_NA
                    FROM House_sendcase H
                    LEFT JOIN House_apply A ON A.HA_id = H.HA_id
                    LEFT JOIN (select U.U_num,U.U_BC,UB.item_D_name BC_NA from User_M u
                    LEFT JOIN Item_list ub ON ub.item_M_code='branch_company' AND ub.item_D_type='Y' AND ub.item_D_code=u.U_BC
                    ) U on A.plan_num=U.U_num
                    WHERE H.del_tag = '0' AND A.del_tag='0'  AND U_BC not in ('BC0700','BC0800')
                    AND isnull(H.get_amount, '') <>''
                    AND convert(varchar(7), (get_amount_date), 126) between @last_YYMM and @this_YYMM";
                if (!string.IsNullOrEmpty(fund_company))
                {
                    T_SQL += " and fund_company=@fund_company";
                    parameters.Add(new SqlParameter("@fund_company", fund_company));
                }
                T_SQL += " GROUP BY Year(get_amount_date),month(get_amount_date),U_BC,BC_NA";
                parameters.Add(new SqlParameter("@last_YYMM", last_YY + "-01"));
                parameters.Add(new SqlParameter("@this_YYMM", this_YYMM));
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
        /// GetTotRateByYYYY/_Ajaxhandler.asp
        /// </summary>
        /// <param name="U_BC">BC0100</param>
        /// <param name="YYYYMM">2024-09</param>
        [HttpGet("GetTotRateByYYYY")]
        public ActionResult<ResultClass<string>> GetTotRateByYYYY(string? U_BC, string YYYYMM)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"
                    select ThisAmtTot,LastAmtTot,CAST(ROUND((ThisAmtTot-LastAmtTot)/LastAmtTot *100,2) AS decimal(5,2)) Rate from (
                    select sum(ThisAmt)ThisAmtTot ,sum(LastAmt) LastAmtTot from ( select  ThisAmt ,LastAmt from (
                    select sum(CAST(House_sendcase.get_amount AS DECIMAL))ThisAmt,Year(get_amount_date) yyyy,U_BC from House_sendcase
                    LEFT JOIN House_apply on House_apply.HA_id = House_sendcase.HA_id
                    LEFT JOIN (select U_num ,U_BC,U_name FROM User_M ) User_M ON User_M.U_num = House_apply.plan_num
                    where House_sendcase.del_tag = '0' AND U_BC not in ('BC0700','BC0800') AND House_apply.del_tag='0'
                    AND isnull(House_sendcase.get_amount,'') <> '' AND convert(varchar(7),(get_amount_date),126)
                    between convert(varchar(4), @YYYYMM+'-01', 126)+'-01' and  @YYYYMM";
                if (!string.IsNullOrEmpty(U_BC))
                {
                    T_SQL += " And U_BC=@U_BC";
                }
                T_SQL += " group by Year(get_amount_date),U_BC ) T";
                T_SQL += " left join (";
                T_SQL += " select sum(CAST(House_sendcase.get_amount AS DECIMAL))LastAmt,Year(get_amount_date) yyyy,U_BC from House_sendcase";
                T_SQL += " LEFT JOIN House_apply on House_apply.HA_id = House_sendcase.HA_id";
                T_SQL += " LEFT JOIN (select U_num ,U_BC,U_name  FROM User_M  ) User_M ON User_M.U_num = House_apply.plan_num";
                T_SQL += " where House_sendcase.del_tag = '0' AND U_BC not in ('BC0700','BC0800') AND House_apply.del_tag='0'";
                T_SQL += " AND isnull(House_sendcase.get_amount,'') <> '' AND convert(varchar(7),(get_amount_date),126)";
                T_SQL += " between convert(varchar(4),  DATEADD(year,-1,@YYYYMM+'-01'), 126)+'-01'";
                T_SQL += " and convert(varchar(7),DATEADD(year,-1,@YYYYMM+'-01'), 126)";
                if (!string.IsNullOrEmpty(U_BC))
                {
                    T_SQL += " And U_BC=@U_BC";
                    parameters.Add(new SqlParameter("@U_BC", U_BC));
                }
                T_SQL += " group by Year(get_amount_date),U_BC ) L on T.U_BC=L.U_BC ) A ) T";
                parameters.Add(new SqlParameter("@YYYYMM", YYYYMM));
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
        /// GetTargetPerfByYYYY/_Ajaxhandler.asp
        /// </summary>
        /// <param name="this_YY">2024</param>
        /// <param name="ThisTot">249855</param>
        /// <param name="U_BC">BC0100</param>
        [HttpGet("GetTargetPerfByYYYY")]
        public ActionResult<ResultClass<string>> GetTargetPerfByYYYY(string this_YY, int ThisTot, string? U_BC)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"
                    select *,CAST(ROUND(CAST(@ThisTot AS decimal(8,0)) / CAST(Target AS decimal(8,0)), 2)*100 AS decimal(3,0)) TargetRate,
                    CAST(ROUND(CAST(Target-@ThisTot AS decimal(8,0)) / CAST(Target AS decimal(8,0)), 2)*100 AS decimal(3,0)) UnTargetRate
                    from ( SELECT KeyVal yyyy,count(ColumnVal) BC_Count, sum(convert(decimal, ColumnVal)) Target FROM LogTable
                    WHERE TableNA = 'Performance_Plot' and KeyVal=@this_YY";
                if (!string.IsNullOrEmpty(U_BC))
                {
                    T_SQL += " and ColumnNA=@U_BC";
                    parameters.Add(new SqlParameter("@U_BC", U_BC));
                }
                T_SQL += " group by KeyVal ) A";
                parameters.Add(new SqlParameter("@this_YY", this_YY));
                parameters.Add(new SqlParameter("@ThisTot", ThisTot));
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
        /// GetSalesByYYYYMM/_Ajaxhandler.asp
        /// </summary>
        /// <param name="YYYYMM">2024-09</param>
        /// <param name="U_BC">BC0100</param>
        [HttpGet("GetSalesByYYYYMM")]
        public ActionResult<ResultClass<string>> GetSalesByYYYYMM(string YYYYMM, string? U_BC)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = "";
                if (!string.IsNullOrEmpty(U_BC))
                {
                    T_SQL += " select @YYYYMM yyyymm,E.U_BC,U_BC_Name,count(*) UserCount,convert(varchar(3),(year(@YYYYMM+'-01')-1911))+'-'+convert(varchar(2),month(@YYYYMM+'-01')) yyymm";
                    T_SQL += " from [dbo].[fun_GetWorkUsers] (@YYYYMM) E";
                    T_SQL += " Left Join  dbo.view_Sales_BC B on E.U_BC=B.U_BC where B.U_BC is not null and E.U_BC=@U_BC";
                    T_SQL += " and U_PFT in ('PFT010','PFT015','PFT020','PFT030','PFT050','PFT060','PFT300')  and E.U_BC=@U_BC group by  E.U_BC,U_BC_Name";
                    T_SQL += " union all";
                    T_SQL += " select convert(varchar(7),  DATEADD(year,-1,@YYYYMM+'-01'), 126) yyyymm, E.U_BC,U_BC_Name ,count(*) UserCount";
                    T_SQL += " ,convert(varchar(3), (year(DATEADD(year,-1,@YYYYMM+'-01'))-1911))+'-'+convert(varchar(2), month(DATEADD(year,-1,@YYYYMM+'-01'))) yyymm";
                    T_SQL += " from [dbo].[fun_GetWorkUsers] (convert(varchar(7),  DATEADD(year,-1,@YYYYMM+'-01'), 126)) E";
                    T_SQL += " Left Join  dbo.view_Sales_BC B on E.U_BC=B.U_BC where B.U_BC is not null";
                    T_SQL += " and U_PFT in ('PFT010','PFT015','PFT020','PFT030','PFT050','PFT060','PFT300')  and E.U_BC=@U_BC group by  E.U_BC,U_BC_Name";
                    parameters.Add(new SqlParameter("@U_BC", U_BC));
                }
                else
                {
                    T_SQL += " select  @YYYYMM yyyymm,count(*) UserCount,convert(varchar(3), (year(@YYYYMM+'-01')-1911))+'-'+convert(varchar(2), month(@YYYYMM+'-01')) yyymm";
                    T_SQL += " from [dbo].[fun_GetWorkUsers] (@YYYYMM) E";
                    T_SQL += " Left Join  dbo.view_Sales_BC B on E.U_BC=B.U_BC where B.U_BC is not null";
                    T_SQL += " and U_PFT in ('PFT010','PFT015','PFT020','PFT030','PFT050','PFT060','PFT300')";
                    T_SQL += " union all";
                    T_SQL += " select convert(varchar(7),DATEADD(year,-1,@YYYYMM+'-01'), 126) yyyy,count(*)UserCount";
                    T_SQL += " ,convert(varchar(3), (year(DATEADD(year,-1,@YYYYMM+'-01'))-1911))+'-'+convert(varchar(2), month(DATEADD(year,-1,@YYYYMM+'-01'))) yyymm";
                    T_SQL += " from [dbo].[fun_GetWorkUsers] (convert(varchar(7), DATEADD(year,-1,@YYYYMM+'-01'), 126)) E";
                    T_SQL += " Left Join  dbo.view_Sales_BC B on E.U_BC=B.U_BC where B.U_BC is not null";
                    T_SQL += " and U_PFT in ('PFT010','PFT015','PFT020','PFT030','PFT050','PFT060','PFT300')";
                }
                parameters.Add(new SqlParameter("@YYYYMM", YYYYMM));
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
        /// GetBCPerfByYYYY/_Ajaxhandler.asp
        /// </summary>
        /// <param name="YYYY">2024</param>
        /// <param name="YYYYMM">2024-09</param>
        /// <param name="fund_Comp">FDCOM003</param>
        [HttpGet("GetBCPerfByYYYY")]
        public ActionResult<ResultClass<string>> GetBCPerfByYYYY(string YYYY, string YYYYMM, string? fund_company)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"
                    SELECT sum(CAST(H.get_amount AS DECIMAL))amount, Year(get_amount_date) yyyy,U_BC,BC_NA,item_sort FROM House_sendcase H
                    LEFT JOIN House_apply A ON A.HA_id = H.HA_id
                    LEFT JOIN (select U.U_num,U.U_BC,UB.item_D_name BC_NA,item_sort from User_M u
                    LEFT JOIN Item_list ub ON ub.item_M_code='branch_company' AND ub.item_D_type='Y' AND ub.item_D_code=u.U_BC
                    ) U on A.plan_num=U.U_num
                    WHERE H.del_tag = '0' AND A.del_tag='0' AND U_BC not in ('BC0700','BC0800') AND isnull(H.get_amount, '') <> ''
                    AND convert(varchar(7), (get_amount_date), 126) between @YYYY+'-01' and @YYYYMM";
                if (!string.IsNullOrEmpty(fund_company))
                {
                    T_SQL += " and fund_company=@fund_company";
                    parameters.Add(new SqlParameter("@fund_company", fund_company));
                }
                T_SQL += " GROUP BY Year(get_amount_date) ,U_BC,BC_NA,item_sort";
                T_SQL += " ORDER BY item_sort ASC";
                parameters.Add(new SqlParameter("@YYYY", YYYY));
                parameters.Add(new SqlParameter("@YYYYMM", YYYYMM));
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
        /// GetPerfByYYYY_BC/_Ajaxhandler.asp
        /// </summary>
        /// <param name="this_YY">2024-09</param>
        /// <param name="last_YY">2023</param>
        /// <param name="U_BC">BC0100</param>
        /// <param name="fund_Comp">FDCOM003</param>
        [HttpGet("GetPerfByYYYY_BC")]
        public ActionResult<ResultClass<string>> GetPerfByYYYY_BC(string this_YY, string last_YY, string U_BC, string? fund_company)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"
                    select M.*,InCount,RateCount,CAST(ROUND(CAST(InCount AS decimal(8,2))/CAST(RateCount AS decimal(8,2)),2)*100 AS decimal(3,0)) TRate
                    from (SELECT sum(CAST(H.get_amount AS DECIMAL)) amount,convert(varchar(7),get_amount_date,126)yyyymm,month(get_amount_date) mm,Year(get_amount_date) yyyy
                    ,U_BC,BC_NA FROM House_sendcase H
                    LEFT JOIN House_apply A ON A.HA_id = H.HA_id
                    LEFT JOIN (select U.U_num,U.U_BC,UB.item_D_name BC_NA from User_M u
                    LEFT JOIN Item_list ub ON ub.item_M_code='branch_company' AND ub.item_D_type='Y' AND ub.item_D_code=u.U_BC
                    ) U on A.plan_num=U.U_num
                    WHERE H.del_tag = '0'  AND A.del_tag='0' AND isnull(H.get_amount, '') <> ''
                    AND convert(varchar(7),(get_amount_date),126) between @last_YY+'-01' and @this_YY";
                if (!string.IsNullOrEmpty(fund_company))
                {
                    T_SQL += " and fund_company=@fund_company";
                    parameters.Add(new SqlParameter("@fund_company", fund_company));
                }
                T_SQL += " GROUP BY convert(varchar(7),get_amount_date,126),Year(get_amount_date),month(get_amount_date),U_BC,BC_NA) M";
                T_SQL += " Left Join (select U_BC,yyyymm,sum(RateCount)RateCount from fun_GetRateCount(@last_YY+'-01',@this_YY)";
                T_SQL += " group by U_BC,yyyymm ) R on M.U_BC=R.U_BC and M.yyyymm=R.yyyymm";
                T_SQL += " Left Join (select U_BC,yyyymm,sum(InCount)InCount from fun_GetInCount(@last_YY+'-01',@this_YY)";
                T_SQL += " group by U_BC,yyyymm ) I on M.U_BC=I.U_BC and M.yyyymm=I.yyyymm";
                T_SQL += " where  M.U_BC=@U_BC ORDER BY M.yyyymm ASC";
                parameters.Add(new SqlParameter("@last_YY", last_YY));
                parameters.Add(new SqlParameter("@this_YY", this_YY));
                parameters.Add(new SqlParameter("@U_BC", U_BC));
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
        /// SetTargetPerfByYYYY_BC/_Ajaxhandler.asp
        /// </summary>
        /// <param name="YYYY">2024</param>
        /// <param name="U_BC">BC0100</param>
        /// <param name="isSet">Y</param>
        [HttpGet("SetTargetPerfByYYYY_BC")]
        public ActionResult<ResultClass<string>> SetTargetPerfByYYYY_BC(int YYYY, string U_BC, string isSet)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                if (isSet == "N")
                    YYYY = YYYY - 1;

                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"
                    select D.*,isnull(M.Target,'0')[Target] from (SELECT item_D_code U_BC,UB.item_D_name U_BC_Name,item_sort
                    FROM Item_list ub
                    WHERE ub.item_M_code='branch_company' AND ub.item_D_type='Y' ) D
                    Left join (SELECT KeyVal yyyy,convert(decimal,ColumnVal) Target,ColumnNA U_BC
                    FROM LogTable L  WHERE TableNA = 'Performance_Plot' and KeyVal=@YYYY) M
                    on m.U_BC=D.U_BC where D.U_BC not in ('BC0700','BC0800')
                    and  D.U_BC=@U_BC order by item_sort";
                parameters.Add(new SqlParameter("@YYYY", YYYY));
                parameters.Add(new SqlParameter("@U_BC", U_BC));
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
        /// 異動年度目標SaveBC_Target/_Ajaxhandler.asp
        /// </summary>
        [HttpPost("SaveBC_Target")]
        public ActionResult<ResultClass<string>> SaveBC_Target(string YYYY, string U_BC, Target_YYYY model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var User_Num = HttpContext.Session.GetString("UserID");

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL_Delete
                var parameters_d = new List<SqlParameter>();
                var T_SQL_D = "Delete LogTable where TableNA='Performance_Plot' and KeyVal=@YYYY and ColumnNA=@U_BC";
                parameters_d.Add(new SqlParameter("@U_BC", U_BC));
                parameters_d.Add(new SqlParameter("@YYYY", YYYY));
                #endregion
                int result_d = _adoData.ExecuteNonQuery(T_SQL_D, parameters_d);
                if (result_d == 0)
                {
                    resultClass.ResultCode = "400";
                    resultClass.ResultMsg = "刪除失敗";
                    return BadRequest(resultClass);
                }
                else
                {
                    #region SQL
                    var parameters = new List<SqlParameter>();
                    var T_SQL = "Insert into LogTable(TableNA,KeyVal,ColumnNA,ColumnVal,LogID,LogDate) Values";
                    T_SQL += "('Performance_Plot',@KeyVal,@ColumnNA,@ColumnVal,@LogID,GETDATE())";
                    parameters.Add(new SqlParameter("@KeyVal", YYYY));
                    parameters.Add(new SqlParameter("@ColumnNA", model.U_BC));
                    parameters.Add(new SqlParameter("@ColumnVal", model.Target));
                    parameters.Add(new SqlParameter("@LogID", User_Num));
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
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass);
            }
        }
        /// <summary>
        /// GetSalesListByYYYYMM/_Ajaxhandler.asp
        /// </summary>
        /// <param name="YYYYMM">2024-09</param>
        /// <param name="U_BC">BC0100</param>
        [HttpGet("GetSalesListByYYYYMM")]
        public ActionResult<ResultClass<string>> GetSalesListByYYYYMM(string U_BC, string YYYYMM)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"
                    select E.u_name,PFT_NA,U_BC_Name,convert(varchar(3),(year(U_arrive_date)-1911))+'-'+convert(varchar(2),month(U_arrive_date))+'-'+convert(varchar(2),Day(U_arrive_date)) U_arrive_date,
                    convert(varchar(3),(year(U_leave_date)-1911))+'-'+convert(varchar(2),month(U_leave_date))+'-'+convert(varchar(2),Day(U_leave_date)) U_leave_date,
                    case when U_leave_date is null then 'Y' else 'N' end isWork,
                    DATEDIFF(MONTH,E.U_arrive_date,case when U_leave_date is null then GETDATE() else U_leave_date end)/12 AS JobYY,
                    DATEDIFF(day,E.U_arrive_date,case when U_leave_date is null then GETDATE() else U_leave_date end ) % 365/30 AS JobMM
                    from [dbo].[fun_GetWorkUsers] (@YYYYMM) E
                    Left Join  dbo.view_Sales_BC B on E.U_BC=B.U_BC
                    Left Join (select item_D_name PFT_NA,item_D_code U_PFT,item_sort from item_list where item_D_type='Y' and item_M_code ='professional_title') I
                    on E.U_PFT=I.U_PFT
                    where B.U_BC is not null and E.U_PFT in ('PFT010','PFT015','PFT020','PFT030','PFT050','PFT060','PFT300')
                    and E.U_BC=@U_BC order by E.U_BC, item_sort,E.U_arrive_date asc";
                parameters.Add(new SqlParameter("@YYYYMM", YYYYMM));
                parameters.Add(new SqlParameter("@U_BC", U_BC));
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
        /// 提供登入者的所屬分公司
        /// </summary>
        [HttpGet("GetUserBC")]
        public ActionResult<ResultClass<string>> GetUserBC()
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var User_U_BC = HttpContext.Session.GetString("User_U_BC");

            if (!string.IsNullOrEmpty(User_U_BC))
            {
                resultClass.ResultCode = "000";
                resultClass.objResult = JsonConvert.SerializeObject(User_U_BC);
                return Ok(resultClass);
            }
            else
            {
                resultClass.ResultCode = "400";
                resultClass.ResultMsg = "查無資料";
                return BadRequest(resultClass);
            }
        }
        #endregion

        #region 客戶類型資料查詢
        /// <summary>
        /// 提供所有工作類型 GetJobKindList/CS_ListByJob.asp
        /// </summary>
        [HttpGet("GetJobKindList")]
        public ActionResult<ResultClass<string>> GetJobKindList()
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = "select item_D_code,item_D_name from Item_list where item_M_code = 'job_kind' AND item_D_type='Y' AND show_tag='0' AND del_tag='0' order by item_sort";
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
        /// 客戶類型資料_查詢 CS_ListByJob_Query/CS_ListByJob.asp
        /// </summary>
        [HttpPost("CS_ListByJob_Query")]
        public ActionResult<ResultClass<string>> CS_ListByJob_Query(CS_ListByJob_req model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"
                    SELECT H.CS_name,H.CS_birthday,H.CS_MTEL1,H.CS_company_name,H.CS_company_tel,I1.job_kind_na
                    ,H.CS_job_title,H.CS_job_years,I.income_na,H.CS_income_everymonth FROM House_apply H
                    left join (select item_D_code cs_income_way,item_D_name income_na from Item_list where item_M_code = 'income_way' AND item_D_type='Y') I
                    on H.cs_income_way=I.cs_income_way
                    left join (select item_D_code cs_job_kind,case when item_D_code='jobk003' then '退休人士(其他)' else item_D_name end job_kind_na
                    from Item_list where item_M_code = 'job_kind' AND item_D_type='Y' ) I1
                    on H.cs_job_kind=I1.cs_job_kind WHERE H.del_tag = '0' AND plan_type='plan_T003'
                    AND plan_type_date between @Pre_Agency_Date_S and @Pre_Agency_Date_E";
                if (!string.IsNullOrEmpty(model.job_kind))
                {
                    T_SQL += " AND H.cs_job_kind=@job_kind";
                    parameters.Add(new SqlParameter("@job_kind", model.job_kind));
                }
                parameters.Add(new SqlParameter("@Pre_Agency_Date_S", model.Pre_Agency_Date_S));
                parameters.Add(new SqlParameter("@Pre_Agency_Date_E", model.Pre_Agency_Date_E));
                #endregion
                DataTable dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
                if (dtResult.Rows.Count > 0)
                {
                    resultClass.ResultCode = "000";
                    DataTable pageData = FuncHandler.GetPage(dtResult, model.page, 100);
                    resultClass.objResult = JsonConvert.SerializeObject(pageData);
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
        /// 客戶類型資料_下載 CS_ListByJob_Excel/CS_ListByJob.asp
        /// </summary>
        [HttpPost("CS_ListByJob_Excel")]
        public IActionResult CS_ListByJob_Excel(CS_ListByJob_req model)
        {
            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"
                    SELECT H.CS_name,H.CS_birthday,H.CS_MTEL1,H.CS_company_name,H.CS_company_tel,I1.job_kind_na
                    ,H.CS_job_title,H.CS_job_years,I.income_na,H.CS_income_everymonth FROM House_apply H
                    left join (select item_D_code cs_income_way,item_D_name income_na from Item_list where item_M_code = 'income_way' AND item_D_type='Y') I
                    on H.cs_income_way=I.cs_income_way
                    left join (select item_D_code cs_job_kind,case when item_D_code='jobk003' then '退休人士(其他)' else item_D_name end job_kind_na
                    from Item_list where item_M_code = 'job_kind' AND item_D_type='Y' ) I1
                    on H.cs_job_kind=I1.cs_job_kind WHERE H.del_tag = '0' AND plan_type='plan_T003'
                    AND plan_type_date between @Pre_Agency_Date_S and @Pre_Agency_Date_E";
                if (!string.IsNullOrEmpty(model.job_kind))
                {
                    T_SQL += " AND H.cs_job_kind=@job_kind";
                    parameters.Add(new SqlParameter("@job_kind", model.job_kind));
                }
                parameters.Add(new SqlParameter("@Pre_Agency_Date_S", model.Pre_Agency_Date_S));
                parameters.Add(new SqlParameter("@Pre_Agency_Date_E", model.Pre_Agency_Date_E));
                #endregion
                DataTable dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
                if (dtResult.Rows.Count > 0)
                {
                    var ExcelList = dtResult.AsEnumerable().Select(row => new CS_ListByJob_Excel
                    {
                        CS_name = row.Field<string>("CS_name"),
                        CS_birthday = row.Field<string>("CS_birthday"),
                        CS_MTEL1 = row.Field<string>("CS_MTEL1"),
                        CS_company_name = row.Field<string>("CS_company_name"),
                        CS_company_tel = row.Field<string>("CS_company_tel"),
                        job_kind_na = row.Field<string>("job_kind_na"),
                        CS_job_title = row.Field<string>("CS_job_title"),
                        CS_job_years = row.Field<string>("CS_job_years"),
                        income_na = row.Field<string>("income_na"),
                        CS_income_everymonth = row.Field<string>("CS_income_everymonth")
                    }).ToList();
                    var Excel_Headers = new Dictionary<string, string>
                    {
                        { "CS_name", "申請人" },
                        { "CS_birthday", "生日" },
                        { "CS_MTEL1", "行動電話" },
                        { "CS_company_name", "公司名稱" },
                        { "CS_company_tel", "公司電話" },
                        { "job_kind_na", "工作類型" },
                        { "CS_job_title", "職稱" },
                        { "CS_job_years", "年資" },
                        { "income_na", "收入方式" },
                        { "CS_income_everymonth", "評估每月收入" }
                    };
                    var fileBytes = FuncHandler.ExportToExcel(ExcelList, Excel_Headers);
                    var fileName = "客戶類型資料" + DateTime.Now.ToString("yyyyMMddHHmm") + ".xlsx";
                    return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
                }
                else
                {
                    return NotFound(); // 檔案不存在時返回 404
                }
            }
            catch (Exception)
            {
                return StatusCode(500);
            }
        }
        #endregion

        #region 客戶來源分析表(網路)
        /// <summary>
        /// GetCallCneterIntData/_Ajaxhandler.asp
        /// </summary>
        /// <param name="InDate_S">113/9/1</param>
        /// <param name="InDate_E">113/9/25</param>
        [HttpGet("GetCallCneterIntData")]
        public ActionResult<ResultClass<string>> GetCallCneterIntData(string InDate_S, string InDate_E)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                InDate_S = FuncHandler.ConvertROCToGregorian(InDate_S);
                InDate_E = FuncHandler.ConvertROCToGregorian(InDate_E);

                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"
                    SELECT li.item_D_name,li.item_D_code, COUNT(*) AS TelSourCount FROM TelemarketingCSList tc
                    left join Telemarketing_M tm on tm.HA_id = tc.TC_ID
                    left join Item_list li on li.item_D_code=tc.TelSour
                    where ISNULL(li.item_D_name,'') <> '' and ISNULL(InDate,'') <> ''
                    and convert(datetime,convert(varchar(4),(convert(varchar(3),InDate,126)+1911)) + SUBSTRING(InDate,4,10)) between @InDate_S and @InDate_E
                    GROUP BY li.item_D_name,li.item_D_code ORDER BY li.item_D_name DESC";
                parameters.Add(new SqlParameter("@InDate_S", InDate_S));
                parameters.Add(new SqlParameter("@InDate_E", InDate_E));
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
        /// GetCallCneterIntDataDetail/_Ajaxhandler.asp
        /// </summary>
        /// <param name="TelSour">SOU04</param>
        /// <param name="InDate_S">113/9/1</param>
        /// <param name="InDate_E">113/9/25</param>
        [HttpGet("GetCallCneterIntDataDetail")]
        public ActionResult<ResultClass<string>> GetCallCneterIntDataDetail(string TelSour, string InDate_S, string InDate_E)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                InDate_S = FuncHandler.ConvertROCToGregorian(InDate_S);
                InDate_E = FuncHandler.ConvertROCToGregorian(InDate_E);

                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"
                    SELECT Case When ISnull(cs_rname,'')='' Then cs_name Else cs_rname END AS cs_name,CS_ID,cs_mtel1,li.item_D_name,cs_register_address
                    ,(select top 1 REPLACE(REPLACE(REPLACE(REPLACE(Memo,CHAR(13), ''),CHAR(10), ''),CHAR(9), ''),CHAR(8), '') from Telemarketing_Log tl where td.TM_id=tl.TM_id and tl.TM_D_id=td.TM_D_id order by add_date desc) as CS_remark
                    FROM view_Telemarketing_source tc
                    left join Telemarketing_M tm on tm.HA_id = tc.HA_id
                    left join view_Telemarketing_Curr td ON tm.TM_id = td.tm_id
                    left join Item_list li on li.item_D_code=tc.TelSour
                    where ISNULL(li.item_D_name,'') <> '' and ISNULL(InDate,'') <> ''
                    and convert(datetime, convert(varchar(4),(convert(varchar(3), InDate, 126)+1911)) + SUBSTRING(InDate,4,10)) between @InDate_S and @InDate_E
                    and tc.TelSour=@TelSour and tc.TM_type='2'
                    order by convert(datetime, convert(varchar(4),(convert(varchar(3), InDate, 126)+1911)) + SUBSTRING(InDate,4,10))";
                parameters.Add(new SqlParameter("@InDate_S", InDate_S));
                parameters.Add(new SqlParameter("@InDate_E", InDate_E));
                parameters.Add(new SqlParameter("@TelSour", TelSour));
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

        #region 重復件數據分析表
        /// <summary>
        /// 提供查詢來源資料 GetTelSourList/Repeat_IntSource_Report.asp
        /// </summary>
        [HttpGet("GetTelSourList")]
        public ActionResult<ResultClass<string>> GetTelSourList()
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = "select item_D_code,item_D_name from Item_list where item_M_code='TelSour' and item_D_type='Y' and del_tag='0'";
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
        /// GetRepeatIntData/_Ajaxhandler.asp
        /// </summary>
        /// <param name="Sour">SOU01</param>
        /// <param name="InDate_S">113/9/1</param>
        /// <param name="InDate_E">113/9/25</param>
        [HttpGet("GetRepeatIntData")]
        public ActionResult<ResultClass<string>> GetRepeatIntData(string Sour, string InDate_S, string InDate_E)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                InDate_S = FuncHandler.ConvertROCToGregorian(InDate_S);
                InDate_E = FuncHandler.ConvertROCToGregorian(InDate_E);

                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"
                    SELECT li.item_D_name,li.item_D_code, COUNT(*) AS AskSourCount
                    FROM TelemarketingCSList tc
                    left join Item_list li on li.item_D_code=tc.TelAsk where ISNULL(li.item_D_name,'') <> ''
                    and ISNULL(repeat, '') <> 'Y' and ISNULL(InDate,'') <> ''
                    and convert(datetime, convert(varchar(4),(convert(varchar(3), InDate, 126)+1911)) + SUBSTRING(InDate,4,10)) between @InDate_S and @InDate_E
                    and TelSour = @Sour and ISNULL(tc.TelAsk,'') <> '' GROUP BY li.item_D_name,item_D_code
                    union all
                    SELECT 'Repeat' AS item_D_name,'Repeat' AS item_D_code, COUNT(*) AS AskSourCount
                    FROM TelemarketingCSList tc
                    WHERE ISNULL(repeat, '') = 'Y' and ISNULL(tc.TelAsk,'') <> '' and ISNULL(InDate,'') <> ''
                    and convert(datetime, convert(varchar(4),(convert(varchar(3), InDate, 126)+1911)) + SUBSTRING(InDate,4,10)) between @InDate_S and @InDate_E
                    and TelSour = @Sour";
                parameters.Add(new SqlParameter("@InDate_S", InDate_S));
                parameters.Add(new SqlParameter("@InDate_E", InDate_E));
                parameters.Add(new SqlParameter("@Sour", Sour));
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
        /// GetRepeatIntDataTable/_Ajaxhandler.asp
        /// </summary>
        /// <param name="Sour">SOU01</param>
        /// <param name="YYY">113</param>
        [HttpGet("GetRepeatIntDataTable")]
        public ActionResult<ResultClass<string>> GetRepeatIntDataTable(string Sour, string YYY)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"
                    SELECT convert(varchar(4),(convert(varchar(3), InDate, 126)+1911))+ '-' +RIGHT('0' + CAST(SUBSTRING(InDate, CHARINDEX('/', InDate) + 1, CHARINDEX('/', InDate, CHARINDEX('/', InDate) + 1) - CHARINDEX('/', InDate) - 1) AS NVARCHAR(2)), 2) AS InDate,
                    SUM(CASE WHEN li.item_D_code = 'ASK03'AND ISNULL(repeat, '') <> 'Y' THEN 1 ELSE 0 END) AS ASK03,
                    SUM(CASE WHEN li.item_D_code = 'ASK01'AND ISNULL(repeat, '') <> 'Y' THEN 1 ELSE 0 END) AS ASK01,
                    SUM(CASE WHEN li.item_D_code = 'ASK02'AND ISNULL(repeat, '') <> 'Y' THEN 1 ELSE 0 END) AS ASK02,
                    SUM(CASE WHEN li.item_D_code = 'ASK04'AND ISNULL(repeat, '') <> 'Y' THEN 1 ELSE 0 END) AS ASK04,
                    SUM(CASE WHEN li.item_D_code = 'ASK05'AND ISNULL(repeat, '') <> 'Y' THEN 1 ELSE 0 END) AS ASK05,
                    SUM(CASE WHEN li.item_D_code = 'ASK06'AND ISNULL(repeat, '') <> 'Y' THEN 1 ELSE 0 END) AS ASK06,
                    SUM(CASE WHEN li.item_D_code IN ('ASK03', 'ASK01', 'ASK02', 'ASK04', 'ASK05','ASK06') and ISNULL(repeat, '') = 'Y' THEN 1 ELSE 0 END) AS repeat
                    FROM TelemarketingCSList tc LEFT JOIN Item_list li ON li.item_D_code = tc.TelAsk
                    WHERE TelSour = @Sour and ISNULL(tc.TelAsk,'') <> '' and left(InDate,3)=@YYY
                    GROUP BY convert(varchar(4),(convert(varchar(3), InDate, 126)+1911))+ '-' +RIGHT('0' + CAST(SUBSTRING(InDate, CHARINDEX('/', InDate) + 1, CHARINDEX('/', InDate, CHARINDEX('/', InDate) + 1) - CHARINDEX('/', InDate) - 1) AS NVARCHAR(2)), 2)
                    ORDER BY convert(varchar(4),(convert(varchar(3), InDate, 126)+1911))+ '-' +RIGHT('0' + CAST(SUBSTRING(InDate, CHARINDEX('/', InDate) + 1, CHARINDEX('/', InDate, CHARINDEX('/', InDate) + 1) - CHARINDEX('/', InDate) - 1) AS NVARCHAR(2)), 2)";
                parameters.Add(new SqlParameter("@Sour", Sour));
                parameters.Add(new SqlParameter("@YYY", YYY));
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
        /// GetRepeatIntDataDetail/_Ajaxhandler.asp
        /// </summary>
        /// <param name="TelAsk">ASK06</param>
        /// <param name="InDate_S">113/9/1</param>
        /// <param name="InDate_E">113/9/25</param>
        /// <param name="Sour">SOU01</param>
        [HttpGet("GetRepeatIntDataDetail")]
        public ActionResult<ResultClass<string>> GetRepeatIntDataDetail(string TelAsk, string InDate_S, string InDate_E, string Sour)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                InDate_S = FuncHandler.ConvertROCToGregorian(InDate_S);
                InDate_E = FuncHandler.ConvertROCToGregorian(InDate_E);

                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = "";
                if (TelAsk != "Repeat")
                {
                    T_SQL += " SELECT Case When ISnull(tc.cs_rname,'')='' Then tc.cs_name Else tc.cs_rname END AS cs_name,tc.CS_ID,tc.cs_mtel1,li.item_D_name,tc.cs_register_address";
                    T_SQL += " ,(select top 1 REPLACE(REPLACE(REPLACE(REPLACE(Memo,CHAR(13), ''),CHAR(10), ''),CHAR(9), ''),CHAR(8), '') from Telemarketing_Log tl where td.TM_id=tl.TM_id and tl.TM_D_id=td.TM_D_id order by add_date desc) as CS_remark";
                    T_SQL += " FROM view_Telemarketing_source tc";
                    T_SQL += " inner join TelemarketingCSList tcs on tcs.TC_ID=tc.ha_id";
                    T_SQL += " left join Telemarketing_M tm on tm.HA_id = tc.HA_id";
                    T_SQL += " left join view_Telemarketing_Curr td ON tm.TM_id = td.tm_id";
                    T_SQL += " left join Item_list li on li.item_D_code=tc.TelAsk where ISNULL(li.item_D_name,'') <> ''";
                    T_SQL += " and ISNULL(repeat, '') <> 'Y' and ISNULL(tc.InDate,'') <> ''";
                    T_SQL += " and convert(datetime, convert(varchar(4),(convert(varchar(3), tc.InDate, 126)+1911)) + SUBSTRING(tc.InDate,4,10)) between @InDate_S and @InDate_E";
                    T_SQL += " and tc.TelSour = @Sour and tc.TM_type='2' and ISNULL(tc.TelAsk,'') <> '' and tc.TelAsk=@TelAsk";
                }
                else
                {
                    T_SQL += " SELECT Case When ISnull(tc.cs_rname,'')='' Then tc.cs_name Else tc.cs_rname END AS cs_name,tc.CS_ID,tc.cs_mtel1,'Repeat' AS item_D_name,tc.cs_register_address";
                    T_SQL += " ,(select top 1 REPLACE(REPLACE(REPLACE(REPLACE(Memo,CHAR(13), ''),CHAR(10), ''),CHAR(9), ''),CHAR(8), '') from Telemarketing_Log tl where td.TM_id=tl.TM_id and tl.TM_D_id=td.TM_D_id order by add_date desc) as CS_remark";
                    T_SQL += " FROM view_Telemarketing_source tc";
                    T_SQL += " inner join TelemarketingCSList tcs on tcs.TC_ID=tc.ha_id";
                    T_SQL += " left join Telemarketing_M tm on tm.HA_id = tc.HA_id";
                    T_SQL += " left join view_Telemarketing_Curr td ON tm.TM_id = td.tm_id";
                    T_SQL += " WHERE ISNULL(repeat, '') = 'Y' and ISNULL(tc.TelAsk,'') <> '' and ISNULL(tc.InDate,'') <> ''";
                    T_SQL += " and convert(datetime, convert(varchar(4),(convert(varchar(3), tc.InDate, 126)+1911)) + SUBSTRING(tc.InDate,4,10)) between @InDate_S and @InDate_E";
                    T_SQL += " and tc.TelSour = @Sour and tc.TM_type='2'";
                }
                parameters.Add(new SqlParameter("@InDate_S", InDate_S));
                parameters.Add(new SqlParameter("@InDate_E", InDate_E));
                parameters.Add(new SqlParameter("@Sour", Sour));
                parameters.Add(new SqlParameter("@TelAsk", TelAsk));
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

        #region 網路電銷房貸狀態分析表
        /// <summary>
        /// GetCallMortgageIntData/_Ajaxhandler.asp
        /// </summary>
        /// <param name="InDate_S">113/9/1</param>
        /// <param name="InDate_E">113/9/25</param>
        [HttpGet("GetCallMortgageIntData")]
        public ActionResult<ResultClass<string>> GetCallMortgageIntData(string InDate_S, string InDate_E)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                InDate_S = FuncHandler.ConvertROCToGregorian(InDate_S);
                InDate_E = FuncHandler.ConvertROCToGregorian(InDate_E);
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"
                    SELECT CASE WHEN Fin_type IN ('FIN_T13', 'FIN_T22', 'FIN_T09', 'FIN_T23', 'FIN_T25') THEN Fin_name WHEN Fin_type IN ('FIN_T02','FIN_T28') THEN 'FIN' END AS Fin_name,
                    CASE WHEN Fin_type IN ('FIN_T13', 'FIN_T22', 'FIN_T09', 'FIN_T23', 'FIN_T25') THEN Fin_type WHEN Fin_type IN ('FIN_T02','FIN_T28') THEN 'FIN_T00' END AS Fin_type,
                    COUNT(*) AS Count FROM Telemarketing_M tm
                    JOIN view_Telemarketing_Curr td ON tm.TM_id = td.tm_id
                    JOIN view_Telemarketing_source ha ON tm.HA_id = ha.HA_id AND ha.TM_type = '2'
                    WHERE TelAsk = 'ASK01' and ISNULL(InDate,'') <> '' and Fin_type IN ('FIN_T13', 'FIN_T22', 'FIN_T09', 'FIN_T23', 'FIN_T25','FIN_T02','FIN_T28')
                    and convert(datetime, convert(varchar(4),(convert(varchar(3), InDate, 126)+1911)) + SUBSTRING(InDate,4,10)) between @InDate_S and @InDate_E
                    GROUP BY CASE  WHEN Fin_type IN ('FIN_T13', 'FIN_T22', 'FIN_T09', 'FIN_T23', 'FIN_T25') THEN Fin_name
                    WHEN Fin_type IN ('FIN_T02','FIN_T28') THEN 'FIN' END
                    ,CASE WHEN Fin_type IN ('FIN_T13', 'FIN_T22', 'FIN_T09', 'FIN_T23', 'FIN_T25') THEN Fin_type WHEN Fin_type IN ('FIN_T02','FIN_T28') THEN 'FIN_T00' END";
                parameters.Add(new SqlParameter("@InDate_S", InDate_S));
                parameters.Add(new SqlParameter("@InDate_E", InDate_E));
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
        /// GetCallMortgageIntDataDetail/_Ajaxhandler.asp
        /// </summary>
        /// <param name="Fin_type">FIN_T09</param>
        /// <param name="InDate_S">113/9/1</param>
        /// <param name="InDate_E">113/9/25</param>
        [HttpGet("GetCallMortgageIntDataDetail")]
        public ActionResult<ResultClass<string>> GetCallMortgageIntDataDetail(string Fin_type, string InDate_S, string InDate_E)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                InDate_S = FuncHandler.ConvertROCToGregorian(InDate_S);
                InDate_E = FuncHandler.ConvertROCToGregorian(InDate_E);
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"
                    select Case When ISnull(cs_rname,'')='' Then cs_name Else cs_rname END AS cs_name,CS_ID,cs_mtel1,Fin_name,cs_register_address
                    ,(select top 1 REPLACE(REPLACE(REPLACE(REPLACE(Memo,CHAR(13), ''),CHAR(10), ''),CHAR(9), ''),CHAR(8), '') from Telemarketing_Log tl where td.TM_id=tl.TM_id and tl.TM_D_id=td.TM_D_id order by add_date desc) as CS_remark
                    from Telemarketing_M tm
                    JOIN view_Telemarketing_Curr td ON tm.TM_id = td.tm_id
                    JOIN view_Telemarketing_source ha ON tm.HA_id = ha.HA_id AND ha.TM_type = '2'
                    WHERE TelAsk = 'ASK01' and ISNULL(InDate,'') <> ''";
                if (Fin_type == "FIN_T00")
                {
                    T_SQL += " and td.Fin_type IN ('FIN_T02','FIN_T28')";
                }
                else
                {
                    T_SQL += " and td.Fin_type = @Fin_type";
                    parameters.Add(new SqlParameter("@Fin_type", Fin_type));
                }
                T_SQL += " and convert(datetime, convert(varchar(4),(convert(varchar(3), InDate, 126)+1911)) + SUBSTRING(InDate,4,10)) between @InDate_S and @InDate_E";
                parameters.Add(new SqlParameter("@InDate_S", InDate_S));
                parameters.Add(new SqlParameter("@InDate_E", InDate_E));
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

        #region 網路電銷來源數據分析表
        /// <summary>
        /// GetIntSourceSendCase/_Ajaxhandler.asp
        /// </summary>
        /// <param name="Sour">SOU01</param>
        /// <param name="YYYYMM_S">2024-08</param>
        /// <param name="YYYYMM_E">2024-09</param>
        [HttpGet("GetIntSourceSendCase")]
        public ActionResult<ResultClass<string>> GetIntSourceSendCase(string Sour, string YYYYMM_S, string YYYYMM_E)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = "exec GetIntSourceSendCase @YYYYMM_S,@YYYYMM_E,@Sour";
                parameters.Add(new SqlParameter("@YYYYMM_S", YYYYMM_S));
                parameters.Add(new SqlParameter("@YYYYMM_E", YYYYMM_E));
                parameters.Add(new SqlParameter("@Sour", Sour));
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

        #region 核准放款/佣金表  區查詢移除
        /// <summary>
        /// 是否具有查詢業務資格 GetCheckSalesRead/Approval_Loan_Sales.asp
        /// </summary>
        [HttpGet("GetCheckSalesRead")]
        public ActionResult<ResultClass<string>> GetCheckSalesRead()
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            try
            {
                var roleNum = HttpContext.Session.GetString("Role_num");
                if (roleNum == "1008" || roleNum == "1014" || roleNum == "1004" || roleNum == "1007" || roleNum == "1001")
                {
                    resultClass.objResult = "Y";
                }
                else
                {
                    resultClass.objResult = "N";
                }

                resultClass.ResultCode = "000";
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
        /// 提供可查詢業務資料 GetApprovalSalesList/Approval_Loan_Sales.asp
        /// </summary>
        [HttpGet("GetApprovalSalesList")]
        public ActionResult<ResultClass<string>> GetApprovalSalesList()
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var roleNum = HttpContext.Session.GetString("Role_num");
            var User_U_BC = HttpContext.Session.GetString("User_U_BC");

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"
                    select bc.item_D_name AS U_BC_name,um.U_num,um.U_name,pft.item_D_name AS U_PFT_name from User_M um
                    LEFT JOIN Item_list bc ON bc.item_M_code = 'branch_company' AND bc.item_D_code = um.U_BC AND bc.item_D_type = 'Y'
                    AND bc.show_tag = '0' AND bc.del_tag = '0'
                    LEFT JOIN Item_list pft ON pft.item_M_code = 'professional_title' AND pft.item_D_code = um.U_PFT
                    AND pft.item_D_type = 'Y' AND pft.show_tag = '0' AND pft.del_tag = '0'
                    where ISNULL(U_leave_date,'') = '' and Role_num <> '1001'";
                if (roleNum == "1008" || roleNum == "1014")
                {
                    T_SQL += " and U_BC=@U_BC";
                    parameters.Add(new SqlParameter("@U_BC", User_U_BC));
                }
                T_SQL += " ORDER BY bc.item_sort,pft.item_sort";
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
        /// 核准放款/佣金表_查詢 Approval_Loan_Sales_Query/Approval_Loan_Sales.asp
        /// </summary>
        /// <param name="model"></param>
        [HttpPost("Approval_Loan_Sales_Query")]
        public ActionResult<ResultClass<string>> Approval_Loan_Sales_Query(Approval_Loan_Sales_req model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var User_Num = HttpContext.Session.GetString("UserID");
            var roleNum = HttpContext.Session.GetString("Role_num");
            var User_U_BC = HttpContext.Session.GetString("User_U_BC");

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"
                    SELECT M.U_BC,M.U_name,case when CancelDate is null then 'N' else 'Y' end isCancel,
                    isnull(Comparison,'')Comparison,project_title,interest_rate_pass refRateI,Loan_rate refRateL,interest_rate_pass,isnull(I.Introducer_PID,'') I_PID,
                    case when H.act_perf_amt is null then 'N' else 'Y' end IsConfirm,isnull(H.Introducer_PID,'')Introducer_PID,
                    isnull(I_Count,0)I_Count,H.HS_id, M.U_BC_name,A.CS_name,
                    isnull(convert(varchar(4),(convert(varchar(4),misaligned_date,126)-1911))+'-'+convert(varchar(2),dbo.PadWithZero(month(misaligned_date)))+'-'+convert(varchar(2),dbo.PadWithZero(day(misaligned_date))),'') misaligned_date,
                    convert(varchar(4),(convert(varchar(4),Send_amount_date,126)-1911))+'-'+convert(varchar(2),dbo.PadWithZero(month(Send_amount_date)))+'-'+convert(varchar(2),dbo.PadWithZero(day(Send_amount_date)))Send_amount_date,
                    convert(varchar(4),(convert(varchar(4),get_amount_date,126)-1911))+'-'+convert(varchar(2),dbo.PadWithZero(month(get_amount_date)))+'-'+convert(varchar(2),dbo.PadWithZero(day(get_amount_date)))get_amount_date,
                    convert(varchar(4),(convert(varchar(4),H.Send_result_date, 126)-1911))+'-'+convert(varchar(2),dbo.PadWithZero(month(H.Send_result_date)))+'-'+convert(varchar(2),dbo.PadWithZero(day(H.Send_result_date)))Send_result_date,
                    H.CS_introducer,M.U_name plan_name,H.pass_amount,get_amount,
                    (SELECT item_D_name FROM Item_list WHERE item_M_code = 'fund_company' AND item_D_type='Y' AND item_D_code = H.fund_company AND show_tag='0' AND del_tag='0') AS show_fund_company,
                    (SELECT item_D_name FROM Item_list WHERE item_M_code = 'project_title' AND item_D_type='Y' AND item_D_code = House_pre_project.project_title AND show_tag='0' AND del_tag='0') AS show_project_title,
                    Loan_rate+'%' Loan_rate,interest_rate_original+'%' interest_rate_original,interest_rate_pass+'%' interest_rate_pass,
                    isnull(charge_flow,0)charge_flow,isnull(charge_agent,0)charge_agent,isnull(charge_check,0)charge_check,isnull(get_amount_final,0)get_amount_final,
                    isnull(H.subsidized_interest,0)subsidized_interest,R.Comm_Remark,I.Bank_name,I.Bank_account Bank_account
                    FROM House_sendcase H
                    LEFT JOIN House_apply A ON A.HA_id = H.HA_id AND A.del_tag='0'
                    LEFT JOIN House_pre_project ON House_pre_project.HP_project_id = H.HP_project_id AND House_pre_project.del_tag='0'
                    LEFT JOIN (select u.*,item_D_name U_BC_name from User_M u LEFT JOIN Item_list ub on ub.item_M_code='branch_company' AND ub.item_D_type='Y' AND ub.item_D_code=u.U_BC)  M ON M.U_num = A.plan_num
                    LEFT JOIN User_M Users ON house_pre_project.fin_user = Users.U_num
                    LEFT JOIN (select * from Introducer_Comm where del_tag='0') I ON replace(H.CS_introducer,';','') = I.Introducer_name
                    and case when H.Introducer_PID is null then I.Introducer_PID else H.Introducer_PID end=  I.Introducer_PID /*判斷退傭人是否已經財務確認押上PID*/
                    LEFT JOIN (select Introducer_name,count(U_ID)I_Count FROM Introducer_Comm group by Introducer_name) I_Cou ON H.CS_introducer = I_Cou.Introducer_name
                    Left join (select item_M_code,item_D_code,item_D_name Comm_Remark from Item_list where item_M_code = 'Return' AND item_D_type='Y') R on H.Comm_Remark=R.item_D_code
                    LEFT JOIN(select KeyVal,Max(LogDate)LogDate from [LogTable] group by KeyVal)L on convert(varchar,H.HS_id)=KeyVal
                    WHERE H.del_tag = '0' AND H.sendcase_handle_type='Y' AND isnull(H.Send_amount, '') <>'' AND get_amount_type= 'GTAT002' ";
                if (roleNum == "1009")
                {
                    T_SQL += " and A.plan_num =@U_num";
                    parameters.Add(new SqlParameter("@U_num", User_Num));
                }
                if (roleNum == "1008" || roleNum == "1014" || roleNum == "1010")
                {
                    T_SQL += " and M.U_BC =@U_BC";
                    parameters.Add(new SqlParameter("@U_BC", User_U_BC));
                }
                if (!string.IsNullOrEmpty(model.sales_num))
                {
                    T_SQL += " and plan_num=@plan_num";
                    parameters.Add(new SqlParameter("@plan_num", model.sales_num));
                }
                T_SQL += " AND convert(varchar(7),get_amount_date,126) =@selYear_S";
                T_SQL += " union all";
                T_SQL += " SELECT M.U_BC,M.U_name,isCancel,'' Comparison,H.show_project_title project_title,H.interest_rate_pass refRateI,NULL refRateL,H.interest_rate_pass,'NA' I_PID,";
                T_SQL += " CASE WHEN H.act_perf_amt IS NULL THEN 'N' ELSE 'Y' END IsConfirm,'NA'Introducer_PID,1 I_Count,case_id HS_id,M.U_BC_name,H.cs_name,'' misaligned_date,'NA' Send_amount_date,";
                T_SQL += " convert(varchar(4),(convert(varchar(4),get_amount_date,126)-1911))+'-'+convert(varchar(2),dbo.PadWithZero(month(get_amount_date)))+'-'+convert(varchar(2),dbo.PadWithZero(day(get_amount_date)))get_amount_date,";
                T_SQL += " 'NA' Send_result_date,'NA' CS_introducer,M.U_name plan_name,H.get_amount pass_amount,H.get_amount*(CASE WHEN isCancel='Y' THEN -1 ELSE 1 END) get_amount";
                T_SQL += " ,H.show_fund_company,H.show_project_title,'NA' Loan_rate,'NA' interest_rate_original,H.interest_rate_pass+'%' interest_rate_pass,0 charge_flow,0 charge_agent";
                T_SQL += " ,0 charge_check,0 get_amount_final,0 subsidized_interest,NULL Comm_Remark,Bank_name,Bank_account";
                T_SQL += " FROM (SELECT *,'NA' Bank_account,'NA' Bank_name,'N' isCancel,NULL misaligned_date FROM House_othercase";
                T_SQL += " WHERE (convert(varchar(7), (get_amount_date), 126)=@selYear_S) UNION ALL SELECT *,'NA' Bank_account,'NA' Bank_name,'Y' isCancel,NULL misaligned_date";
                T_SQL += " FROM House_othercase WHERE CancelDate IS NOT NULL) H";
                T_SQL += " LEFT JOIN(SELECT u.U_BC,U_num,U_name,ub.item_D_name U_BC_name,pt.item_sort,U_arrive_date FROM User_M u";
                T_SQL += " LEFT JOIN Item_list ub ON ub.item_M_code='branch_company'";
                T_SQL += " AND ub.item_D_type='Y' AND ub.item_D_code=u.U_BC";
                T_SQL += " LEFT JOIN Item_list pt ON pt.item_M_code='professional_title'";
                T_SQL += " AND pt.item_D_type='Y'AND pt.item_D_code=u.U_PFT)M ON H.plan_num=M.U_num";
                T_SQL += " WHERE H.del_tag = '0'";
                if (roleNum == "1009")
                {
                    T_SQL += " and A.plan_num =@U_num";
                    parameters.Add(new SqlParameter("@U_num", User_Num));
                }
                if (roleNum == "1008" || roleNum == "1014" || roleNum == "1010")
                {
                    T_SQL += " and M.U_BC =@U_BC";
                    parameters.Add(new SqlParameter("@U_BC", User_U_BC));
                }
                if (!string.IsNullOrEmpty(model.sales_num))
                {
                    T_SQL += " and plan_num=@plan_num";
                    parameters.Add(new SqlParameter("@plan_num", model.sales_num));
                }
                T_SQL += " and convert(varchar(7),get_amount_date,126) =@selYear_S";
                if (model.OrderByStyle == "1")
                {
                    T_SQL += " order by get_amount_date asc";
                }
                else
                {
                    T_SQL += " order by M.U_BC,M.U_name";
                }
                parameters.Add(new SqlParameter("@selYear_S", model.YYYYMM));
                #endregion

                DataTable dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
                if (dtResult.Rows.Count > 0)
                {
                    var ApprovalLoanSalesList = dtResult.AsEnumerable().Select(row => new Approval_Loan_Sales_res
                    {
                        U_BC_name = row.Field<string>("U_BC_name"),
                        Send_amount_date = row.Field<string>("Send_amount_date"),
                        CS_name = row.Field<string>("CS_name"),
                        CS_introducer = row.Field<string>("CS_introducer"),
                        Bank_name = row.Field<string>("Bank_name"),
                        Bank_account = row.Field<string>("Bank_account"),
                        plan_name = row.Field<string>("plan_name"),
                        Send_result_date = row.Field<string>("Send_result_date"),
                        pass_amount = row.Field<string>("pass_amount"),
                        get_amount_date = row.Field<string>("get_amount_date"),
                        get_amount = row.Field<int>("get_amount"),
                        show_project_title = row.Field<string>("show_fund_company") + row.Field<string>("show_project_title"),
                        Loan_rate = row.Field<string>("Loan_rate"),
                        interest_rate_original = row.Field<string>("interest_rate_original"),
                        interest_rate_pass = row.Field<string>("interest_rate_pass"),
                        charge_flow = row.Field<int>("charge_flow"),
                        charge_agent = row.Field<int>("charge_agent"),
                        charge_check = row.Field<int>("charge_check"),
                        get_amount_final = row.Field<int>("get_amount_final"),
                        subsidized_interest = row.Field<decimal>("subsidized_interest"),
                        Comm_Remark = row.Field<string>("Comm_Remark"),
                        Comparison = row.Field<string>("Comparison"),
                        isCancel = row.Field<string>("isCancel"),
                        Introducer_PID = row.Field<string>("Introducer_PID"),
                        I_Count = row.Field<int>("I_Count")
                    }).ToList();



                    resultClass.ResultCode = "000";
                    resultClass.objResult = JsonConvert.SerializeObject(ApprovalLoanSalesList);
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
        /// 核准放款/佣金表_下載 Approval_Loan_Sales_Excel/Approval_Loan_Sales.asp
        /// </summary>
        /// <param name="model"></param>
        [HttpPost("Approval_Loan_Sales_Excel")]
        public IActionResult Approval_Loan_Sales_Excel(Approval_Loan_Sales_req model)
        {
            var User_Num = HttpContext.Session.GetString("UserID");
            var roleNum = HttpContext.Session.GetString("Role_num");
            var User_U_BC = HttpContext.Session.GetString("User_U_BC");
            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"
                    SELECT M.U_BC,M.U_name,case when CancelDate is null then 'N' else 'Y' end isCancel,
                    isnull(Comparison,'')Comparison,project_title,interest_rate_pass refRateI,Loan_rate refRateL,interest_rate_pass,isnull(I.Introducer_PID,'') I_PID,
                    case when H.act_perf_amt is null then 'N' else 'Y' end IsConfirm,isnull(H.Introducer_PID,'')Introducer_PID,
                    isnull(I_Count,0)I_Count,H.HS_id, M.U_BC_name,A.CS_name,
                    isnull(convert(varchar(4),(convert(varchar(4),misaligned_date,126)-1911))+'-'+convert(varchar(2),dbo.PadWithZero(month(misaligned_date)))+'-'+convert(varchar(2),dbo.PadWithZero(day(misaligned_date))),'') misaligned_date,
                    convert(varchar(4),(convert(varchar(4),Send_amount_date,126)-1911))+'-'+convert(varchar(2),dbo.PadWithZero(month(Send_amount_date)))+'-'+convert(varchar(2),dbo.PadWithZero(day(Send_amount_date)))Send_amount_date,
                    convert(varchar(4),(convert(varchar(4),get_amount_date,126)-1911))+'-'+convert(varchar(2),dbo.PadWithZero(month(get_amount_date)))+'-'+convert(varchar(2),dbo.PadWithZero(day(get_amount_date)))get_amount_date,
                    convert(varchar(4),(convert(varchar(4),H.Send_result_date, 126)-1911))+'-'+convert(varchar(2),dbo.PadWithZero(month(H.Send_result_date)))+'-'+convert(varchar(2),dbo.PadWithZero(day(H.Send_result_date)))Send_result_date,
                    H.CS_introducer,M.U_name plan_name,H.pass_amount,get_amount,
                    (SELECT item_D_name FROM Item_list WHERE item_M_code = 'fund_company' AND item_D_type='Y' AND item_D_code = H.fund_company AND show_tag='0' AND del_tag='0') AS show_fund_company,
                    (SELECT item_D_name FROM Item_list WHERE item_M_code = 'project_title' AND item_D_type='Y' AND item_D_code = House_pre_project.project_title AND show_tag='0' AND del_tag='0') AS show_project_title,
                    Loan_rate+'%' Loan_rate,interest_rate_original+'%' interest_rate_original,interest_rate_pass+'%' interest_rate_pass,
                    isnull(charge_flow,0)charge_flow,isnull(charge_agent,0)charge_agent,isnull(charge_check,0)charge_check,isnull(get_amount_final,0)get_amount_final,
                    isnull(H.subsidized_interest,0)subsidized_interest,R.Comm_Remark,I.Bank_name,I.Bank_account Bank_account
                    FROM House_sendcase H
                    LEFT JOIN House_apply A ON A.HA_id = H.HA_id AND A.del_tag='0'
                    LEFT JOIN House_pre_project ON House_pre_project.HP_project_id = H.HP_project_id AND House_pre_project.del_tag='0'
                    LEFT JOIN (select u.*,item_D_name U_BC_name from User_M u LEFT JOIN Item_list ub on ub.item_M_code='branch_company' AND ub.item_D_type='Y' AND ub.item_D_code=u.U_BC)  M ON M.U_num = A.plan_num
                    LEFT JOIN User_M Users ON house_pre_project.fin_user = Users.U_num
                    LEFT JOIN (select * from Introducer_Comm where del_tag='0') I ON replace(H.CS_introducer,';','') = I.Introducer_name
                    and case when H.Introducer_PID is null then I.Introducer_PID else H.Introducer_PID end=  I.Introducer_PID /*判斷退傭人是否已經財務確認押上PID*/
                    LEFT JOIN (select Introducer_name,count(U_ID)I_Count FROM Introducer_Comm group by Introducer_name) I_Cou ON H.CS_introducer = I_Cou.Introducer_name
                    Left join (select item_M_code,item_D_code,item_D_name Comm_Remark from Item_list where item_M_code = 'Return' AND item_D_type='Y') R on H.Comm_Remark=R.item_D_code
                    LEFT JOIN(select KeyVal,Max(LogDate)LogDate from [LogTable] group by KeyVal)L on convert(varchar,H.HS_id)=KeyVal
                    WHERE H.del_tag = '0' AND H.sendcase_handle_type='Y' AND isnull(H.Send_amount, '') <>'' AND get_amount_type= 'GTAT002'";
                if (roleNum == "1009")
                {
                    T_SQL += " and A.plan_num =@U_num";
                    parameters.Add(new SqlParameter("@U_num", User_Num));
                }
                if (roleNum == "1008" || roleNum == "1014" || roleNum == "1010")
                {
                    T_SQL += " and M.U_BC =@U_BC";
                    parameters.Add(new SqlParameter("@U_BC", User_U_BC));
                }
                if (!string.IsNullOrEmpty(model.sales_num))
                {
                    T_SQL += " and plan_num=@plan_num";
                    parameters.Add(new SqlParameter("@plan_num", model.sales_num));
                }
                T_SQL += " AND convert(varchar(7),get_amount_date,126) =@selYear_S";
                T_SQL += " union all";
                T_SQL += " SELECT M.U_BC,M.U_name,isCancel,'' Comparison,H.show_project_title project_title,H.interest_rate_pass refRateI,NULL refRateL,H.interest_rate_pass,'NA' I_PID,";
                T_SQL += " CASE WHEN H.act_perf_amt IS NULL THEN 'N' ELSE 'Y' END IsConfirm,'NA'Introducer_PID,1 I_Count,case_id HS_id,M.U_BC_name,H.cs_name,'' misaligned_date,'NA' Send_amount_date,";
                T_SQL += " convert(varchar(4),(convert(varchar(4),get_amount_date,126)-1911))+'-'+convert(varchar(2),dbo.PadWithZero(month(get_amount_date)))+'-'+convert(varchar(2),dbo.PadWithZero(day(get_amount_date)))get_amount_date,";
                T_SQL += " 'NA' Send_result_date,'NA' CS_introducer,M.U_name plan_name,H.get_amount pass_amount,H.get_amount*(CASE WHEN isCancel='Y' THEN -1 ELSE 1 END) get_amount";
                T_SQL += " ,H.show_fund_company,H.show_project_title,'NA' Loan_rate,'NA' interest_rate_original,H.interest_rate_pass+'%' interest_rate_pass,0 charge_flow,0 charge_agent";
                T_SQL += " ,0 charge_check,0 get_amount_final,0 subsidized_interest,NULL Comm_Remark,Bank_name,Bank_account";
                T_SQL += " FROM (SELECT *,'NA' Bank_account,'NA' Bank_name,'N' isCancel,NULL misaligned_date FROM House_othercase";
                T_SQL += " WHERE (convert(varchar(7), (get_amount_date), 126)=@selYear_S) UNION ALL SELECT *,'NA' Bank_account,'NA' Bank_name,'Y' isCancel,NULL misaligned_date";
                T_SQL += " FROM House_othercase WHERE CancelDate IS NOT NULL) H";
                T_SQL += " LEFT JOIN(SELECT u.U_BC,U_num,U_name,ub.item_D_name U_BC_name,pt.item_sort,U_arrive_date FROM User_M u";
                T_SQL += " LEFT JOIN Item_list ub ON ub.item_M_code='branch_company'";
                T_SQL += " AND ub.item_D_type='Y' AND ub.item_D_code=u.U_BC";
                T_SQL += " LEFT JOIN Item_list pt ON pt.item_M_code='professional_title'";
                T_SQL += " AND pt.item_D_type='Y'AND pt.item_D_code=u.U_PFT)M ON H.plan_num=M.U_num";
                T_SQL += " WHERE H.del_tag = '0'";
                if (roleNum == "1009")
                {
                    T_SQL += " and A.plan_num =@U_num";
                    parameters.Add(new SqlParameter("@U_num", User_Num));
                }
                if (roleNum == "1008" || roleNum == "1014" || roleNum == "1010")
                {
                    T_SQL += " and M.U_BC =@U_BC";
                    parameters.Add(new SqlParameter("@U_BC", User_U_BC));
                }
                if (!string.IsNullOrEmpty(model.sales_num))
                {
                    T_SQL += " and plan_num=@plan_num";
                    parameters.Add(new SqlParameter("@plan_num", model.sales_num));
                }
                T_SQL += " and convert(varchar(7),get_amount_date,126) =@selYear_S";
                if (model.OrderByStyle == "1")
                {
                    T_SQL += " order by get_amount_date asc";
                }
                else
                {
                    T_SQL += " order by M.U_BC,M.U_name";
                }
                parameters.Add(new SqlParameter("@selYear_S", model.YYYYMM));
                #endregion
                DataTable dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
                if (dtResult.Rows.Count > 0)
                {
                    var ExcelList = dtResult.AsEnumerable().Select(row => new Approval_Loan_Sales_Excel
                    {
                        U_BC_name = row.Field<string>("U_BC_name"),
                        Send_amount_date = row.Field<string>("Send_amount_date"),
                        CS_name = row.Field<string>("CS_name"),
                        CS_introducer = row.Field<string>("CS_introducer"),
                        Bank_name = row.Field<string>("Bank_name"),
                        Bank_account = row.Field<string>("Bank_account"),
                        plan_name = row.Field<string>("plan_name"),
                        Send_result_date = row.Field<string>("Send_result_date"),
                        pass_amount = row.Field<string>("pass_amount"),
                        get_amount_date = row.Field<string>("get_amount_date"),
                        get_amount = row.Field<int>("get_amount"),
                        show_project_title = row.Field<string>("show_fund_company") + row.Field<string>("show_project_title"),
                        Loan_rate = row.Field<string>("Loan_rate"),
                        interest_rate_original = row.Field<string>("interest_rate_original"),
                        interest_rate_pass = row.Field<string>("interest_rate_pass"),
                        charge_flow = row.Field<int>("charge_flow"),
                        charge_agent = row.Field<int>("charge_agent"),
                        charge_check = row.Field<int>("charge_check"),
                        get_amount_final = row.Field<int>("get_amount_final"),
                        subsidized_interest = row.Field<decimal>("subsidized_interest"),
                        Comm_Remark = row.Field<string>("Comm_Remark"),
                        Comparison = row.Field<string>("Comparison")
                    }).ToList();
                    var Excel_Headers = new Dictionary<string, string>
                    {
                        { "U_BC_name", "區" },
                        { "Send_amount_date", "進件日" },
                        { "CS_name", "申請人" },
                        { "CS_introducer", "介紹人" },
                        { "Bank_name", "銀行名稱" },
                        { "Bank_account", "銀行帳號" },
                        { "plan_name", "業務" },
                        { "Send_result_date", "核准日" },
                        { "pass_amount", "核准金額(萬)" },
                        { "get_amount_date", "撥款日" },
                        { "get_amount", "撥款金額(萬)" },
                        { "show_project_title", "專案" },
                        { "Loan_rate", "貸款成數(%)" },
                        { "interest_rate_original", "原始利率(%)" },
                        { "interest_rate_pass", "承作利率(%)" },
                        { "charge_flow", "收費" },
                        { "charge_agent", "代書費用" },
                        { "charge_check", "對保費" },
                        { "get_amount_final", "結餘" },
                        { "subsidized_interest", "補貼息" },
                        { "Comm_Remark", "退傭備註" },
                        { "Comparison", "委對/外對" }
                    };
                    var fileBytes = FuncHandler.ExportToExcel(ExcelList, Excel_Headers);
                    fileBytes = FuncHandler.ApprovalLoanSalesExcelFooter(fileBytes);
                    var fileName = "核准放款表_傭金表" + DateTime.Now.ToString("yyyyMMddHHmm") + ".xlsx";
                    return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
                }
                else
                {
                    return NotFound(); // 檔案不存在時返回 404
                }
            }
            catch (Exception ex)
            {
                ResultClass<string> resultClass = new ResultClass<string>();
                resultClass.ResultCode = "500";
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass);
            }
        }
        #endregion

        #region 請假單報表
        /// <summary>
        /// 請假單報表_查詢 Flow_rest_report_query/Flow_rest_report.asp
        /// </summary>
        [HttpPost("Flow_Rest_Report_Query")]
        public ActionResult<ResultClass<string>> Flow_Rest_Report_Query(Flow_rest_report_req model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var User_Num = HttpContext.Session.GetString("UserID");
            var roleNum = HttpContext.Session.GetString("Role_num");

            try
            {
                //特殊權限判定
                string[] str = new string[] { "7005", "7007" };
                SpecialClass specialClass = FuncHandler.CheckSpecial(str, User_Num);
                if (specialClass.special_check == "N")
                    model.U_num = User_Num;

                ADOData _adoData = new ADOData();
                #region SQL_早退資料
                var parameters_ey = new List<SqlParameter>();
                var T_SQL_EY = @"
                    select userID,sum(early)early from (
                    SELECT [userID],cast(year(cast(yyyymm +'01' as datetime)) as varchar(4)) +'/'+ad.attendance_date　attendance_date,
                    CASE WHEN isnull([getoffwork_time], '')='' THEN 0 WHEN [getoffwork_time]<'18:00' THEN DATEDIFF(MINUTE, [getoffwork_time], '18:00') ELSE 0 END early
                    FROM [dbo].[attendance] ad
                    LEFT JOIN (SELECT U_BC,U_num,Role_num FROM [dbo].[User_M] U WHERE 1=1  ) U ON ad.userID=U.U_num
                    LEFT JOIN (SELECT FR_U_num,convert(varchar, FR_date_begin, 111) FR_date_S,convert(varchar, FR_date_end, 111) FR_date_E
                    ,count(FR_U_num) RestCount
                    FROM Flow_rest WHERE del_tag = '0' AND FR_cancel<>'Y' GROUP BY FR_U_num,convert(varchar, FR_date_begin, 111),
                    convert(varchar, FR_date_end, 111)) R ON ad.userID=R.FR_U_num
                    AND @YYYY+'/'+ad.[attendance_date] BETWEEN R.FR_date_S AND FR_date_E
                    WHERE convert(varchar,convert(datetime,  @YYYY+'/'+[attendance_date]), 111)
                    not in (SELECT convert(varchar,convert(datetime, [HDate]), 111) FROM [dbo].[Holidays])
                    AND [getoffwork_time]<'18:00' and [userID] <> '' and R.FR_date_S is null and Role_num in (SELECT R_num FROM Role_M where LE_tag='Y')) E
                    where CAST(attendance_date AS datetime) between @yyyymmdd_s and @yyyymmdd_e group by userID";
                parameters_ey.Add(new SqlParameter("@YYYY", FuncHandler.ConvertROCToGregorian(model.Flow_Rest_Date_S).Split('/')[0]));
                parameters_ey.Add(new SqlParameter("@yyyymmdd_s", FuncHandler.ConvertROCToGregorian(model.Flow_Rest_Date_S)));
                parameters_ey.Add(new SqlParameter("@yyyymmdd_e", FuncHandler.ConvertROCToGregorian(model.Flow_Rest_Date_E)));
                #endregion
                DataTable dtResult_Ey = _adoData.ExecuteQuery(T_SQL_EY, parameters_ey);
                var EarlyList = dtResult_Ey.AsEnumerable().Select(row => new Flow_rest_report_res_early
                {
                    U_num = row.Field<string>("userID"),
                    Sum_early = row.Field<int>("early")
                }).ToList();

                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"
                    select FR_U_num,U_BC_name,isnull(LE_tag,'N')LE_tag,isLeave,(select U_name FROM User_M where U_num=x.FR_U_num) as FR_U_name
                    ,sum(FR_kind_FRK001) as '事假',sum(FR_kind_FRK002) as '病假',sum(FR_kind_FRK003) as '公假',sum(FR_kind_FRK004) as '補休'
                    ,sum(FR_kind_FRK005) as '特休',sum(FR_kind_FRK006) as '婚假',sum(FR_kind_FRK007) as '喪假',sum(FR_kind_FRK008) as '產假'
                    ,sum(FR_kind_FRK009) as '陪產假',sum(FR_kind_FRK010) as '產檢假',sum(FR_kind_FRK011) as '家庭照顧假',sum(FR_kind_FRK012) as '生理假'
                    ,sum(FR_kind_FRK013) as '公傷假',sum(FR_kind_FRK014) as '疫苗假',sum(FR_kind_FRＫ015) as '防疫照顧假',sum(FR_kind_FRK018) as '居家上班'
                    ,sum(FR_kind_FRK016) as '外出(公出)',sum(FR_kind_FRK017) as '忘打卡',sum(FR_kind_FRK019) as '育嬰假',sum(FR_kind_FRK020) as '陪產檢假'
                    ,sum(FR_kind_FRK999) as '早退' from (select *,(case when Flow_rest.FR_kind='FRK001' then FR_total_hour else 0 end) as 'FR_kind_FRK001'
                    ,(case when Flow_rest.FR_kind='FRK002' then FR_total_hour else 0 end) as 'FR_kind_FRK002'
                    ,(case when Flow_rest.FR_kind='FRK003' then FR_total_hour else 0 end) as 'FR_kind_FRK003'
                    ,(case when Flow_rest.FR_kind='FRK004' then FR_total_hour else 0 end) as 'FR_kind_FRK004'
                    ,(case when Flow_rest.FR_kind='FRK005' then FR_total_hour else 0 end) as 'FR_kind_FRK005'
                    ,(case when Flow_rest.FR_kind='FRK006' then FR_total_hour else 0 end) as 'FR_kind_FRK006'
                    ,(case when Flow_rest.FR_kind='FRK007' then FR_total_hour else 0 end) as 'FR_kind_FRK007'
                    ,(case when Flow_rest.FR_kind='FRK008' then FR_total_hour else 0 end) as 'FR_kind_FRK008'
                    ,(case when Flow_rest.FR_kind='FRK009' then FR_total_hour else 0 end) as 'FR_kind_FRK009'
                    ,(case when Flow_rest.FR_kind='FRK010' then FR_total_hour else 0 end) as 'FR_kind_FRK010'
                    ,(case when Flow_rest.FR_kind='FRK011' then FR_total_hour else 0 end) as 'FR_kind_FRK011'
                    ,(case when Flow_rest.FR_kind='FRK012' then FR_total_hour else 0 end) as 'FR_kind_FRK012'
                    ,(case when Flow_rest.FR_kind='FRK013' then FR_total_hour else 0 end) as 'FR_kind_FRK013'
                    ,(case when Flow_rest.FR_kind='FRK014' then FR_total_hour else 0 end) as 'FR_kind_FRK014'
                    ,(case when Flow_rest.FR_kind='FRK015' then FR_total_hour else 0 end) as 'FR_kind_FRK015'
                    ,(case when Flow_rest.FR_kind='FRK018' then FR_total_hour else 0 end) as 'FR_kind_FRK018'
                    ,(case when Flow_rest.FR_kind='FRK016' then FR_total_hour else 0 end) as 'FR_kind_FRK016'
                    ,(case when Flow_rest.FR_kind='FRK017' then FR_total_hour else 0 end) as 'FR_kind_FRK017'
                    ,(case when Flow_rest.FR_kind='FRK019' then FR_total_hour else 0 end) as 'FR_kind_FRK019'
                    ,(case when Flow_rest.FR_kind='FRK020' then FR_total_hour else 0 end) as 'FR_kind_FRK020'
                    ,(case when Flow_rest.FR_kind='FRK999' then FR_total_hour else 0 end) as 'FR_kind_FRK999'
                    ,(select item_D_name from Item_list where item_M_code='branch_company' AND item_D_type='Y' AND item_D_code=User_M.U_BC AND del_tag='0') as U_BC_name
                    ,(select item_D_code from Item_list where item_M_code='branch_company' AND item_D_type='Y' AND item_D_code=User_M.U_BC AND del_tag='0') as U_BC_code
                    ,(select item_sort from Item_list where item_M_code='branch_company' AND item_D_type='Y' AND item_D_code=User_M.U_BC AND del_tag='0') as U_BC_sort
                    from Flow_rest
                    left join (select U_num,U_BC,LE_tag,case when M.U_leave_date is null then 'N' else 'Y' end isLeave FROM User_M M
                    left join Role_M R on M.Role_num=R.R_num) User_M ON User_M.U_num = Flow_rest.FR_U_num
                    where del_tag = '0' AND FR_sign_type = 'FSIGN002'";
                if (!string.IsNullOrEmpty(model.U_num))
                {
                    T_SQL += " and FR_U_num = @U_num";
                    parameters.Add(new SqlParameter("@U_num", model.U_num));
                }
                if (!string.IsNullOrEmpty(model.LE_tag))
                {
                    T_SQL += " and isnull(LE_tag,'N')='Y'";
                }
                T_SQL += " AND ((FR_date_begin >= @yyyymmdd_s+' 00:00:00' AND FR_date_begin <= @yyyymmdd_e+' 23:59:59' )";
                T_SQL += " OR (FR_date_end >= @yyyymmdd_s+' 00:00:00' AND FR_date_end <= @yyyymmdd_e+' 23:59:59' ) ";
                T_SQL += " OR (FR_date_begin <= @yyyymmdd_s+' 00:00:00' AND FR_date_end >= @yyyymmdd_e+' 23:59:59'))) x ";
                T_SQL += " where isleave='N'";
                if (!string.IsNullOrEmpty(model.U_BC))
                {
                    T_SQL += " and x.U_BC_code = @U_BC";
                    parameters.Add(new SqlParameter("@U_BC", model.U_BC));
                }
                T_SQL += " group by FR_U_num,LE_tag,isLeave,U_BC_name,U_BC_sort order by U_BC_sort,FR_U_num";
                parameters.Add(new SqlParameter("@yyyymmdd_s", FuncHandler.ConvertROCToGregorian(model.Flow_Rest_Date_S)));
                parameters.Add(new SqlParameter("@yyyymmdd_e", FuncHandler.ConvertROCToGregorian(model.Flow_Rest_Date_E)));
                #endregion
                var dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
                var Restlist = dtResult.AsEnumerable().Select(row => new Flow_rest_report_Excel
                {
                    U_BC_name = row.Field<string>("U_BC_name"),
                    FR_U_num = row.Field<string>("FR_U_num"),
                    FR_U_name = row.Field<string>("U_BC_name"),
                    SUM_FR_kind_FRK001 = row.Field<decimal>("事假"),
                    SUM_FR_kind_FRK002 = row.Field<decimal>("病假"),
                    SUM_FR_kind_FRK003 = row.Field<decimal>("公假"),
                    SUM_FR_kind_FRK004 = row.Field<decimal>("補休"),
                    SUM_FR_kind_FRK005 = row.Field<decimal>("特休"),
                    SUM_FR_kind_FRK006 = row.Field<decimal>("婚假"),
                    SUM_FR_kind_FRK007 = row.Field<decimal>("喪假"),
                    SUM_FR_kind_FRK008 = row.Field<decimal>("產假"),
                    SUM_FR_kind_FRK009 = row.Field<decimal>("陪產假"),
                    SUM_FR_kind_FRK010 = row.Field<decimal>("產檢假"),
                    SUM_FR_kind_FRK011 = row.Field<decimal>("家庭照顧假"),
                    SUM_FR_kind_FRK012 = row.Field<decimal>("生理假"),
                    SUM_FR_kind_FRK013 = row.Field<decimal>("公傷假"),
                    SUM_FR_kind_FRK014 = row.Field<decimal>("疫苗假"),
                    SUM_FR_kind_FRK015 = row.Field<decimal>("防疫照顧假"),
                    SUM_FR_kind_FRK018 = row.Field<decimal>("居家上班"),
                    SUM_FR_kind_FRK016 = row.Field<decimal>("外出(公出)"),
                    SUM_FR_kind_FRK017 = row.Field<decimal>("忘打卡"),
                    SUM_FR_kind_FRK019 = row.Field<decimal>("育嬰假"),
                    SUM_FR_kind_FRK020 = row.Field<decimal>("陪產檢假"),
                    SUM_FR_kind_FRK999 = row.Field<decimal>("早退")
                }).ToList();

                var mergedList = Restlist.Select(a => new {
                    a.U_BC_name,
                    a.FR_U_num,
                    a.FR_U_name,
                    a.SUM_FR_kind_FRK001,
                    a.SUM_FR_kind_FRK002,
                    a.SUM_FR_kind_FRK003,
                    a.SUM_FR_kind_FRK004,
                    a.SUM_FR_kind_FRK005,
                    a.SUM_FR_kind_FRK006,
                    a.SUM_FR_kind_FRK007,
                    a.SUM_FR_kind_FRK008,
                    a.SUM_FR_kind_FRK009,
                    a.SUM_FR_kind_FRK010,
                    a.SUM_FR_kind_FRK011,
                    a.SUM_FR_kind_FRK012,
                    a.SUM_FR_kind_FRK013,
                    a.SUM_FR_kind_FRK014,
                    a.SUM_FR_kind_FRK015,
                    a.SUM_FR_kind_FRK018,
                    a.SUM_FR_kind_FRK016,
                    a.SUM_FR_kind_FRK017,
                    a.SUM_FR_kind_FRK019,
                    a.SUM_FR_kind_FRK020,
                    SUM_FR_kind_FRK999 = EarlyList.FirstOrDefault(b => b.U_num == a.FR_U_num)?.Sum_early
                }).ToList();

                resultClass.ResultCode = "000";
                resultClass.objResult = JsonConvert.SerializeObject(mergedList);
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
        /// 請假單報表_下載 Flow_rest_report_query/Flow_rest_report.asp
        /// </summary>
        [HttpPost("Flow_Rest_Report_Excel")]
        public IActionResult Flow_Rest_Report_Excel(Flow_rest_report_req model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var User_Num = HttpContext.Session.GetString("UserID");
            var roleNum = HttpContext.Session.GetString("Role_num");

            try
            {
                //特殊權限判定
                string[] str = new string[] { "7005", "7007" };
                SpecialClass specialClass = FuncHandler.CheckSpecial(str, User_Num);
                if (specialClass.special_check == "N")
                    model.U_num = User_Num;

                ADOData _adoData = new ADOData();
                #region SQL_早退資料
                var parameters_ey = new List<SqlParameter>();
                var T_SQL_EY = @"
                    select userID,sum(early)early from (
                    SELECT [userID],cast(year(cast(yyyymm +'01' as datetime)) as varchar(4)) +'/'+ad.attendance_date　attendance_date,
                    CASE WHEN isnull([getoffwork_time], '')='' THEN 0 WHEN [getoffwork_time]<'18:00' THEN DATEDIFF(MINUTE, [getoffwork_time], '18:00') ELSE 0 END early
                    FROM [dbo].[attendance] ad
                    LEFT JOIN (SELECT U_BC,U_num,Role_num FROM [dbo].[User_M] U WHERE 1=1  ) U ON ad.userID=U.U_num
                    LEFT JOIN (SELECT FR_U_num,convert(varchar, FR_date_begin, 111) FR_date_S,convert(varchar, FR_date_end, 111) FR_date_E
                    ,count(FR_U_num) RestCount
                    FROM Flow_rest WHERE del_tag = '0' AND FR_cancel<>'Y' GROUP BY FR_U_num,convert(varchar, FR_date_begin, 111),
                    convert(varchar, FR_date_end, 111)) R ON ad.userID=R.FR_U_num
                    AND @YYYY+'/'+ad.[attendance_date] BETWEEN R.FR_date_S AND FR_date_E
                    WHERE convert(varchar,convert(datetime,  @YYYY+'/'+[attendance_date]), 111)
                    not in (SELECT convert(varchar,convert(datetime, [HDate]), 111) FROM [dbo].[Holidays])
                    AND [getoffwork_time]<'18:00' and [userID] <> '' and R.FR_date_S is null and Role_num in (SELECT R_num FROM Role_M where LE_tag='Y')) E
                    where CAST(attendance_date AS datetime) between @yyyymmdd_s and @yyyymmdd_e group by userID";
                parameters_ey.Add(new SqlParameter("@YYYY", FuncHandler.ConvertROCToGregorian(model.Flow_Rest_Date_S).Split('/')[0]));
                parameters_ey.Add(new SqlParameter("@yyyymmdd_s", FuncHandler.ConvertROCToGregorian(model.Flow_Rest_Date_S)));
                parameters_ey.Add(new SqlParameter("@yyyymmdd_e", FuncHandler.ConvertROCToGregorian(model.Flow_Rest_Date_E)));
                #endregion
                DataTable dtResult_Ey = _adoData.ExecuteQuery(T_SQL_EY, parameters_ey);
                var EarlyList = dtResult_Ey.AsEnumerable().Select(row => new Flow_rest_report_res_early
                {
                    U_num = row.Field<string>("userID"),
                    Sum_early = row.Field<int>("early")
                }).ToList();

                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"
                    select FR_U_num,U_BC_name,isnull(LE_tag,'N')LE_tag,isLeave,(select U_name FROM User_M where U_num=x.FR_U_num) as FR_U_name
                    ,sum(FR_kind_FRK001) as '事假',sum(FR_kind_FRK002) as '病假',sum(FR_kind_FRK003) as '公假',sum(FR_kind_FRK004) as '補休'
                    ,sum(FR_kind_FRK005) as '特休',sum(FR_kind_FRK006) as '婚假',sum(FR_kind_FRK007) as '喪假',sum(FR_kind_FRK008) as '產假'
                    ,sum(FR_kind_FRK009) as '陪產假',sum(FR_kind_FRK010) as '產檢假',sum(FR_kind_FRK011) as '家庭照顧假',sum(FR_kind_FRK012) as '生理假'
                    ,sum(FR_kind_FRK013) as '公傷假',sum(FR_kind_FRK014) as '疫苗假',sum(FR_kind_FRＫ015) as '防疫照顧假',sum(FR_kind_FRK018) as '居家上班'
                    ,sum(FR_kind_FRK016) as '外出(公出)',sum(FR_kind_FRK017) as '忘打卡',sum(FR_kind_FRK019) as '育嬰假',sum(FR_kind_FRK020) as '陪產檢假'
                    ,sum(FR_kind_FRK999) as '早退' from (select *,(case when Flow_rest.FR_kind='FRK001' then FR_total_hour else 0 end) as 'FR_kind_FRK001'
                    ,(case when Flow_rest.FR_kind='FRK002' then FR_total_hour else 0 end) as 'FR_kind_FRK002'
                    ,(case when Flow_rest.FR_kind='FRK003' then FR_total_hour else 0 end) as 'FR_kind_FRK003'
                    ,(case when Flow_rest.FR_kind='FRK004' then FR_total_hour else 0 end) as 'FR_kind_FRK004'
                    ,(case when Flow_rest.FR_kind='FRK005' then FR_total_hour else 0 end) as 'FR_kind_FRK005'
                    ,(case when Flow_rest.FR_kind='FRK006' then FR_total_hour else 0 end) as 'FR_kind_FRK006'
                    ,(case when Flow_rest.FR_kind='FRK007' then FR_total_hour else 0 end) as 'FR_kind_FRK007'
                    ,(case when Flow_rest.FR_kind='FRK008' then FR_total_hour else 0 end) as 'FR_kind_FRK008'
                    ,(case when Flow_rest.FR_kind='FRK009' then FR_total_hour else 0 end) as 'FR_kind_FRK009'
                    ,(case when Flow_rest.FR_kind='FRK010' then FR_total_hour else 0 end) as 'FR_kind_FRK010'
                    ,(case when Flow_rest.FR_kind='FRK011' then FR_total_hour else 0 end) as 'FR_kind_FRK011'
                    ,(case when Flow_rest.FR_kind='FRK012' then FR_total_hour else 0 end) as 'FR_kind_FRK012'
                    ,(case when Flow_rest.FR_kind='FRK013' then FR_total_hour else 0 end) as 'FR_kind_FRK013'
                    ,(case when Flow_rest.FR_kind='FRK014' then FR_total_hour else 0 end) as 'FR_kind_FRK014'
                    ,(case when Flow_rest.FR_kind='FRK015' then FR_total_hour else 0 end) as 'FR_kind_FRK015'
                    ,(case when Flow_rest.FR_kind='FRK018' then FR_total_hour else 0 end) as 'FR_kind_FRK018'
                    ,(case when Flow_rest.FR_kind='FRK016' then FR_total_hour else 0 end) as 'FR_kind_FRK016'
                    ,(case when Flow_rest.FR_kind='FRK017' then FR_total_hour else 0 end) as 'FR_kind_FRK017'
                    ,(case when Flow_rest.FR_kind='FRK019' then FR_total_hour else 0 end) as 'FR_kind_FRK019'
                    ,(case when Flow_rest.FR_kind='FRK020' then FR_total_hour else 0 end) as 'FR_kind_FRK020'
                    ,(case when Flow_rest.FR_kind='FRK999' then FR_total_hour else 0 end) as 'FR_kind_FRK999'
                    ,(select item_D_name from Item_list where item_M_code='branch_company' AND item_D_type='Y' AND item_D_code=User_M.U_BC AND del_tag='0') as U_BC_name
                    ,(select item_D_code from Item_list where item_M_code='branch_company' AND item_D_type='Y' AND item_D_code=User_M.U_BC AND del_tag='0') as U_BC_code
                    ,(select item_sort from Item_list where item_M_code='branch_company' AND item_D_type='Y' AND item_D_code=User_M.U_BC AND del_tag='0') as U_BC_sort
                    from Flow_rest
                    left join (select U_num,U_BC,LE_tag,case when M.U_leave_date is null then 'N' else 'Y' end isLeave FROM User_M M
                    left join Role_M R on M.Role_num=R.R_num) User_M ON User_M.U_num = Flow_rest.FR_U_num
                    where del_tag = '0' AND FR_sign_type = 'FSIGN002'";
                if (!string.IsNullOrEmpty(model.U_num))
                {
                    T_SQL += " and FR_U_num = @U_num";
                    parameters.Add(new SqlParameter("@U_num", model.U_num));
                }
                if (!string.IsNullOrEmpty(model.LE_tag))
                {
                    T_SQL += " and isnull(LE_tag,'N')='Y'";
                }
                T_SQL += " AND ((FR_date_begin >= @yyyymmdd_s+' 00:00:00' AND FR_date_begin <= @yyyymmdd_e+' 23:59:59' )";
                T_SQL += " OR (FR_date_end >= @yyyymmdd_s+' 00:00:00' AND FR_date_end <= @yyyymmdd_e+' 23:59:59' ) ";
                T_SQL += " OR (FR_date_begin <= @yyyymmdd_s+' 00:00:00' AND FR_date_end >= @yyyymmdd_e+' 23:59:59'))) x ";
                T_SQL += " where isleave='N'";
                if (!string.IsNullOrEmpty(model.U_BC))
                {
                    T_SQL += " and x.U_BC_code = @U_BC";
                    parameters.Add(new SqlParameter("@U_BC", model.U_BC));
                }
                T_SQL += " group by FR_U_num,LE_tag,isLeave,U_BC_name,U_BC_sort order by U_BC_sort,FR_U_num";
                parameters.Add(new SqlParameter("@yyyymmdd_s", FuncHandler.ConvertROCToGregorian(model.Flow_Rest_Date_S)));
                parameters.Add(new SqlParameter("@yyyymmdd_e", FuncHandler.ConvertROCToGregorian(model.Flow_Rest_Date_E)));
                #endregion
                var dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
                var Restlist = dtResult.AsEnumerable().Select(row => new Flow_rest_report_Excel
                {
                    U_BC_name = row.Field<string>("U_BC_name"),
                    FR_U_num = row.Field<string>("FR_U_num"),
                    FR_U_name = row.Field<string>("U_BC_name"),
                    SUM_FR_kind_FRK001 = row.Field<decimal>("事假"),
                    SUM_FR_kind_FRK002 = row.Field<decimal>("病假"),
                    SUM_FR_kind_FRK003 = row.Field<decimal>("公假"),
                    SUM_FR_kind_FRK004 = row.Field<decimal>("補休"),
                    SUM_FR_kind_FRK005 = row.Field<decimal>("特休"),
                    SUM_FR_kind_FRK006 = row.Field<decimal>("婚假"),
                    SUM_FR_kind_FRK007 = row.Field<decimal>("喪假"),
                    SUM_FR_kind_FRK008 = row.Field<decimal>("產假"),
                    SUM_FR_kind_FRK009 = row.Field<decimal>("陪產假"),
                    SUM_FR_kind_FRK010 = row.Field<decimal>("產檢假"),
                    SUM_FR_kind_FRK011 = row.Field<decimal>("家庭照顧假"),
                    SUM_FR_kind_FRK012 = row.Field<decimal>("生理假"),
                    SUM_FR_kind_FRK013 = row.Field<decimal>("公傷假"),
                    SUM_FR_kind_FRK014 = row.Field<decimal>("疫苗假"),
                    SUM_FR_kind_FRK015 = row.Field<decimal>("防疫照顧假"),
                    SUM_FR_kind_FRK018 = row.Field<decimal>("居家上班"),
                    SUM_FR_kind_FRK016 = row.Field<decimal>("外出(公出)"),
                    SUM_FR_kind_FRK017 = row.Field<decimal>("忘打卡"),
                    SUM_FR_kind_FRK019 = row.Field<decimal>("育嬰假"),
                    SUM_FR_kind_FRK020 = row.Field<decimal>("陪產檢假"),
                    SUM_FR_kind_FRK999 = row.Field<decimal>("早退")
                }).ToList();

                var ExcelList = Restlist.Select(a => new {
                    a.U_BC_name,
                    a.FR_U_num,
                    a.FR_U_name,
                    a.SUM_FR_kind_FRK001,
                    a.SUM_FR_kind_FRK002,
                    a.SUM_FR_kind_FRK003,
                    a.SUM_FR_kind_FRK004,
                    a.SUM_FR_kind_FRK005,
                    a.SUM_FR_kind_FRK006,
                    a.SUM_FR_kind_FRK007,
                    a.SUM_FR_kind_FRK008,
                    a.SUM_FR_kind_FRK009,
                    a.SUM_FR_kind_FRK010,
                    a.SUM_FR_kind_FRK011,
                    a.SUM_FR_kind_FRK012,
                    a.SUM_FR_kind_FRK013,
                    a.SUM_FR_kind_FRK014,
                    a.SUM_FR_kind_FRK015,
                    a.SUM_FR_kind_FRK018,
                    a.SUM_FR_kind_FRK016,
                    a.SUM_FR_kind_FRK017,
                    a.SUM_FR_kind_FRK019,
                    a.SUM_FR_kind_FRK020,
                    SUM_FR_kind_FRK999 = EarlyList.FirstOrDefault(b => b.U_num == a.FR_U_num)?.Sum_early
                }).ToList();
                var Excel_Headers = new Dictionary<string, string>
                    {
                        { "U_BC_name", "公司別" },
                        { "FR_U_num", "員編" },
                        { "FR_U_name", "姓名" },
                        { "SUM_FR_kind_FRK001", "事假" },
                        { "SUM_FR_kind_FRK002", "病假" },
                        { "SUM_FR_kind_FRK003", "公假" },
                        { "SUM_FR_kind_FRK004", "補休" },
                        { "SUM_FR_kind_FRK005", "特休" },
                        { "SUM_FR_kind_FRK006", "婚假" },
                        { "SUM_FR_kind_FRK007", "喪假" },
                        { "SUM_FR_kind_FRK008", "產假" },
                        { "SUM_FR_kind_FRK009", "陪產假" },
                        { "SUM_FR_kind_FRK010", "產檢假" },
                        { "SUM_FR_kind_FRK011", "家庭照顧假" },
                        { "SUM_FR_kind_FRK012", "生理假" },
                        { "SUM_FR_kind_FRK013", "公傷假" },
                        { "SUM_FR_kind_FRK014", "疫苗假" },
                        { "SUM_FR_kind_FRK015", "防疫照顧假" },
                        { "SUM_FR_kind_FRK018", "居家上班" },
                        { "SUM_FR_kind_FRK016", "外出(公出)" },
                        { "SUM_FR_kind_FRK017", "忘打卡" },
                        { "SUM_FR_kind_FRK019", "育嬰假" },
                        { "SUM_FR_kind_FRK020", "陪產檢假" },
                        { "SUM_FR_kind_FRK999", "早退(分)" }
                    };

                var fileBytes = FuncHandler.ExportToExcel(ExcelList, Excel_Headers);
                var fileName = "請假單報表" + DateTime.Now.ToString("yyyyMMddHHmm") + ".xlsx";
                return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass);
            }

        }
        #endregion

        #region 業務件數統計表
        /// <summary>
        /// 顯示區查詢權限 GetAreaShowType/customer_qty_count.asp
        /// </summary>
        [HttpGet("GetAreaShowType")]
        public ActionResult<ResultClass<string>> GetAreaShowType()
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var User_Num = HttpContext.Session.GetString("UserID");
            var roleNum = HttpContext.Session.GetString("Role_num");
            SpecialClass specialClass = new SpecialClass();

            try
            {
                string[] strarr = new string[] { "1006", "1007", "1001", "1005" };
                if (strarr.Contains(roleNum))
                {
                    specialClass.special_check = "Y";
                    specialClass.BC_Strings = "xxx,BC0100,BC0200,BC0600,BC0300,BC0500,BC0400,BC0900";
                }
                else
                {
                    specialClass.special_check = "N";
                    specialClass.BC_Strings = "";
                }
                resultClass.ResultCode = "000";
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
        /// 業務件數_查詢 Customer_qty_count_Query/customer_qty_count.asp
        /// </summary>
        [HttpPost("Customer_qty_count_Query")]
        public ActionResult<ResultClass<string>> Customer_qty_count_Query(customer_qty_count_req model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var roleNum = HttpContext.Session.GetString("Role_num");
            var User_U_BC = HttpContext.Session.GetString("User_U_BC");

            try
            {
                var Qty_Date_S = FuncHandler.ConvertROCToGregorian(model.Qty_Date_S);
                var Qty_Date_E = FuncHandler.ConvertROCToGregorian(model.Qty_Date_E);
                //特殊權限判定
                var sql_UDBC = "";
                var sql_UBC = "";
                string[] strarr = new string[] { "1006", "1007", "1001", "1005" };
                if (strarr.Contains(roleNum))
                {
                    sql_UBC = "xxx, BC0100, BC0200, BC0600, BC0300, BC0500, BC0400";
                    sql_UDBC = "BC0900";
                }
                else
                {
                    sql_UBC = User_U_BC;
                }
                if (!string.IsNullOrEmpty(model.U_BC))
                {
                    sql_UBC = model.U_BC;
                    sql_UDBC = "";
                }
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"
                    select * from fun_customer_count(@Qty_Date_S,@Qty_Date_E, @sql_UBC,@sql_UDBC)
                    ORDER BY case when u_bc = 'BC0900' then 'BC0101' WHEN u_bc = 'BC0600' THEN 'BC0201' WHEN u_bc = 'BC0500' THEN 'BC0301' else u_bc end,item_sort";
                parameters.Add(new SqlParameter("@Qty_Date_S", Qty_Date_S));
                parameters.Add(new SqlParameter("@Qty_Date_E", Qty_Date_E));
                parameters.Add(new SqlParameter("@sql_UBC", sql_UBC));
                parameters.Add(new SqlParameter("@sql_UDBC", sql_UDBC));
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
        /// 業務件數_下載 Customer_qty_count_Excel/customer_qty_count.asp
        /// </summary>
        [HttpPost("Customer_qty_count_Excel")]
        public IActionResult Customer_qty_count_Excel(customer_qty_count_req model)
        {
            var roleNum = HttpContext.Session.GetString("Role_num");
            var User_U_BC = HttpContext.Session.GetString("User_U_BC");

            try
            {
                var Qty_Date_S = FuncHandler.ConvertROCToGregorian(model.Qty_Date_S);
                var Qty_Date_E = FuncHandler.ConvertROCToGregorian(model.Qty_Date_E);
                //特殊權限判定
                var sql_UDBC = "";
                var sql_UBC = "";
                string[] strarr = new string[] { "1006", "1007", "1001", "1005" };
                if (strarr.Contains(roleNum))
                {
                    sql_UBC = "xxx, BC0100, BC0200, BC0600, BC0300, BC0500, BC0400";
                    sql_UDBC = "BC0900";
                }
                else
                {
                    sql_UBC = User_U_BC;
                }
                if (!string.IsNullOrEmpty(model.U_BC))
                {
                    sql_UBC = model.U_BC;
                    sql_UDBC = "";
                }
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"
                    select * from fun_customer_count(@Qty_Date_S,@Qty_Date_E, @sql_UBC,@sql_UDBC)
                    ORDER BY case when u_bc = 'BC0900' then 'BC0101' WHEN u_bc = 'BC0600' THEN 'BC0201' WHEN u_bc = 'BC0500' THEN 'BC0301' else u_bc end,item_sort";
                parameters.Add(new SqlParameter("@Qty_Date_S", Qty_Date_S));
                parameters.Add(new SqlParameter("@Qty_Date_E", Qty_Date_E));
                parameters.Add(new SqlParameter("@sql_UBC", sql_UBC));
                parameters.Add(new SqlParameter("@sql_UDBC", sql_UDBC));
                #endregion
                DataTable dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
                var ExcelList = dtResult.AsEnumerable().Select(row => new customer_qty_count_Excel {
                    com_name = row.Field<string>("com_name"),
                    title_name = row.Field<string>("title_name"),
                    U_name = row.Field<string>("U_name"),
                    count = row.Field<int>("_count")
                }).ToList();
                var Excel_Headers = new Dictionary<string, string>
                {
                    { "com_name", "公司別" },
                    { "title_name", "職稱" },
                    { "U_name", "業務" },
                    { "count", "預估件數" }
                };

                var fileBytes = FuncHandler.CustomerQtyCountExcel(ExcelList, Excel_Headers, model);
                var fileName = "業務件數報表" + DateTime.Now.ToString("yyyyMMddHHmm") + ".xlsx";
                return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            }
            catch (Exception ex)
            {
                ResultClass<string> resultClass = new ResultClass<string>();
                resultClass.ResultCode = "500";
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass);
            }
        }
        /// <summary>
        /// 業務件數明細 Customer_qty_count_Detail/customer_qty_count_AJAX.asp
        /// </summary>
        /// <param name="U_num">K0311 OR BC0900 </param>
        /// <param name="Qty_Date_S">113/09/01</param>
        /// <param name="Qty_Date_E">113/09/30</param>
        [HttpGet("Customer_qty_count_Detail")]
        public ActionResult<ResultClass<string>> Customer_qty_count_Detail(string U_num, string Qty_Date_S, string Qty_Date_E)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                Qty_Date_S = FuncHandler.ConvertROCToGregorian(Qty_Date_S);
                Qty_Date_E = FuncHandler.ConvertROCToGregorian(Qty_Date_E);

                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"
                    select ha.HA_id,hp.HP_id,ha.add_date,ha.cs_name,hp.pre_address,ha.plan_date,il.item_D_name,um.U_name
                    from User_M um
                    left join House_apply ha on um.U_num = ha.plan_num
                    left join house_pre hp on ha.HA_id = hp.HA_id
                    left join Item_list il on il.item_D_code = hp.pre_process_type
                    and il.item_M_code = 'pre_process_type' and il.item_D_type = 'Y'
                    WHERE um.del_tag = '0' and hp.del_tag = '0' and ha.del_tag = '0'
                    AND hp.pre_process_type IN ('PRCT0002', 'PRCT0003', 'PRCT0005')";
                if (U_num == "BC0900")
                {
                    T_SQL += " AND um.u_bc IN ('BC0900')";
                }
                else
                {
                    T_SQL += " AND ha.plan_num=@plan_num";
                    parameters.Add(new SqlParameter("@plan_num", U_num));
                }

                T_SQL += " AND (ha.add_date >= @Qty_Date_S + ' 00:00:00' AND ha.add_date <= @Qty_Date_E + ' 23:59:59') ";
                T_SQL += " group by ha.HA_id,hp.HP_id,ha.add_date,ha.cs_name,hp.pre_address,ha.plan_date,il.item_D_name,um.U_name";
                parameters.Add(new SqlParameter("@Qty_Date_S", Qty_Date_S));
                parameters.Add(new SqlParameter("@Qty_Date_E", Qty_Date_E));
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

        #region 進件預估核准撥款/業務件數統計表 Incoming_parts.asp
        /// <summary>
        /// 顯示區查詢權限 GetIncomingShowType/Incoming_parts.asp
        /// </summary>
        [HttpGet("GetIncomingShowType")]
        public ActionResult<ResultClass<string>> GetIncomingShowType()
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var User_Num = HttpContext.Session.GetString("UserID");
            var roleNum = HttpContext.Session.GetString("Role_num");
            SpecialClass specialClass = new SpecialClass();

            try
            {
                string[] strarr = new string[] { "1006", "1007", "1001", "1005" };
                if (strarr.Contains(roleNum))
                {
                    specialClass.special_check = "Y";
                    specialClass.BC_Strings = "xxx,BC0100,BC0200,BC0600,BC0300,BC0500,BC0400,BC0900";
                }
                else
                {
                    specialClass.special_check = "N";
                    specialClass.BC_Strings = "";
                }
                resultClass.ResultCode = "000";
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
        /// 提供工作日天數資料 GetMonthDay/Incoming_parts.asp
        /// </summary>
        [HttpPost("GetMonthDay")]
        public ActionResult<ResultClass<string>> GetMonthDay(Incoming_parts_req model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                var Inc_Date_S = FuncHandler.ConvertROCToGregorian(model.Inc_Date_S);
                var Inc_Date_E = FuncHandler.ConvertROCToGregorian(model.Inc_Date_E);

                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = "select * from fun_get_month_day (@Inc_Date_S, @Inc_Date_E)";
                parameters.Add(new SqlParameter("@Inc_Date_S", Inc_Date_S));
                parameters.Add(new SqlParameter("@Inc_Date_E", Inc_Date_E));
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
        /// 進件預估核准撥款_查詢 Incoming_PartA_Query/Incoming_parts.asp
        /// </summary>
        [HttpPost("Incoming_PartA_Query")]
        public ActionResult<ResultClass<string>> Incoming_PartA_Query(Incoming_parts_req model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var roleNum = HttpContext.Session.GetString("Role_num");
            var User_U_BC = HttpContext.Session.GetString("User_U_BC");

            try
            {
                var Inc_Date_S = FuncHandler.ConvertROCToGregorian(model.Inc_Date_S);
                var Inc_Date_E = FuncHandler.ConvertROCToGregorian(model.Inc_Date_E);

                //檢查權限
                if (string.IsNullOrEmpty(model.U_BC))
                {
                    string[] strarr = new string[] { "1006", "1007", "1001", "1005" };
                    if (strarr.Contains(roleNum))
                    {
                        model.U_BC = "xxx,BC0100,BC0200,BC0600,BC0300,BC0500,BC0400,BC0900";
                    }
                    else
                    {
                        model.U_BC = User_U_BC;
                    }
                }

                ADOData _adoData = new ADOData();
                #region SQL_工作天
                var parameters_dy = new List<SqlParameter>();
                var T_SQL_DY = "select * from fun_get_month_day (@Inc_Date_S, @Inc_Date_E)";
                parameters_dy.Add(new SqlParameter("@Inc_Date_S", Inc_Date_S));
                parameters_dy.Add(new SqlParameter("@Inc_Date_E", Inc_Date_E));
                #endregion
                DataTable dtResult_DY = _adoData.ExecuteQuery(T_SQL_DY, parameters_dy);
                var daysArray = dtResult_DY.AsEnumerable().Select(row => row.Field<string>("DateValue")).ToArray();

                #region SQL_明細
                var parameters_d = new List<SqlParameter>();
                var T_SQL_D = @"
                    SELECT house_apply.plan_num+'-'+convert(varchar(10), Send_amount_date,111)as [Key], count(*) as Val
                    FROM house_sendcase he
                    LEFT JOIN house_apply ON house_apply.ha_id = he.ha_id
                    LEFT JOIN house_pre_project ON house_pre_project.hp_project_id = he.hp_project_id 
                    WHERE he.del_tag = '0' AND (he.Send_amount_date >= @Inc_Date_S +' 00:00:00.000' AND he.Send_amount_date <= @Inc_Date_E+' 23:00:00.000')
                    and he.Send_amount_date is not null and he.sendcase_handle_type = 'Y'
                    and house_apply.del_tag = '0' AND house_pre_project.del_tag = '0'
                    group by convert(varchar(10), Send_amount_date,111),house_apply.plan_num
                    order by plan_num,convert(varchar(10), Send_amount_date,111)";
                parameters_d.Add(new SqlParameter("@Inc_Date_S", Inc_Date_S));
                parameters_d.Add(new SqlParameter("@Inc_Date_E", Inc_Date_E));
                #endregion
                DataTable dtResult_d = _adoData.ExecuteQuery(T_SQL_D, parameters_d);
                var partsKeyValList = dtResult_d.AsEnumerable().Select(row => new Incoming_Parts_KeyVal
                {
                    Key = row.Field<string>("Key"),
                    Value = row.Field<int>("Val")

                }).ToList();

                #region SQL_統計
                var parameters = new List<SqlParameter>();
                var T_SQL = @"
                    select case when il_1.item_d_code = 'BC0600' then 'BC0201' when il_1.item_d_code = 'BC0500' then 'BC0301' else il_1.item_d_code end AS com_id
                    ,il_1.item_D_name as com_name,User_M.U_num as plan_num,il.item_D_name as ti_name,il.item_sort,User_M.U_name,0 as _count
                    from User_M join Item_list il on il.item_D_code = User_M.U_PFT and il.del_tag = '0' and il.item_M_code = 'professional_title'
                    join Item_list il_1 on User_M.U_BC = il_1.item_D_code and il_1.del_tag = '0' and il_1.item_M_code = 'branch_company'
                    where User_M.u_bc IN (SELECT SplitValue FROM dbo.SplitStringFunction(@U_Check_BC_txt)) and (ISNULL(User_M.U_leave_date,'') = ''
                    or User_M.U_leave_date between @Inc_Date_S+' 00:00:00' and @Inc_Date_E+' 23:59:59')
                    and User_M.u_num not in ( SELECT house_apply.plan_num FROM house_sendcase he
                    LEFT JOIN house_apply ON house_apply.ha_id = he.ha_id AND house_apply.del_tag = '0'
                    LEFT JOIN house_pre_project ON house_pre_project.hp_project_id = he.hp_project_id AND house_pre_project.del_tag = '0'
                    LEFT JOIN (SELECT u_num, U_BC FROM user_m WHERE  del_tag = '0') User_M ON User_M.u_num = house_apply.plan_num
                    WHERE he.del_tag = '0'
                    AND (he.Send_amount_date >= @Inc_Date_S +' 00:00:00' AND he.Send_amount_date <= @Inc_Date_E +' 23:59:59' )
                    AND User_M.u_bc IN (SELECT SplitValue FROM dbo.SplitStringFunction(@U_Check_BC_txt))
                    AND he.send_amount_date IS NOT NULL AND he.sendcase_handle_type = 'Y' GROUP BY house_apply.plan_num )
                    and il.item_sort between '120' and '170' and User_M.U_arrive_date <= @Inc_Date_E
                    union
                    select case when il_1.item_d_code = 'BC0600' then 'BC0201' when il_1.item_d_code = 'BC0500' then 'BC0301' else il_1.item_d_code end AS com_id
                    ,il_1.item_D_name as com_name,house_apply.plan_num,il.item_D_name as ti_name,il.item_sort,User_M.U_name,count(*) as _count
                    FROM house_sendcase he LEFT JOIN house_apply ON house_apply.ha_id = he.ha_id AND house_apply.del_tag = '0'
                    LEFT JOIN house_pre_project ON house_pre_project.hp_project_id = he.hp_project_id AND house_pre_project.del_tag = '0'
                    LEFT JOIN (SELECT u_num,u_bc,U_name,U_PFT FROM user_m WHERE del_tag = '0') User_M ON User_M.u_num = house_apply.plan_num
                    join Item_list il on il.item_D_code = User_M.U_PFT and il.del_tag = '0' and il.item_M_code = 'professional_title'
                    join Item_list il_1 on il_1.item_D_code = User_M.U_BC and il_1.del_tag = '0' and il_1.item_M_code = 'branch_company'
                    WHERE he.del_tag = '0'
                    AND (he.Send_amount_date >= @Inc_Date_S+' 00:00:00' AND he.Send_amount_date <= @Inc_Date_E+' 23:59:59' )
                    AND User_M.u_bc IN (SELECT SplitValue FROM dbo.SplitStringFunction(@U_Check_BC_txt))
                    and he.Send_amount_date is not null and he.sendcase_handle_type = 'Y'
                    group by house_apply.plan_num,User_M.U_name,il.item_D_name,il.item_sort,il_1.item_D_name,il_1.item_D_code
                    order by com_id, item_sort";
                parameters.Add(new SqlParameter("@U_Check_BC_txt", model.U_BC));
                parameters.Add(new SqlParameter("@Inc_Date_S", Inc_Date_S));
                parameters.Add(new SqlParameter("@Inc_Date_E", Inc_Date_E));
                #endregion
                DataTable dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
                var incomingPartList = new List<Incoming_Part_res>();
                foreach (var row in dtResult.AsEnumerable())
                {
                    var partA = new Incoming_Part_res
                    {
                        plan_num = row.Field<string>("plan_num") ?? string.Empty,
                        com_name = row.Field<string>("com_name") ?? string.Empty,
                        ti_name = row.Field<string>("ti_name") ?? string.Empty,
                        totalcount = row.Field<int?>("_count") ?? 0,
                        U_name = row.Field<string>("U_name") ?? string.Empty
                    };

                    for (int i = 0; i < daysArray.Length; i++)
                    {
                        partA.DateValues.Add(daysArray[i], 0);
                    }
                    incomingPartList.Add(partA);
                }

                foreach (var partA in incomingPartList)
                {
                    foreach (var kv in partsKeyValList)
                    {
                        var keyParts = kv.Key.Split('-');

                        var partNum = keyParts[0];
                        var datePart = keyParts[1];

                        if (partNum == partA.plan_num && partA.DateValues.ContainsKey(datePart))
                        {
                            partA.DateValues[datePart] = kv.Value;
                        }
                    }
                }

                resultClass.ResultCode = "000";
                resultClass.objResult = JsonConvert.SerializeObject(incomingPartList);
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
        /// 業務件數統計表_查詢 Incoming_PartB_Query/Incoming_parts.asp
        /// </summary>
        [HttpPost("Incoming_PartB_Query")]
        public ActionResult<ResultClass<string>> Incoming_PartB_Query(Incoming_parts_req model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var roleNum = HttpContext.Session.GetString("Role_num");
            var User_U_BC = HttpContext.Session.GetString("User_U_BC");

            try
            {
                var Inc_Date_S = FuncHandler.ConvertROCToGregorian(model.Inc_Date_S);
                var Inc_Date_E = FuncHandler.ConvertROCToGregorian(model.Inc_Date_E);

                //檢查權限
                if (string.IsNullOrEmpty(model.U_BC))
                {
                    string[] strarr = new string[] { "1006", "1007", "1001", "1005" };
                    if (strarr.Contains(roleNum))
                    {
                        model.U_BC = "xxx,BC0100,BC0200,BC0600,BC0300,BC0500,BC0400,BC0900";
                    }
                    else
                    {
                        model.U_BC = User_U_BC;
                    }
                }

                ADOData _adoData = new ADOData();
                #region SQL_工作天
                var parameters_dy = new List<SqlParameter>();
                var T_SQL_DY = "select * from fun_get_month_day (@Inc_Date_S, @Inc_Date_E)";
                parameters_dy.Add(new SqlParameter("@Inc_Date_S", Inc_Date_S));
                parameters_dy.Add(new SqlParameter("@Inc_Date_E", Inc_Date_E));
                #endregion
                DataTable dtResult_DY = _adoData.ExecuteQuery(T_SQL_DY, parameters_dy);
                var daysArray = dtResult_DY.AsEnumerable().Select(row => row.Field<string>("DateValue")).ToArray();

                #region SQL_明細
                var parameters_d = new List<SqlParameter>();
                var T_SQL_D = @"
                    select ha.plan_num+'-'+convert(varchar(10), ha.add_date,111)as [Key], count(*) as Val
                    from User_M um
                    left join House_apply ha on um.U_num = ha.plan_num left join house_pre hp on ha.HA_id = hp.HA_id
                    WHERE um.del_tag = '0' and hp.del_tag = '0' and ha.del_tag = '0' AND hp.pre_process_type IN ('PRCT0002','PRCT0003','PRCT0005')
                    AND (ha.add_date >= @Inc_Date_S+' 00:00:00.000' AND ha.add_date <= @Inc_Date_E+' 23:00:00.000' )
                    group by ha.plan_num,convert(varchar(10), ha.add_date,111)";
                parameters_d.Add(new SqlParameter("@Inc_Date_S", Inc_Date_S));
                parameters_d.Add(new SqlParameter("@Inc_Date_E", Inc_Date_E));
                #endregion
                DataTable dtResult_d = _adoData.ExecuteQuery(T_SQL_D, parameters_d);
                var partsKeyValList = dtResult_d.AsEnumerable().Select(row => new Incoming_Parts_KeyVal
                {
                    Key = row.Field<string>("Key"),
                    Value = row.Field<int>("Val")

                }).ToList();

                #region SQL_統計
                var parameters = new List<SqlParameter>();
                var T_SQL = @"
                    select case when um.u_bc = 'BC0600' then 'BC0201' when um.u_bc = 'BC0500' then 'BC0301' else um.u_bc end u_bc
                    ,il.item_D_name as com_name,il_1.item_D_name as title_name,um.u_num,um.U_name,0 as _count,il_1.item_sort from User_M um 
                    left join Item_list il on um.U_BC = il.item_D_code and il.item_M_code = 'branch_company' and il.item_D_type = 'Y'
                    left join Item_list il_1 on um.U_PFT = il_1.item_D_code and il_1.item_M_code = 'professional_title' and il_1.item_D_type = 'Y'
                    where um.del_tag = '0' AND um.u_bc IN (SELECT SplitValue FROM dbo.SplitStringFunction(@U_Check_BC_txt))
                    and (ISNULL(U_leave_date,'') = '' or U_leave_date between @Inc_Date_S+' 00:00:00' and @Inc_Date_E+' 23:59:59')
                    and um.u_num not in ( select ha.plan_num from User_M um left join House_apply ha on um.U_num = ha.plan_num
                    left join house_pre hp on ha.HA_id = hp.HA_id
                    WHERE um.del_tag = '0' and hp.del_tag = '0' and ha.del_tag = '0' AND hp.pre_process_type IN ('PRCT0002','PRCT0003','PRCT0005')
                    AND (ha.add_date >= @Inc_Date_S + ' 00:00:00' AND ha.add_date <= @Inc_Date_E + ' 23:59:59' )
                    group by ha.plan_num ) and il_1.item_sort between '120' and '170' 
                    union
                    select case when u_bc = 'BC0600' then 'BC0201' when u_bc = 'BC0500' then 'BC0301' else u_bc end u_bc
                    ,com_name,title_name,plan_num,U_name,count(*) _count,item_sort from (
                    select ha.HA_id,um.u_bc,il.item_D_name as com_name,il_1.item_D_name as title_name,ha.plan_num,um.U_name
                    ,il_1.item_sort,hp.pre_address,ha.add_date from User_M um
                    left join House_apply ha on um.U_num = ha.plan_num left join house_pre hp on ha.HA_id = hp.HA_id
                    left join Item_list il on um.U_BC = il.item_D_code and il.item_M_code = 'branch_company' and il.item_D_type = 'Y'
                    left join Item_list il_1 on um.U_PFT = il_1.item_D_code and il_1.item_M_code = 'professional_title' and il_1.item_D_type = 'Y'
                    WHERE um.del_tag = '0' and hp.del_tag = '0' and ha.del_tag = '0' AND hp.pre_process_type IN ('PRCT0002','PRCT0003','PRCT0005')
                    AND (ha.add_date >= @Inc_Date_S + ' 00:00:00' AND ha.add_date <= @Inc_Date_E + ' 23:59:59' )
                    AND um.u_bc IN (SELECT SplitValue FROM dbo.SplitStringFunction(@U_Check_BC_txt))
                    group by ha.HA_id,um.u_bc,il.item_D_name,il_1.item_D_name,ha.plan_num,um.U_name,il_1.item_sort,hp.pre_address
                    ,ha.add_date) a group by a.u_bc,a.com_name,a.title_name,a.plan_num,a.U_name,a.item_sort order by U_BC,item_sort";
                parameters.Add(new SqlParameter("@U_Check_BC_txt", model.U_BC));
                parameters.Add(new SqlParameter("@Inc_Date_S", Inc_Date_S));
                parameters.Add(new SqlParameter("@Inc_Date_E", Inc_Date_E));
                #endregion
                DataTable dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
                var incomingPartList = new List<Incoming_Part_res>();
                foreach (var row in dtResult.AsEnumerable())
                {
                    var partA = new Incoming_Part_res
                    {
                        plan_num = row.Field<string>("u_num") ?? string.Empty,
                        com_name = row.Field<string>("com_name") ?? string.Empty,
                        ti_name = row.Field<string>("title_name") ?? string.Empty,
                        totalcount = row.Field<int?>("_count") ?? 0,
                        U_name = row.Field<string>("U_name") ?? string.Empty
                    };

                    for (int i = 0; i < daysArray.Length; i++)
                    {
                        partA.DateValues.Add(daysArray[i], 0);
                    }
                    incomingPartList.Add(partA);
                }

                foreach (var partA in incomingPartList)
                {
                    foreach (var kv in partsKeyValList)
                    {
                        var keyParts = kv.Key.Split('-');

                        var partNum = keyParts[0];
                        var datePart = keyParts[1];

                        if (partNum == partA.plan_num && partA.DateValues.ContainsKey(datePart))
                        {
                            partA.DateValues[datePart] = kv.Value;
                        }
                    }
                }

                resultClass.ResultCode = "000";
                resultClass.objResult = JsonConvert.SerializeObject(incomingPartList);
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
        /// 進件預估核准撥款/業務件數統計表_下載 Incoming_parts_Excel/Incoming_parts.asp
        /// </summary>
        [HttpPost("Incoming_parts_Excel")]
        public IActionResult Incoming_parts_Excel(Incoming_parts_req model)
        {
            var roleNum = HttpContext.Session.GetString("Role_num");
            var User_U_BC = HttpContext.Session.GetString("User_U_BC");

            try
            {
                var Inc_Date_S = FuncHandler.ConvertROCToGregorian(model.Inc_Date_S);
                var Inc_Date_E = FuncHandler.ConvertROCToGregorian(model.Inc_Date_E);

                //檢查權限
                if (string.IsNullOrEmpty(model.U_BC))
                {
                    string[] strarr = new string[] { "1006", "1007", "1001", "1005" };
                    if (strarr.Contains(roleNum))
                    {
                        model.U_BC = "xxx,BC0100,BC0200,BC0600,BC0300,BC0500,BC0400,BC0900";
                    }
                    else
                    {
                        model.U_BC = User_U_BC;
                    }
                }

                ADOData _adoData = new ADOData();
                #region SQL_工作天
                var parameters_dy = new List<SqlParameter>();
                var T_SQL_DY = "select * from fun_get_month_day (@Inc_Date_S, @Inc_Date_E)";
                parameters_dy.Add(new SqlParameter("@Inc_Date_S", Inc_Date_S));
                parameters_dy.Add(new SqlParameter("@Inc_Date_E", Inc_Date_E));
                #endregion
                DataTable dtResult_DY = _adoData.ExecuteQuery(T_SQL_DY, parameters_dy);
                var daysArray = dtResult_DY.AsEnumerable().Select(row => row.Field<string>("DateValue")).ToArray();

                #region 進件預估核准
                #region SQL_明細
                var parameters_da = new List<SqlParameter>();
                var T_SQL_DA = @"
                    SELECT house_apply.plan_num+'-'+convert(varchar(10), Send_amount_date,111)as [Key], count(*) as Val
                    FROM house_sendcase he
                    LEFT JOIN house_apply ON house_apply.ha_id = he.ha_id
                    LEFT JOIN house_pre_project ON house_pre_project.hp_project_id = he.hp_project_id 
                    WHERE he.del_tag = '0' AND (he.Send_amount_date >= @Inc_Date_S +' 00:00:00.000' AND he.Send_amount_date <= @Inc_Date_E+' 23:00:00.000')
                    and he.Send_amount_date is not null and he.sendcase_handle_type = 'Y'
                    and house_apply.del_tag = '0' AND house_pre_project.del_tag = '0'
                    group by convert(varchar(10), Send_amount_date,111),house_apply.plan_num
                    order by plan_num,convert(varchar(10), Send_amount_date,111)";
                parameters_da.Add(new SqlParameter("@Inc_Date_S", Inc_Date_S));
                parameters_da.Add(new SqlParameter("@Inc_Date_E", Inc_Date_E));
                #endregion
                DataTable dtResult_da = _adoData.ExecuteQuery(T_SQL_DA, parameters_da);
                var partsKeyValListA = dtResult_da.AsEnumerable().Select(row => new Incoming_Parts_KeyVal
                {
                    Key = row.Field<string>("Key"),
                    Value = row.Field<int>("Val")

                }).ToList();

                #region SQL_統計
                var parameters_a = new List<SqlParameter>();
                var T_SQL_A = @"
                    select case when il_1.item_d_code = 'BC0600' then 'BC0201' when il_1.item_d_code = 'BC0500' then 'BC0301' else il_1.item_d_code end AS com_id
                    ,il_1.item_D_name as com_name,User_M.U_num as plan_num,il.item_D_name as ti_name,il.item_sort,User_M.U_name,0 as _count
                    from User_M join Item_list il on il.item_D_code = User_M.U_PFT and il.del_tag = '0' and il.item_M_code = 'professional_title'
                    join Item_list il_1 on User_M.U_BC = il_1.item_D_code and il_1.del_tag = '0' and il_1.item_M_code = 'branch_company'
                    where User_M.u_bc IN (SELECT SplitValue FROM dbo.SplitStringFunction(@U_Check_BC_txt)) and (ISNULL(User_M.U_leave_date,'') = ''
                    or User_M.U_leave_date between @Inc_Date_S+' 00:00:00' and @Inc_Date_E+' 23:59:59')
                    and User_M.u_num not in ( SELECT house_apply.plan_num FROM house_sendcase he
                    LEFT JOIN house_apply ON house_apply.ha_id = he.ha_id AND house_apply.del_tag = '0'
                    LEFT JOIN house_pre_project ON house_pre_project.hp_project_id = he.hp_project_id AND house_pre_project.del_tag = '0'
                    LEFT JOIN (SELECT u_num, U_BC FROM user_m WHERE  del_tag = '0') User_M ON User_M.u_num = house_apply.plan_num
                    WHERE he.del_tag = '0'
                    AND (he.Send_amount_date >= @Inc_Date_S +' 00:00:00' AND he.Send_amount_date <= @Inc_Date_E +' 23:59:59' )
                    AND User_M.u_bc IN (SELECT SplitValue FROM dbo.SplitStringFunction(@U_Check_BC_txt))
                    AND he.send_amount_date IS NOT NULL AND he.sendcase_handle_type = 'Y' GROUP BY house_apply.plan_num )
                    and il.item_sort between '120' and '170' and User_M.U_arrive_date <= @Inc_Date_E
                    union
                    select case when il_1.item_d_code = 'BC0600' then 'BC0201' when il_1.item_d_code = 'BC0500' then 'BC0301' else il_1.item_d_code end AS com_id
                    ,il_1.item_D_name as com_name,house_apply.plan_num,il.item_D_name as ti_name,il.item_sort,User_M.U_name,count(*) as _count
                    FROM house_sendcase he LEFT JOIN house_apply ON house_apply.ha_id = he.ha_id AND house_apply.del_tag = '0'
                    LEFT JOIN house_pre_project ON house_pre_project.hp_project_id = he.hp_project_id AND house_pre_project.del_tag = '0'
                    LEFT JOIN (SELECT u_num,u_bc,U_name,U_PFT FROM user_m WHERE del_tag = '0') User_M ON User_M.u_num = house_apply.plan_num
                    join Item_list il on il.item_D_code = User_M.U_PFT and il.del_tag = '0' and il.item_M_code = 'professional_title'
                    join Item_list il_1 on il_1.item_D_code = User_M.U_BC and il_1.del_tag = '0' and il_1.item_M_code = 'branch_company'
                    WHERE he.del_tag = '0'
                    AND (he.Send_amount_date >= @Inc_Date_S+' 00:00:00' AND he.Send_amount_date <= @Inc_Date_E+' 23:59:59' )
                    AND User_M.u_bc IN (SELECT SplitValue FROM dbo.SplitStringFunction(@U_Check_BC_txt))
                    and he.Send_amount_date is not null and he.sendcase_handle_type = 'Y'
                    group by house_apply.plan_num,User_M.U_name,il.item_D_name,il.item_sort,il_1.item_D_name,il_1.item_D_code
                    order by com_id, item_sort";
                parameters_a.Add(new SqlParameter("@U_Check_BC_txt", model.U_BC));
                parameters_a.Add(new SqlParameter("@Inc_Date_S", Inc_Date_S));
                parameters_a.Add(new SqlParameter("@Inc_Date_E", Inc_Date_E));
                #endregion
                DataTable dtResulta = _adoData.ExecuteQuery(T_SQL_A, parameters_a);
                var incomingPartListA = new List<Incoming_Part_res>();
                foreach (var row in dtResulta.AsEnumerable())
                {
                    var partA = new Incoming_Part_res
                    {
                        plan_num = row.Field<string>("plan_num") ?? string.Empty,
                        com_name = row.Field<string>("com_name") ?? string.Empty,
                        ti_name = row.Field<string>("ti_name") ?? string.Empty,
                        totalcount = row.Field<int?>("_count") ?? 0,
                        U_name = row.Field<string>("U_name") ?? string.Empty
                    };

                    for (int i = 0; i < daysArray.Length; i++)
                    {
                        partA.DateValues.Add(daysArray[i], 0);
                    }
                    incomingPartListA.Add(partA);
                }

                foreach (var partA in incomingPartListA)
                {
                    foreach (var kv in partsKeyValListA)
                    {
                        var keyParts = kv.Key.Split('-');

                        var partNum = keyParts[0];
                        var datePart = keyParts[1];

                        if (partNum == partA.plan_num && partA.DateValues.ContainsKey(datePart))
                        {
                            partA.DateValues[datePart] = kv.Value;
                        }
                    }
                }
                #endregion

                #region 業務件數
                #region SQL_明細
                var parameters_db = new List<SqlParameter>();
                var T_SQL_DB = @"
                    select ha.plan_num+'-'+convert(varchar(10), ha.add_date,111)as [Key], count(*) as Val
                    from User_M um
                    left join House_apply ha on um.U_num = ha.plan_num left join house_pre hp on ha.HA_id = hp.HA_id
                    WHERE um.del_tag = '0' and hp.del_tag = '0' and ha.del_tag = '0' AND hp.pre_process_type IN ('PRCT0002','PRCT0003','PRCT0005')
                    AND (ha.add_date >= @Inc_Date_S+' 00:00:00.000' AND ha.add_date <= @Inc_Date_E+' 23:00:00.000' )
                    group by ha.plan_num,convert(varchar(10), ha.add_date,111)";
                parameters_db.Add(new SqlParameter("@Inc_Date_S", Inc_Date_S));
                parameters_db.Add(new SqlParameter("@Inc_Date_E", Inc_Date_E));
                #endregion
                DataTable dtResult_db = _adoData.ExecuteQuery(T_SQL_DB, parameters_db);
                var partsKeyValListB = dtResult_db.AsEnumerable().Select(row => new Incoming_Parts_KeyVal
                {
                    Key = row.Field<string>("Key"),
                    Value = row.Field<int>("Val")

                }).ToList();

                #region SQL_統計
                var parameters_b = new List<SqlParameter>();
                var T_SQL_B = @"
                    select case when um.u_bc = 'BC0600' then 'BC0201' when um.u_bc = 'BC0500' then 'BC0301' else um.u_bc end u_bc
                    ,il.item_D_name as com_name,il_1.item_D_name as title_name,um.u_num,um.U_name,0 as _count,il_1.item_sort from User_M um 
                    left join Item_list il on um.U_BC = il.item_D_code and il.item_M_code = 'branch_company' and il.item_D_type = 'Y'
                    left join Item_list il_1 on um.U_PFT = il_1.item_D_code and il_1.item_M_code = 'professional_title' and il_1.item_D_type = 'Y'
                    where um.del_tag = '0' AND um.u_bc IN (SELECT SplitValue FROM dbo.SplitStringFunction(@U_Check_BC_txt))
                    and (ISNULL(U_leave_date,'') = '' or U_leave_date between @Inc_Date_S+' 00:00:00' and @Inc_Date_E+' 23:59:59')
                    and um.u_num not in ( select ha.plan_num from User_M um left join House_apply ha on um.U_num = ha.plan_num
                    left join house_pre hp on ha.HA_id = hp.HA_id
                    WHERE um.del_tag = '0' and hp.del_tag = '0' and ha.del_tag = '0' AND hp.pre_process_type IN ('PRCT0002','PRCT0003','PRCT0005')
                    AND (ha.add_date >= @Inc_Date_S + ' 00:00:00' AND ha.add_date <= @Inc_Date_E + ' 23:59:59' )
                    group by ha.plan_num ) and il_1.item_sort between '120' and '170' 
                    union
                    select case when u_bc = 'BC0600' then 'BC0201' when u_bc = 'BC0500' then 'BC0301' else u_bc end u_bc
                    ,com_name,title_name,plan_num,U_name,count(*) _count,item_sort from (
                    select ha.HA_id,um.u_bc,il.item_D_name as com_name,il_1.item_D_name as title_name,ha.plan_num,um.U_name
                    ,il_1.item_sort,hp.pre_address,ha.add_date from User_M um
                    left join House_apply ha on um.U_num = ha.plan_num left join house_pre hp on ha.HA_id = hp.HA_id
                    left join Item_list il on um.U_BC = il.item_D_code and il.item_M_code = 'branch_company' and il.item_D_type = 'Y'
                    left join Item_list il_1 on um.U_PFT = il_1.item_D_code and il_1.item_M_code = 'professional_title' and il_1.item_D_type = 'Y'
                    WHERE um.del_tag = '0' and hp.del_tag = '0' and ha.del_tag = '0' AND hp.pre_process_type IN ('PRCT0002','PRCT0003','PRCT0005')
                    AND (ha.add_date >= @Inc_Date_S + ' 00:00:00' AND ha.add_date <= @Inc_Date_E + ' 23:59:59' )
                    AND um.u_bc IN (SELECT SplitValue FROM dbo.SplitStringFunction(@U_Check_BC_txt))
                    group by ha.HA_id,um.u_bc,il.item_D_name,il_1.item_D_name,ha.plan_num,um.U_name,il_1.item_sort,hp.pre_address
                    ,ha.add_date) a group by a.u_bc,a.com_name,a.title_name,a.plan_num,a.U_name,a.item_sort order by U_BC,item_sort";
                parameters_b.Add(new SqlParameter("@U_Check_BC_txt", model.U_BC));
                parameters_b.Add(new SqlParameter("@Inc_Date_S", Inc_Date_S));
                parameters_b.Add(new SqlParameter("@Inc_Date_E", Inc_Date_E));
                #endregion
                DataTable dtResult = _adoData.ExecuteQuery(T_SQL_B, parameters_b);
                var incomingPartListB = new List<Incoming_Part_res>();
                foreach (var row in dtResult.AsEnumerable())
                {
                    var partA = new Incoming_Part_res
                    {
                        plan_num = row.Field<string>("u_num") ?? string.Empty,
                        com_name = row.Field<string>("com_name") ?? string.Empty,
                        ti_name = row.Field<string>("title_name") ?? string.Empty,
                        totalcount = row.Field<int?>("_count") ?? 0,
                        U_name = row.Field<string>("U_name") ?? string.Empty
                    };

                    for (int i = 0; i < daysArray.Length; i++)
                    {
                        partA.DateValues.Add(daysArray[i], 0);
                    }
                    incomingPartListB.Add(partA);
                }

                foreach (var partA in incomingPartListB)
                {
                    foreach (var kv in partsKeyValListB)
                    {
                        var keyParts = kv.Key.Split('-');

                        var partNum = keyParts[0];
                        var datePart = keyParts[1];

                        if (partNum == partA.plan_num && partA.DateValues.ContainsKey(datePart))
                        {
                            partA.DateValues[datePart] = kv.Value;
                        }
                    }
                }
                #endregion

                var fileBytes = FuncHandler.IncomingPartsExcel(incomingPartListA, incomingPartListB, daysArray, model);
                var fileName = "每周業務件數統計表" + DateTime.Now.ToString("yyyyMMddHHmm") + ".xlsx";
                return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            }
            catch (Exception ex)
            {
                ResultClass<string> resultClass = new ResultClass<string>();
                resultClass.ResultCode = "500";
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass);
            }
        }
        /// <summary>
        /// 進件預估核准撥款_明細 Incoming_partA_Detail/get_detail.asp
        /// </summary>
        /// <param name="U_num">K0067</param>
        /// <param name="Inc_Date_S">113/10/1</param>
        /// <param name="Inc_Date_E">113/10/8</param>
        [HttpGet("Incoming_partA_Detail")]
        public ActionResult<ResultClass<string>> Incoming_partA_Detail(string U_num, string Inc_Date_S, string Inc_Date_E)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                Inc_Date_S = FuncHandler.ConvertROCToGregorian(Inc_Date_S);
                Inc_Date_E = FuncHandler.ConvertROCToGregorian(Inc_Date_E);
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"
                    SELECT he.HS_id,house_apply.cs_name,house_apply.cs_mtel1
                    ,(SELECT item_d_name FROM item_list WHERE item_m_code = 'appraise_company' AND item_d_type = 'Y'
                    AND item_d_code = he.appraise_company AND show_tag = '0' AND del_tag = '0') AS show_appraise_company
                    ,(SELECT item_d_name FROM item_list WHERE item_m_code = 'project_title' AND item_d_type = 'Y'
                    AND item_d_code = house_pre_project.project_title AND show_tag = '0' AND del_tag = '0') AS show_project_title
                    ,house_pre_project.project_apply_amount,(SELECT u_name FROM user_m WHERE u_num = he.sendcase_handle_num
                    AND del_tag = '0') AS sendcase_handle_name
                    ,he.sendcase_handle_date,he.Send_amount,convert(varchar,he.Send_amount_date,111) Send_amount_date,user_m.u_name 
                    FROM house_sendcase he
                    LEFT JOIN house_apply ON house_apply.ha_id = he.ha_id AND house_apply.del_tag = '0'
                    LEFT JOIN house_pre_project ON house_pre_project.hp_project_id = he.hp_project_id AND house_pre_project.del_tag = '0'
                    LEFT JOIN user_m ON User_M.u_num = house_apply.plan_num
                    WHERE  he.del_tag = '0' AND house_apply.plan_num = @U_num
                    AND (he.Send_amount_date >= @Inc_Date_S+' 00:00:00' AND he.Send_amount_date <= @Inc_Date_E+' 23:59:59')
                    ORDER BY he.Send_amount_date, he.hs_id";
                parameters.Add(new SqlParameter("@U_num", U_num));
                parameters.Add(new SqlParameter("@Inc_Date_S", Inc_Date_S));
                parameters.Add(new SqlParameter("@Inc_Date_E", Inc_Date_E));
                #endregion
                DataTable dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
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
        /// 業務件數統計表_明細 Incoming_partB_Detail/customer_qty_count_AJAX.asp
        /// </summary>
        /// <param name="U_num">K0067</param>
        /// <param name="Inc_Date_S">113/10/1</param>
        /// <param name="Inc_Date_E">113/10/8</param>
        [HttpGet("Incoming_partB_Detail")]
        public ActionResult<ResultClass<string>> Incoming_partB_Detail(string U_num, string Inc_Date_S, string Inc_Date_E)
        {
            ResultClass<string> resultClass=new ResultClass<string>();

            try
            {
                Inc_Date_S = FuncHandler.ConvertROCToGregorian(Inc_Date_S);
                Inc_Date_E = FuncHandler.ConvertROCToGregorian(Inc_Date_E);

                ADOData _adoData = new ADOData();
                #region SQL
                var parameters=new List<SqlParameter>();
                var T_SQL = @"
                    select ha.HA_id,hp.HP_id,ha.add_date,ha.cs_name,hp.pre_address,ha.plan_date,il.item_D_name,um.U_name from User_M um
                    left join House_apply ha on um.U_num = ha.plan_num
                    left join house_pre hp on ha.HA_id = hp.HA_id
                    left join Item_list il on il.item_D_code = hp.pre_process_type and il.item_M_code = 'pre_process_type' and il.item_D_type = 'Y'
                    WHERE um.del_tag = '0' and hp.del_tag = '0' and ha.del_tag = '0'
                    AND hp.pre_process_type IN ('PRCT0002','PRCT0003','PRCT0005')
                    AND(ha.add_date >= @Inc_Date_S + ' 00:00:00' AND ha.add_date <= @Inc_Date_E + ' 23:59:59')
                    and ha.plan_num = @U_num
                    group by ha.HA_id,hp.HP_id,ha.add_date,ha.cs_name,hp.pre_address,ha.plan_date,il.item_D_name,um.U_name";
                parameters.Add(new SqlParameter("@U_num", U_num));
                parameters.Add(new SqlParameter("@Inc_Date_S", Inc_Date_S));
                parameters.Add(new SqlParameter("@Inc_Date_E", Inc_Date_E));
                #endregion
                DataTable dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
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
        #endregion
    }
}
