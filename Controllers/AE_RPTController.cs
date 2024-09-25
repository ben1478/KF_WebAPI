using KF_WebAPI.BaseClass;
using KF_WebAPI.BaseClass.AE;
using KF_WebAPI.FunctionHandler;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Primitives;
using Newtonsoft.Json;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing.Drawing2D;
using System.Reflection;

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
                var T_SQL = "SELECT bc.item_D_name AS U_BC_name,um.U_num,um.U_name,pft.item_D_name AS U_PFT_name FROM User_M um";
                T_SQL = T_SQL + " LEFT JOIN Item_list bc ON bc.item_M_code = 'branch_company'  AND bc.item_D_code = um.U_BC AND bc.item_D_type = 'Y' AND bc.show_tag = '0' AND bc.del_tag = '0'";
                T_SQL = T_SQL + " LEFT JOIN Item_list pft ON pft.item_M_code = 'professional_title' AND pft.item_D_code = um.U_PFT AND pft.item_D_type = 'Y' AND pft.show_tag = '0' AND pft.del_tag = '0'";
                T_SQL = T_SQL + " WHERE um.del_tag = '0' AND bc.item_D_name is not null ";
                T_SQL = T_SQL + " AND U_num IN (Select distinct group_D_code from view_User_group)";
                T_SQL = T_SQL + " AND isnull(U_type,'')='' AND (U_leave_date is null OR U_leave_date >= DATEADD(MONTH, -2, GETDATE()))";
                T_SQL = T_SQL + " AND um.U_BC IN (SELECT SplitValue FROM dbo.SplitStringFunction(@U_Check_BC_txt))";
                T_SQL = T_SQL + " ORDER BY bc.item_sort,pft.item_sort";
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
                var T_SQL = "Select distinct group_D_code from view_User_group Where group_M_code = @group_M_code";
                T_SQL = T_SQL + " And GETDATE() between group_M_start_day and group_M_end_day";
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
                var T_SQL = "select House_apply.CS_name,House_sendcase.HS_id";
                T_SQL = T_SQL + " ,(select U_name FROM User_M where U_num = House_apply.plan_num AND del_tag='0') as plan_name";
                T_SQL = T_SQL + " ,(select item_D_name from Item_list where item_M_code = 'appraise_company' AND item_D_type='Y' AND item_D_code = House_sendcase.appraise_company  AND show_tag='0' AND del_tag='0') as show_appraise_company";
                T_SQL = T_SQL + " ,(select item_D_name from Item_list where item_M_code = 'project_title' AND item_D_type='Y' AND item_D_code = House_pre_project.project_title  AND show_tag='0' AND del_tag='0') as show_project_title";
                T_SQL = T_SQL + " ,get_amount_date,get_amount,Loan_rate,interest_rate_original,interest_rate_pass,charge_M,charge_flow,charge_agent";
                T_SQL = T_SQL + " ,charge_check,get_amount_final,House_apply.CS_introducer,HS_note,House_pre_project.project_title,exception_type,exception_rate";
                T_SQL = T_SQL + " from House_sendcase";
                T_SQL = T_SQL + " LEFT JOIN House_apply on House_apply.HA_id = House_sendcase.HA_id AND House_apply.del_tag='0' ";
                T_SQL = T_SQL + " LEFT JOIN House_pre_project on House_pre_project.HP_project_id = House_sendcase.HP_project_id AND House_pre_project.del_tag='0' ";
                T_SQL = T_SQL + " LEFT JOIN (select U_num ,U_BC FROM User_M where del_tag='0' ) User_M ON User_M.U_num = House_apply.plan_num";
                T_SQL = T_SQL + " where House_sendcase.del_tag = '0' AND isnull(House_sendcase.get_amount,'')<>'' ";
                T_SQL = T_SQL + " AND ( get_amount_date between @start_date AND @end_date )";
                T_SQL = T_SQL + " AND ( House_apply.plan_num = @U_num )";
                #region 提供可查詢公司權限
                string[] str_all = new string[] { "7011" };
                SpecialClass specialClass = FuncHandler.CheckSpecial(str_all, User_Num);
                if (specialClass.special_check == "N")
                {
                    string[] str = new string[] { "7020", "7021", "7022", "7023", "7024", "7025" };
                    specialClass = FuncHandler.CheckSpecial(str, User_Num);
                    if (specialClass.special_check == "Y")
                    {
                        T_SQL = T_SQL + "AND User_M.U_BC in ( @Check_U_BC )";
                        parameters.Add(new SqlParameter("@Check_U_BC", specialClass.BC_Strings));
                    }
                }
                #endregion
                T_SQL = T_SQL + " order by House_sendcase.HS_id desc";
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
                            var T_SQL_Ru = "select FR_D_ratio_A,FR_D_ratio_B,FR_D_rate,FR_D_discount*10 AS show_FR_D_discount,FR_D_replace,FR_D_discount from Feat_rule";
                            T_SQL_Ru = T_SQL_Ru + " Where show_tag='0' AND del_tag='0' AND FR_D_type='Y'";
                            T_SQL_Ru = T_SQL_Ru + " AND FR_M_code = @FR_M_code AND FR_D_rate=@FR_D_rate";
                            T_SQL_Ru = T_SQL_Ru + " AND ( FR_D_ratio_A <=@Loan_rate AND FR_D_ratio_B >=@Loan_rate )";

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
                var T_SQL = "SELECT bc.item_D_name AS U_BC_name,um.U_num,um.U_name,pft.item_D_name AS U_PFT_name FROM User_M um";
                T_SQL = T_SQL + " LEFT JOIN Item_list bc ON bc.item_M_code = 'branch_company'  AND bc.item_D_code = um.U_BC AND bc.item_D_type = 'Y' AND bc.show_tag = '0' AND bc.del_tag = '0'";
                T_SQL = T_SQL + " LEFT JOIN Item_list pft ON pft.item_M_code = 'professional_title' AND pft.item_D_code = um.U_PFT AND pft.item_D_type = 'Y' AND pft.show_tag = '0' AND pft.del_tag = '0'";
                T_SQL = T_SQL + " WHERE um.del_tag = '0' AND bc.item_D_name is not null ";
                T_SQL = T_SQL + " AND U_num IN (select group_M_code from User_group where del_tag='0' AND group_M_type='Y')";
                T_SQL = T_SQL + " AND isnull(U_type,'')='' AND (U_leave_date is null OR U_leave_date >= DATEADD(MONTH, -2, GETDATE()))";
                T_SQL = T_SQL + " AND um.U_BC IN (SELECT SplitValue FROM dbo.SplitStringFunction(@U_Check_BC_txt))";
                T_SQL = T_SQL + " ORDER BY bc.item_sort,pft.item_sort";
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
            var T_SQL = "select House_apply.CS_name,House_sendcase.HS_id";
            T_SQL = T_SQL + " ,(select U_name FROM User_M where U_num = House_apply.plan_num AND del_tag='0') as plan_name";
            T_SQL = T_SQL + " ,(select item_D_name from Item_list where item_M_code = 'appraise_company' AND item_D_type='Y' AND item_D_code = House_sendcase.appraise_company  AND show_tag='0' AND del_tag='0') as show_appraise_company";
            T_SQL = T_SQL + " ,(select item_D_name from Item_list where item_M_code = 'project_title' AND item_D_type='Y' AND item_D_code = House_pre_project.project_title  AND show_tag='0' AND del_tag='0') as show_project_title";
            T_SQL = T_SQL + " ,get_amount_date,get_amount,Loan_rate,interest_rate_original,interest_rate_pass,charge_M,charge_flow,charge_agent";
            T_SQL = T_SQL + " ,charge_check,get_amount_final,House_apply.CS_introducer,HS_note,House_pre_project.project_title,exception_type,exception_rate";
            T_SQL = T_SQL + " from House_sendcase";
            T_SQL = T_SQL + " LEFT JOIN House_apply on House_apply.HA_id = House_sendcase.HA_id AND House_apply.del_tag='0' ";
            T_SQL = T_SQL + " LEFT JOIN House_pre_project on House_pre_project.HP_project_id = House_sendcase.HP_project_id AND House_pre_project.del_tag='0'";
            T_SQL = T_SQL + " LEFT JOIN (select U_num ,U_BC FROM User_M where del_tag='0' ) User_M ON User_M.U_num = House_apply.plan_num";
            T_SQL = T_SQL + " where House_sendcase.del_tag = '0' AND isnull(House_sendcase.get_amount,'')<>''";
            T_SQL = T_SQL + " AND ( get_amount_date between @start_date AND @end_date )";
            T_SQL = T_SQL + " AND House_apply.plan_num IN (Select distinct group_D_code from view_User_group Where group_M_code = @group_M_code and del_tag='0')";
            T_SQL = T_SQL + " order by plan_num,HS_id desc";
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
                var T_SQL = "SELECT bc.item_D_name AS U_BC_name,um.U_num,um.U_name,pft.item_D_name AS U_PFT_nam, um.U_BC FROM User_M um";
                T_SQL = T_SQL + " LEFT JOIN Item_list bc ON bc.item_M_code = 'branch_company'  AND bc.item_D_code = um.U_BC AND bc.item_D_type = 'Y' AND bc.show_tag = '0' AND bc.del_tag = '0'";
                T_SQL = T_SQL + " LEFT JOIN Item_list pft ON pft.item_M_code = 'professional_title' AND pft.item_D_code = um.U_PFT AND pft.item_D_type = 'Y' AND pft.show_tag = '0' AND pft.del_tag = '0'";
                T_SQL = T_SQL + " WHERE um.del_tag = '0' AND bc.item_D_name is not null ";
                T_SQL = T_SQL + " AND U_num IN (select group_M_code from User_group where del_tag='0' AND group_M_type='Y')";
                T_SQL = T_SQL + " AND isnull(U_type,'')='' AND (U_leave_date is null OR U_leave_date >= DATEADD(MONTH, -2, GETDATE()))";
                T_SQL = T_SQL + " AND um.U_BC IN (SELECT SplitValue FROM dbo.SplitStringFunction(@U_Check_BC_txt))";
                T_SQL = T_SQL + " ORDER BY bc.item_sort,pft.item_sort";
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
                    var T_SQL = "select group_M_id,group_M_name,group_D_name,group_D_code,sa.U_PFT_sort,sa.U_PFT_name";
                    T_SQL = T_SQL + " ,isnull((select target_quota from Feat_target ft where ft.del_tag='0' and ft.group_id=ug.group_id and ft.target_ym= @YYYYMM),0) target_quota";
                    T_SQL = T_SQL + " ,@YYYYMM AS target_ym";
                    T_SQL = T_SQL + " FROM view_User_group ug";
                    T_SQL = T_SQL + " join view_User_sales leader on leader.U_num = ug.group_M_code ";
                    T_SQL = T_SQL + " join view_User_sales sa on sa.U_num = ug.group_D_code";
                    T_SQL = T_SQL + " where getdate() between ug.group_M_start_day and ug.group_M_end_day and leader.U_BC =@U_BC";
                    T_SQL = T_SQL + " order by leader.U_BC,group_M_name,sa.U_PFT_sort,sa.U_PFT_name,group_D_code";
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
                    var T_SQL_U = "update Feat_target set target_quota=@target_quota,edit_date=getdate(),edit_num=@edit_num,edit_ip=@edit_ip  ";
                    T_SQL_U = T_SQL_U + " where U_num=@U_num and target_ym=@target_ym";
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
                    var T_SQL_IN = "Insert into Feat_target(target_ym,target_quota,group_id,U_num,del_tag,add_date,add_num,add_ip,edit_date)";
                    T_SQL_IN = T_SQL_IN + " Values (@target_ym,@target_quota,@group_id,@U_num,'0',GETDATE(),@add_num,@add_ip,GETDATE())";
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
                return StatusCode(500, resultClass);
            }
        }
        /// <summary>
        /// 放款公司列表 GetCompanyList/_fn.asp
        /// </summary>
        [HttpGet("GetCompanyList")]
        public ActionResult<ResultClass<string>> GetCompanyList()
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = "select item_D_code,item_D_name from Item_list where item_M_code='fund_company' and item_D_type='Y' and del_tag='0'";
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
                var T_SQL = "select item_D_code,item_D_name from Item_list where item_M_code='branch_company' and item_D_type='Y' ";
                T_SQL = T_SQL + " and item_D_code in (select SplitValue from dbo.SplitStringFunction　((";
                T_SQL = T_SQL + " select 'zz'+REPLACE(U_Check_BC,'#',',') from User_M where U_num=@U_num))) order by item_sort";
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
                var T_SQL = "with A(leader_name,plan_name,group_id,group_M_id,group_M_title,group_M_start_day,group_M_end_day,day_incase_num,month_incase_num";
                T_SQL = T_SQL + " ,day_get_amount_num,day_get_amount,month_pass_num,month_get_amount_num,month_get_amount,month_pass_amount,month_pre_amount ) as (SELECT isnull(ug.group_M_name,'未分組') leader_name";
                T_SQL = T_SQL + " ,f.plan_name,ug.group_id,ug.group_M_id,ug.group_M_title,ug.group_M_start_day,ug.group_M_end_day";
                //日進件數 
                T_SQL = T_SQL + " ,(case when convert(varchar,Send_amount_date,112) = @reportDate_n  and Send_amount_date between ug.group_M_start_day and ug.group_M_end_day and Send_amount_date between ug.group_D_start_day and ug.group_D_end_day then 1 else 0 end ) as day_incase_num";
                //月進件數
                T_SQL = T_SQL + " ,(case when left(convert(varchar,Send_amount_date,112),6) = LEFT(@reportDate_n,6) AND convert(varchar,Send_amount_date,112) <= @reportDate_n  and Send_amount_date between ug.group_M_start_day and ug.group_M_end_day and Send_amount_date between ug.group_D_start_day and ug.group_D_end_day then 1 else 0 end ) as month_incase_num";
                //日撥款數 
                T_SQL = T_SQL + " ,(case when convert(varchar,get_amount_date,112) = @reportDate_n  and get_amount_date between ug.group_M_start_day and ug.group_M_end_day and get_amount_date between ug.group_D_start_day and ug.group_D_end_day then 1 else 0 end ) as day_get_amount_num";
                //日撥款額 
                T_SQL = T_SQL + " ,(case when convert(varchar,get_amount_date,112) = @reportDate_n  and get_amount_date between ug.group_M_start_day and ug.group_M_end_day and get_amount_date between ug.group_D_start_day and ug.group_D_end_day then get_amount else 0 end ) as day_get_amount";
                //月核准數
                T_SQL = T_SQL + " ,(case when left(convert(varchar,Send_result_date,112),6) = LEFT(@reportDate_n,6) AND convert(varchar,Send_result_date,112) <= @reportDate_n  and Send_result_date between ug.group_M_start_day and ug.group_M_end_day and Send_result_date between ug.group_D_start_day and ug.group_D_end_day AND Send_result_type = 'SRT002'  then 1 else 0 end ) as month_pass_num";
                //月撥款數
                T_SQL = T_SQL + " ,(case when left(convert(varchar,get_amount_date,112),6) = LEFT(@reportDate_n,6) AND convert(varchar,get_amount_date,112) <= @reportDate_n  and get_amount_date between ug.group_M_start_day and ug.group_M_end_day and get_amount_date between ug.group_D_start_day and ug.group_D_end_day then 1 else 0 end ) as month_get_amount_num";
                //月撥款額
                T_SQL = T_SQL + " ,(case when left(convert(varchar,get_amount_date,112),6) = LEFT(@reportDate_n,6) AND convert(varchar,get_amount_date,112) <= @reportDate_n  and get_amount_date between ug.group_M_start_day and ug.group_M_end_day and get_amount_date between ug.group_D_start_day and ug.group_D_end_day then get_amount else 0 end ) as month_get_amount";
                //已核未撥
                T_SQL = T_SQL + " ,(case when left(convert(varchar,Send_result_date,112),6) in (LEFT(@reportDate_n,6),@reportDate_b) AND convert(varchar,Send_result_date,112) <= @reportDate_n  and Send_result_date between ug.group_M_start_day and ug.group_M_end_day and Send_result_date between ug.group_D_start_day and ug.group_D_end_day AND Send_result_type = 'SRT002' AND isnull(check_amount_type,'') NOT IN('CKAT003') AND isnull(get_amount_type,'') NOT IN('GTAT002','GTAT003')  then pass_amount else 0 end ) as month_pass_amount";
                //預核額度
                T_SQL = T_SQL + " ,(case when left(convert(varchar,Send_result_date,112),6) in (LEFT(@reportDate_n,6),@reportDate_b) AND convert(varchar,Send_result_date,112) <= @reportDate_n  and Send_result_date between ug.group_M_start_day and ug.group_M_end_day and Send_result_date between ug.group_D_start_day and ug.group_D_end_day AND Send_result_type = 'SRT005' then pass_amount else 0 end ) as month_pre_amount";
                T_SQL = T_SQL + " FROM viewFeats f LEFT JOIN view_User_group ug ON ug.group_D_code = f.plan_num AND((Send_amount_date between ug.group_M_start_day and ug.group_M_end_day and Send_amount_date between ug.group_D_start_day and ug.group_D_end_day) ";
                T_SQL = T_SQL + " OR(Send_result_date between ug.group_M_start_day and ug.group_M_end_day and Send_result_date between ug.group_D_start_day and ug.group_D_end_day) OR(get_amount_date between ug.group_M_start_day and ug.group_M_end_day and get_amount_date between ug.group_D_start_day and ug.group_D_end_day) )  ";
                T_SQL = T_SQL + " where 1=1  AND(left(convert(varchar,Send_amount_date,112),6) = LEFT(@reportDate_n,6) OR left(convert(varchar,Send_result_date,112),6) in (LEFT(@reportDate_n,6),@reportDate_b) OR left(convert(varchar,get_amount_date,112),6) = LEFT(@reportDate_n,6)  )  and CONVERT(DATE,@reportDate_n,112) between ug.group_M_start_day and ug.group_M_end_day ";
                T_SQL = T_SQL + " AND f.fund_company IN (SELECT SplitValue FROM dbo.SplitStringFunction(@company)) AND U_BC = @U_BC )";
                T_SQL = T_SQL + " select @U_BC U_BC,leader_name,plan_name,group_M_id,group_M_title,sum(day_incase_num) as day_incase_num,sum(month_incase_num) as month_incase_num,sum(day_incase_num) as day_incase_num,sum(day_get_amount_num) as day_get_amount_num,sum(day_get_amount) as day_get_amount,sum(month_pass_num) as month_pass_num,sum(month_get_amount_num) as month_get_amount_num,sum(month_get_amount) as month_get_amount,sum(month_pass_amount) as month_pass_amount,sum(month_pre_amount) as month_pre_amount  FROM A  group by leader_name,plan_name,group_M_id,group_M_title  union";
                //顯示沒業績的組員
                T_SQL = T_SQL + " select @U_BC U_BC,leader_name,plan_name,group_M_id,group_M_title,0 as day_incase_num,0 as month_incase_num,0 as day_incase_num,0 as day_get_amount_num,0 as day_get_amount,0 as month_pass_num,0 as month_get_amount_num,0 as month_get_amount,0 as month_pass_amount,0 as month_pre_amount  from (select group_M_name leader_name,group_D_name plan_name,group_M_id,group_M_title  from view_User_group ug  where group_M_id in(select distinct group_M_id FROM A where @reportDate_n between A.group_M_start_day and A.group_M_end_day)  and group_id not in(select distinct group_id FROM A)   group by group_M_name,group_D_name,group_M_id,group_M_title ) B order by leader_name,plan_name";
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
                var T_SQL = "with A ( leader_name,plan_name,plan_num,group_id,group_M_id,group_M_title ,U_PFT_sort,U_PFT_name ,day_incase_num_FDCOM001,month_incase_num_FDCOM001 ";
                T_SQL = T_SQL + " ,day_incase_num_FDCOM003,month_incase_num_FDCOM003 ,day_incase_num_FDCOM004,month_incase_num_FDCOM004 ,day_incase_num_FDCOM005,month_incase_num_FDCOM005 ";
                T_SQL = T_SQL + " ,day_get_amount_num,day_get_amount,month_pass_num,month_get_amount_num ,month_get_amount_FDCOM001,month_get_amount_FDCOM003,month_get_amount_FDCOM004,month_get_amount_FDCOM005,month_pass_amount_FDCOM001,month_pre_amount_FDCOM001,month_pass_amount_FDCOM003,month_pass_amount_FDCOM004,month_pass_amount_FDCOM005,advance_payment_AE ) as ( SELECT isnull(ug.group_M_name,'未分組') leader_name";
                T_SQL = T_SQL + " ,ug.group_D_name,ug.group_D_code, ug.group_id, ug.group_M_id,ug.group_M_title ,sa.U_PFT_sort,sa.U_PFT_name";
                //新鑫 日進件數 
                T_SQL = T_SQL + " ,sum(case when 'FDCOM001'=fund_company and convert(varchar, Send_amount_date, 112) = @reportDate_n then 1 else 0 end ) as day_incase_num_FDCOM001";
                //新鑫 月進件數
                T_SQL = T_SQL + " ,sum(case when 'FDCOM001'=fund_company and left(convert(varchar, Send_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, Send_amount_date, 112) <= @reportDate_n then 1 else 0 end ) as month_incase_num_FDCOM001";
                //國&#23791; 日進件數
                T_SQL = T_SQL + " ,sum(case when 'FDCOM003'=fund_company and convert(varchar, Send_amount_date, 112) = @reportDate_n then 1 else 0 end ) as day_incase_num_FDCOM003";
                //國&#23791; 月進件數
                T_SQL = T_SQL + " ,sum(case when 'FDCOM003'=fund_company and left(convert(varchar, Send_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, Send_amount_date, 112) <= @reportDate_n then 1 else 0 end ) as month_incase_num_FDCOM003";
                //和潤 日進件數
                T_SQL = T_SQL + " ,sum(case when 'FDCOM004'=fund_company and convert(varchar, Send_amount_date, 112) = @reportDate_n then 1 else 0 end ) as day_incase_num_FDCOM004";
                //和潤 月進件數
                T_SQL = T_SQL + " ,sum(case when 'FDCOM004'=fund_company and left(convert(varchar, Send_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, Send_amount_date, 112) <= @reportDate_n then 1 else 0 end ) as month_incase_num_FDCOM004";
                //福斯 日進件數
                T_SQL = T_SQL + " ,sum(case when 'FDCOM005'=fund_company and convert(varchar, Send_amount_date, 112) = @reportDate_n then 1 else 0 end ) as day_incase_num_FDCOM005";
                //福斯 月進件數
                T_SQL = T_SQL + " ,sum(case when 'FDCOM005'=fund_company and left(convert(varchar, Send_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, Send_amount_date, 112) <= @reportDate_n then 1 else 0 end ) as month_incase_num_FDCOM005";
                //日撥款數
                T_SQL = T_SQL + " ,sum(case when convert(varchar, get_amount_date, 112) = @reportDate_n  then 1 else 0 end ) as day_get_amount_num";
                //日撥款額
                T_SQL = T_SQL + " ,sum(case when convert(varchar, get_amount_date, 112) = @reportDate_n  then get_amount else 0 end ) as day_get_amount";
                //月核准數 
                T_SQL = T_SQL + " ,sum(case when left(convert(varchar, Send_result_date, 112),6)=LEFT(@reportDate_n,6) AND convert(varchar, Send_result_date, 112) <= @reportDate_n  AND Send_result_type in ('SRT002','SRT005') then 1 else 0 end ) as month_pass_num";
                //月撥款數
                T_SQL = T_SQL + " ,sum(case when left(convert(varchar, get_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, get_amount_date, 112) <= @reportDate_n  then 1 else 0 end ) as month_get_amount_num";
                //新鑫 月撥款額
                T_SQL = T_SQL + " ,sum(case when 'FDCOM001'=fund_company and left(convert(varchar, get_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, get_amount_date, 112) <= @reportDate_n  then get_amount else 0 end ) as month_get_amount_FDCOM001";
                //國&#23791; 月撥款額
                T_SQL = T_SQL + " ,sum(case when 'FDCOM003'=fund_company and left(convert(varchar, get_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, get_amount_date, 112) <= @reportDate_n  then get_amount else 0 end ) as month_get_amount_FDCOM003";
                //和潤 月撥款額
                T_SQL = T_SQL + " ,sum(case when 'FDCOM004'=fund_company and left(convert(varchar, get_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, get_amount_date, 112) <= @reportDate_n  then get_amount else 0 end ) as month_get_amount_FDCOM004";
                //福斯 月撥款額
                T_SQL = T_SQL + " ,sum(case when 'FDCOM005'=fund_company and left(convert(varchar, get_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, get_amount_date, 112) <= @reportDate_n  then get_amount else 0 end ) as month_get_amount_FDCOM005";
                //新鑫 已核未撥
                T_SQL = T_SQL + " ,sum(case when 'FDCOM001'=fund_company and left(convert(varchar, Send_result_date, 112),6) in (LEFT(@reportDate_n,6),@reportDate_b) AND convert(varchar, Send_result_date, 112) <= @reportDate_n  AND Send_result_type = 'SRT002' AND isnull(check_amount_type,'') NOT IN ('CKAT003') AND isnull(get_amount_type,'') NOT IN ('GTAT002','GTAT003') then pass_amount else 0 end ) as month_pass_amount_FDCOM001";
                //新鑫 預核額度
                T_SQL = T_SQL + " ,sum(case when 'FDCOM001'=fund_company and left(convert(varchar, Send_result_date, 112),6) in (LEFT(@reportDate_n,6),@reportDate_b) AND convert(varchar, Send_result_date, 112) <= @reportDate_n  AND Send_result_type = 'SRT005' then pass_amount else 0 end ) as month_pre_amount_FDCOM001";
                //國&#23791; 已核未撥
                T_SQL = T_SQL + " ,sum(case when 'FDCOM003'=fund_company and left(convert(varchar, Send_result_date, 112),6) in (LEFT(@reportDate_n,6),@reportDate_b) AND convert(varchar, Send_result_date, 112) <= @reportDate_n  AND Send_result_type = 'SRT002' AND isnull(check_amount_type,'') NOT IN ('CKAT003') AND isnull(get_amount_type,'') NOT IN ('GTAT002','GTAT003') then pass_amount else 0 end ) as month_pass_amount_FDCOM003";
                //和潤 已核未撥
                T_SQL = T_SQL + " ,sum(case when 'FDCOM004'=fund_company and left(convert(varchar, Send_result_date, 112),6) in (LEFT(@reportDate_n,6),@reportDate_b) AND convert(varchar, Send_result_date, 112) <= @reportDate_n  AND Send_result_type = 'SRT002' AND isnull(check_amount_type,'') NOT IN ('CKAT003') AND isnull(get_amount_type,'') NOT IN ('GTAT002','GTAT003') then pass_amount else 0 end ) as month_pass_amount_FDCOM004";
                //福斯 已核未撥
                T_SQL = T_SQL + " ,sum(case when 'FDCOM005'=fund_company and left(convert(varchar, Send_result_date, 112),6) in (LEFT(@reportDate_n,6),@reportDate_b) AND convert(varchar, Send_result_date, 112) <= @reportDate_n  AND Send_result_type = 'SRT002' AND isnull(check_amount_type,'') NOT IN ('CKAT003') AND isnull(get_amount_type,'') NOT IN ('GTAT002','GTAT003') then pass_amount else 0 end ) as month_pass_amount_FDCOM005";
                //國&#23791; 代墊款(萬) 
                T_SQL = T_SQL + " ,sum(case when 'FDCOM003'=fund_company and left(convert(varchar, get_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, get_amount_date, 112) <= @reportDate_n  then advance_payment_AE else 0 end ) as advance_payment_AE";
                T_SQL = T_SQL + " FROM view_User_group ug join view_User_sales leader on leader.U_num = ug.group_M_code AND leader.U_BC = @U_BC ";
                T_SQL = T_SQL + " join view_User_sales sa on sa.U_num = ug.group_D_code ";
                T_SQL = T_SQL + " left join viewFeats f on ug.group_D_code = f.plan_num AND ( left(convert(varchar, f.Send_amount_date, 112),6) = LEFT(@reportDate_n,6) OR left(convert(varchar, f.Send_result_date, 112),6) in (LEFT(@reportDate_n,6),@reportDate_b) OR left(convert(varchar, f.get_amount_date, 112),6) = LEFT(@reportDate_n,6) ) ";
                T_SQL = T_SQL + " where 1=1 and @reportDate_n between ug.group_M_start_day and ug.group_M_end_day and @reportDate_n between ug.group_D_start_day and ug.group_D_end_day  group by isnull(ug.group_M_name,'未分組'),ug.group_D_name,ug.group_D_code, ug.group_id, ug.group_M_id,ug.group_M_title,sa.U_PFT_sort,sa.U_PFT_name)";
                T_SQL = T_SQL + " select @U_BC U_BC,a.* ,isnull(ft.target_quota,0) target_quota FROM A  ";
                T_SQL = T_SQL + " left join Feat_target ft on ft.del_tag='0'  and ft.U_num=A.plan_num  and ft.group_id=A.group_id and ft.target_ym=LEFT(@reportDate_n,6) ";
                T_SQL = T_SQL + " order by A.leader_name,A.U_PFT_sort,A.U_PFT_name,A.plan_num";
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
                var T_SQL_COMP = "select item_D_code,item_D_name from Item_list where item_M_code='branch_company' and item_D_type='Y' ";
                T_SQL_COMP = T_SQL_COMP + " and item_D_code in (select SplitValue from dbo.SplitStringFunction　((";
                T_SQL_COMP = T_SQL_COMP + " select 'zz'+REPLACE(U_Check_BC,'#',',') from User_M where U_num=@U_num))) order by item_sort";
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
                        var T_SQL = "with A ( leader_name,plan_name,plan_num,group_id,group_M_id,group_M_title ,U_PFT_sort,U_PFT_name ,day_incase_num_FDCOM001,month_incase_num_FDCOM001 ";
                        T_SQL = T_SQL + " ,day_incase_num_FDCOM003,month_incase_num_FDCOM003 ,day_incase_num_FDCOM004,month_incase_num_FDCOM004 ,day_incase_num_FDCOM005,month_incase_num_FDCOM005 ";
                        T_SQL = T_SQL + " ,day_get_amount_num,day_get_amount,month_pass_num,month_get_amount_num ,month_get_amount_FDCOM001,month_get_amount_FDCOM003,month_get_amount_FDCOM004,month_get_amount_FDCOM005,month_pass_amount_FDCOM001,month_pre_amount_FDCOM001,month_pass_amount_FDCOM003,month_pass_amount_FDCOM004,month_pass_amount_FDCOM005,advance_payment_AE ) as ( SELECT isnull(ug.group_M_name,'未分組') leader_name";
                        T_SQL = T_SQL + " ,ug.group_D_name,ug.group_D_code, ug.group_id, ug.group_M_id,ug.group_M_title ,sa.U_PFT_sort,sa.U_PFT_name";
                        //新鑫 日進件數 
                        T_SQL = T_SQL + " ,sum(case when 'FDCOM001'=fund_company and convert(varchar, Send_amount_date, 112) = @reportDate_n then 1 else 0 end ) as day_incase_num_FDCOM001";
                        //新鑫 月進件數
                        T_SQL = T_SQL + " ,sum(case when 'FDCOM001'=fund_company and left(convert(varchar, Send_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, Send_amount_date, 112) <= @reportDate_n then 1 else 0 end ) as month_incase_num_FDCOM001";
                        //國&#23791; 日進件數
                        T_SQL = T_SQL + " ,sum(case when 'FDCOM003'=fund_company and convert(varchar, Send_amount_date, 112) = @reportDate_n then 1 else 0 end ) as day_incase_num_FDCOM003";
                        //國&#23791; 月進件數
                        T_SQL = T_SQL + " ,sum(case when 'FDCOM003'=fund_company and left(convert(varchar, Send_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, Send_amount_date, 112) <= @reportDate_n then 1 else 0 end ) as month_incase_num_FDCOM003";
                        //和潤 日進件數
                        T_SQL = T_SQL + " ,sum(case when 'FDCOM004'=fund_company and convert(varchar, Send_amount_date, 112) = @reportDate_n then 1 else 0 end ) as day_incase_num_FDCOM004";
                        //和潤 月進件數
                        T_SQL = T_SQL + " ,sum(case when 'FDCOM004'=fund_company and left(convert(varchar, Send_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, Send_amount_date, 112) <= @reportDate_n then 1 else 0 end ) as month_incase_num_FDCOM004";
                        //福斯 日進件數
                        T_SQL = T_SQL + " ,sum(case when 'FDCOM005'=fund_company and convert(varchar, Send_amount_date, 112) = @reportDate_n then 1 else 0 end ) as day_incase_num_FDCOM005";
                        //福斯 月進件數
                        T_SQL = T_SQL + " ,sum(case when 'FDCOM005'=fund_company and left(convert(varchar, Send_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, Send_amount_date, 112) <= @reportDate_n then 1 else 0 end ) as month_incase_num_FDCOM005";
                        //日撥款數
                        T_SQL = T_SQL + " ,sum(case when convert(varchar, get_amount_date, 112) = @reportDate_n  then 1 else 0 end ) as day_get_amount_num";
                        //日撥款額
                        T_SQL = T_SQL + " ,sum(case when convert(varchar, get_amount_date, 112) = @reportDate_n  then get_amount else 0 end ) as day_get_amount";
                        //月核准數 
                        T_SQL = T_SQL + " ,sum(case when left(convert(varchar, Send_result_date, 112),6)=LEFT(@reportDate_n,6) AND convert(varchar, Send_result_date, 112) <= @reportDate_n  AND Send_result_type in ('SRT002','SRT005') then 1 else 0 end ) as month_pass_num";
                        //月撥款數
                        T_SQL = T_SQL + " ,sum(case when left(convert(varchar, get_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, get_amount_date, 112) <= @reportDate_n  then 1 else 0 end ) as month_get_amount_num";
                        //新鑫 月撥款額
                        T_SQL = T_SQL + " ,sum(case when 'FDCOM001'=fund_company and left(convert(varchar, get_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, get_amount_date, 112) <= @reportDate_n  then get_amount else 0 end ) as month_get_amount_FDCOM001";
                        //國&#23791; 月撥款額
                        T_SQL = T_SQL + " ,sum(case when 'FDCOM003'=fund_company and left(convert(varchar, get_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, get_amount_date, 112) <= @reportDate_n  then get_amount else 0 end ) as month_get_amount_FDCOM003";
                        //和潤 月撥款額
                        T_SQL = T_SQL + " ,sum(case when 'FDCOM004'=fund_company and left(convert(varchar, get_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, get_amount_date, 112) <= @reportDate_n  then get_amount else 0 end ) as month_get_amount_FDCOM004";
                        //福斯 月撥款額
                        T_SQL = T_SQL + " ,sum(case when 'FDCOM005'=fund_company and left(convert(varchar, get_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, get_amount_date, 112) <= @reportDate_n  then get_amount else 0 end ) as month_get_amount_FDCOM005";
                        //新鑫 已核未撥
                        T_SQL = T_SQL + " ,sum(case when 'FDCOM001'=fund_company and left(convert(varchar, Send_result_date, 112),6) in (LEFT(@reportDate_n,6),@reportDate_b) AND convert(varchar, Send_result_date, 112) <= @reportDate_n  AND Send_result_type = 'SRT002' AND isnull(check_amount_type,'') NOT IN ('CKAT003') AND isnull(get_amount_type,'') NOT IN ('GTAT002','GTAT003') then pass_amount else 0 end ) as month_pass_amount_FDCOM001";
                        //新鑫 預核額度
                        T_SQL = T_SQL + " ,sum(case when 'FDCOM001'=fund_company and left(convert(varchar, Send_result_date, 112),6) in (LEFT(@reportDate_n,6),@reportDate_b) AND convert(varchar, Send_result_date, 112) <= @reportDate_n  AND Send_result_type = 'SRT005' then pass_amount else 0 end ) as month_pre_amount_FDCOM001";
                        //國&#23791; 已核未撥
                        T_SQL = T_SQL + " ,sum(case when 'FDCOM003'=fund_company and left(convert(varchar, Send_result_date, 112),6) in (LEFT(@reportDate_n,6),@reportDate_b) AND convert(varchar, Send_result_date, 112) <= @reportDate_n  AND Send_result_type = 'SRT002' AND isnull(check_amount_type,'') NOT IN ('CKAT003') AND isnull(get_amount_type,'') NOT IN ('GTAT002','GTAT003') then pass_amount else 0 end ) as month_pass_amount_FDCOM003";
                        //和潤 已核未撥
                        T_SQL = T_SQL + " ,sum(case when 'FDCOM004'=fund_company and left(convert(varchar, Send_result_date, 112),6) in (LEFT(@reportDate_n,6),@reportDate_b) AND convert(varchar, Send_result_date, 112) <= @reportDate_n  AND Send_result_type = 'SRT002' AND isnull(check_amount_type,'') NOT IN ('CKAT003') AND isnull(get_amount_type,'') NOT IN ('GTAT002','GTAT003') then pass_amount else 0 end ) as month_pass_amount_FDCOM004";
                        //福斯 已核未撥
                        T_SQL = T_SQL + " ,sum(case when 'FDCOM005'=fund_company and left(convert(varchar, Send_result_date, 112),6) in (LEFT(@reportDate_n,6),@reportDate_b) AND convert(varchar, Send_result_date, 112) <= @reportDate_n  AND Send_result_type = 'SRT002' AND isnull(check_amount_type,'') NOT IN ('CKAT003') AND isnull(get_amount_type,'') NOT IN ('GTAT002','GTAT003') then pass_amount else 0 end ) as month_pass_amount_FDCOM005";
                        //國&#23791; 代墊款(萬) 
                        T_SQL = T_SQL + " ,sum(case when 'FDCOM003'=fund_company and left(convert(varchar, get_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, get_amount_date, 112) <= @reportDate_n  then advance_payment_AE else 0 end ) as advance_payment_AE";
                        T_SQL = T_SQL + " FROM view_User_group ug join view_User_sales leader on leader.U_num = ug.group_M_code AND leader.U_BC = @U_BC ";
                        T_SQL = T_SQL + " join view_User_sales sa on sa.U_num = ug.group_D_code ";
                        T_SQL = T_SQL + " left join viewFeats f on ug.group_D_code = f.plan_num AND ( left(convert(varchar, f.Send_amount_date, 112),6) = LEFT(@reportDate_n,6) OR left(convert(varchar, f.Send_result_date, 112),6) in (LEFT(@reportDate_n,6),@reportDate_b) OR left(convert(varchar, f.get_amount_date, 112),6) = LEFT(@reportDate_n,6) ) ";
                        T_SQL = T_SQL + " where 1=1 and @reportDate_n between ug.group_M_start_day and ug.group_M_end_day and @reportDate_n between ug.group_D_start_day and ug.group_D_end_day  group by isnull(ug.group_M_name,'未分組'),ug.group_D_name,ug.group_D_code, ug.group_id, ug.group_M_id,ug.group_M_title,sa.U_PFT_sort,sa.U_PFT_name)";
                        T_SQL = T_SQL + " select @U_BC U_BC,a.* ,isnull(ft.target_quota,0) target_quota FROM A  ";
                        T_SQL = T_SQL + " left join Feat_target ft on ft.del_tag='0'  and ft.U_num=A.plan_num  and ft.group_id=A.group_id and ft.target_ym=LEFT(@reportDate_n,6) ";
                        T_SQL = T_SQL + " order by A.leader_name,A.U_PFT_sort,A.U_PFT_name,A.plan_num";
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
                return StatusCode(500);
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
                var T_SQL = "select item_D_code,item_D_name,COUNT(*) from view_User_group ug";
                T_SQL = T_SQL + " join view_User_sales leader on leader.U_num = ug.group_M_code";
                T_SQL = T_SQL + " left join Item_list on Item_list.item_D_code=leader.U_BC and item_M_code='branch_company' and item_D_type='Y'";
                T_SQL = T_SQL + " group by item_D_code,item_D_name,item_sort order by item_sort";
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
                return StatusCode(500, resultClass);
            }
        }
        /// <summary>
        /// 業績報表_日報表(202106合計版)_查詢 Feat_daily_report_v2021_Query/Feat_daily_report_v2021.asp
        /// </summary>
        [HttpPost("Feat_daily_report_v2021_Query")]
        public ActionResult<ResultClass<string>> Feat_daily_report_v2021_Query(FeatDailyReport_req model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = "with A ( leader_name,plan_name,plan_num,group_id,group_M_id,group_M_title ,U_PFT_sort,U_PFT_name ,day_incase_num_FDCOM001,month_incase_num_FDCOM001 ";
                T_SQL = T_SQL + " ,day_incase_num_FDCOM003,month_incase_num_FDCOM003 ,day_incase_num_FDCOM004,month_incase_num_FDCOM004 ,day_incase_num_FDCOM005,month_incase_num_FDCOM005 ";
                T_SQL = T_SQL + " ,day_get_amount_num,day_get_amount,month_pass_num,month_get_amount_num ,month_get_amount_FDCOM001,month_get_amount_FDCOM003,month_get_amount_FDCOM004,month_get_amount_FDCOM005,month_pass_amount_FDCOM001,month_pre_amount_FDCOM001,month_pass_amount_FDCOM003,month_pass_amount_FDCOM004,month_pass_amount_FDCOM005,advance_payment_AE ) as ( SELECT isnull(ug.group_M_name,'未分組') leader_name";
                T_SQL = T_SQL + " ,ug.group_D_name,ug.group_D_code, ug.group_id, ug.group_M_id,ug.group_M_title ,sa.U_PFT_sort,sa.U_PFT_name";
                //新鑫 日進件數 
                T_SQL = T_SQL + " ,sum(case when 'FDCOM001'=fund_company and convert(varchar, Send_amount_date, 112) = @reportDate_n then 1 else 0 end ) as day_incase_num_FDCOM001";
                //新鑫 月進件數
                T_SQL = T_SQL + " ,sum(case when 'FDCOM001'=fund_company and left(convert(varchar, Send_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, Send_amount_date, 112) <= @reportDate_n then 1 else 0 end ) as month_incase_num_FDCOM001";
                //國&#23791; 日進件數
                T_SQL = T_SQL + " ,sum(case when 'FDCOM003'=fund_company and convert(varchar, Send_amount_date, 112) = @reportDate_n then 1 else 0 end ) as day_incase_num_FDCOM003";
                //國&#23791; 月進件數
                T_SQL = T_SQL + " ,sum(case when 'FDCOM003'=fund_company and left(convert(varchar, Send_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, Send_amount_date, 112) <= @reportDate_n then 1 else 0 end ) as month_incase_num_FDCOM003";
                //和潤 日進件數
                T_SQL = T_SQL + " ,sum(case when 'FDCOM004'=fund_company and convert(varchar, Send_amount_date, 112) = @reportDate_n then 1 else 0 end ) as day_incase_num_FDCOM004";
                //和潤 月進件數
                T_SQL = T_SQL + " ,sum(case when 'FDCOM004'=fund_company and left(convert(varchar, Send_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, Send_amount_date, 112) <= @reportDate_n then 1 else 0 end ) as month_incase_num_FDCOM004";
                //福斯 日進件數
                T_SQL = T_SQL + " ,sum(case when 'FDCOM005'=fund_company and convert(varchar, Send_amount_date, 112) = @reportDate_n then 1 else 0 end ) as day_incase_num_FDCOM005";
                //福斯 月進件數
                T_SQL = T_SQL + " ,sum(case when 'FDCOM005'=fund_company and left(convert(varchar, Send_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, Send_amount_date, 112) <= @reportDate_n then 1 else 0 end ) as month_incase_num_FDCOM005";
                //日撥款數
                T_SQL = T_SQL + " ,sum(case when convert(varchar, get_amount_date, 112) = @reportDate_n  then 1 else 0 end ) as day_get_amount_num";
                //日撥款額
                T_SQL = T_SQL + " ,sum(case when convert(varchar, get_amount_date, 112) = @reportDate_n  then get_amount else 0 end ) as day_get_amount";
                //月核准數 
                T_SQL = T_SQL + " ,sum(case when left(convert(varchar, Send_result_date, 112),6)=LEFT(@reportDate_n,6) AND convert(varchar, Send_result_date, 112) <= @reportDate_n  AND Send_result_type in ('SRT002','SRT005') then 1 else 0 end ) as month_pass_num";
                //月撥款數
                T_SQL = T_SQL + " ,sum(case when left(convert(varchar, get_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, get_amount_date, 112) <= @reportDate_n  then 1 else 0 end ) as month_get_amount_num";
                //新鑫 月撥款額
                T_SQL = T_SQL + " ,sum(case when 'FDCOM001'=fund_company and left(convert(varchar, get_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, get_amount_date, 112) <= @reportDate_n  then get_amount else 0 end ) as month_get_amount_FDCOM001";
                //國&#23791; 月撥款額
                T_SQL = T_SQL + " ,sum(case when 'FDCOM003'=fund_company and left(convert(varchar, get_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, get_amount_date, 112) <= @reportDate_n  then get_amount else 0 end ) as month_get_amount_FDCOM003";
                //和潤 月撥款額
                T_SQL = T_SQL + " ,sum(case when 'FDCOM004'=fund_company and left(convert(varchar, get_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, get_amount_date, 112) <= @reportDate_n  then get_amount else 0 end ) as month_get_amount_FDCOM004";
                //福斯 月撥款額
                T_SQL = T_SQL + " ,sum(case when 'FDCOM005'=fund_company and left(convert(varchar, get_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, get_amount_date, 112) <= @reportDate_n  then get_amount else 0 end ) as month_get_amount_FDCOM005";
                //新鑫 已核未撥
                T_SQL = T_SQL + " ,sum(case when 'FDCOM001'=fund_company and left(convert(varchar, Send_result_date, 112),6) in (LEFT(@reportDate_n,6),@reportDate_b) AND convert(varchar, Send_result_date, 112) <= @reportDate_n  AND Send_result_type = 'SRT002' AND isnull(check_amount_type,'') NOT IN ('CKAT003') AND isnull(get_amount_type,'') NOT IN ('GTAT002','GTAT003') then pass_amount else 0 end ) as month_pass_amount_FDCOM001";
                //新鑫 預核額度
                T_SQL = T_SQL + " ,sum(case when 'FDCOM001'=fund_company and left(convert(varchar, Send_result_date, 112),6) in (LEFT(@reportDate_n,6),@reportDate_b) AND convert(varchar, Send_result_date, 112) <= @reportDate_n  AND Send_result_type = 'SRT005' then pass_amount else 0 end ) as month_pre_amount_FDCOM001";
                //國&#23791; 已核未撥
                T_SQL = T_SQL + " ,sum(case when 'FDCOM003'=fund_company and left(convert(varchar, Send_result_date, 112),6) in (LEFT(@reportDate_n,6),@reportDate_b) AND convert(varchar, Send_result_date, 112) <= @reportDate_n  AND Send_result_type = 'SRT002' AND isnull(check_amount_type,'') NOT IN ('CKAT003') AND isnull(get_amount_type,'') NOT IN ('GTAT002','GTAT003') then pass_amount else 0 end ) as month_pass_amount_FDCOM003";
                //和潤 已核未撥
                T_SQL = T_SQL + " ,sum(case when 'FDCOM004'=fund_company and left(convert(varchar, Send_result_date, 112),6) in (LEFT(@reportDate_n,6),@reportDate_b) AND convert(varchar, Send_result_date, 112) <= @reportDate_n  AND Send_result_type = 'SRT002' AND isnull(check_amount_type,'') NOT IN ('CKAT003') AND isnull(get_amount_type,'') NOT IN ('GTAT002','GTAT003') then pass_amount else 0 end ) as month_pass_amount_FDCOM004";
                //福斯 已核未撥
                T_SQL = T_SQL + " ,sum(case when 'FDCOM005'=fund_company and left(convert(varchar, Send_result_date, 112),6) in (LEFT(@reportDate_n,6),@reportDate_b) AND convert(varchar, Send_result_date, 112) <= @reportDate_n  AND Send_result_type = 'SRT002' AND isnull(check_amount_type,'') NOT IN ('CKAT003') AND isnull(get_amount_type,'') NOT IN ('GTAT002','GTAT003') then pass_amount else 0 end ) as month_pass_amount_FDCOM005";
                //國&#23791; 代墊款(萬) 
                T_SQL = T_SQL + " ,sum(case when 'FDCOM003'=fund_company and left(convert(varchar, get_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, get_amount_date, 112) <= @reportDate_n  then advance_payment_AE else 0 end ) as advance_payment_AE";
                T_SQL = T_SQL + " FROM view_User_group ug join view_User_sales leader on leader.U_num = ug.group_M_code AND leader.U_BC = @U_BC ";
                T_SQL = T_SQL + " join view_User_sales sa on sa.U_num = ug.group_D_code ";
                T_SQL = T_SQL + " left join viewFeats f on ug.group_D_code = f.plan_num AND ( left(convert(varchar, f.Send_amount_date, 112),6) = LEFT(@reportDate_n,6) OR left(convert(varchar, f.Send_result_date, 112),6) in (LEFT(@reportDate_n,6),@reportDate_b) OR left(convert(varchar, f.get_amount_date, 112),6) = LEFT(@reportDate_n,6) ) ";
                T_SQL = T_SQL + " where 1=1 and @reportDate_n between ug.group_M_start_day and ug.group_M_end_day and @reportDate_n between ug.group_D_start_day and ug.group_D_end_day  group by isnull(ug.group_M_name,'未分組'),ug.group_D_name,ug.group_D_code, ug.group_id, ug.group_M_id,ug.group_M_title,sa.U_PFT_sort,sa.U_PFT_name)";
                T_SQL = T_SQL + " select SUM(a.day_incase_num_FDCOM001) as day_incase_num_FDCOM001_total,SUM(a.month_incase_num_FDCOM001) as month_incase_num_FDCOM001_total";
                T_SQL = T_SQL + " ,SUM(a.day_incase_num_FDCOM003) as day_incase_num_FDCOM003_total,SUM(a.month_incase_num_FDCOM003) as month_incase_num_FDCOM003_total";
                T_SQL = T_SQL + " ,SUM(a.day_get_amount_num) as day_get_amount_num_total,SUM(a.day_get_amount) as day_get_amount_total";
                T_SQL = T_SQL + " ,SUM(a.month_pass_num) as month_pass_num_total,SUM(a.month_get_amount_num) as month_get_amount_num_total";
                T_SQL = T_SQL + " ,SUM(a.month_get_amount_FDCOM001) as month_get_amount_FDCOM001_total,SUM(a.month_get_amount_FDCOM003) as month_get_amount_FDCOM003_total";
                T_SQL = T_SQL + " ,SUM(a.month_pass_amount_FDCOM001) as month_pass_amount_FDCOM001_total,SUM(a.month_pre_amount_FDCOM001) as month_pre_amount_FDCOM001_total";
                T_SQL = T_SQL + " ,SUM(a.month_pass_amount_FDCOM003) as month_pass_amount_FDCOM003_total,SUM(a.advance_payment_AE) as advance_payment_AE_total";
                T_SQL = T_SQL + " ,SUM(ft.target_quota) as target_quota_total";
                T_SQL = T_SQL + " ,CONVERT(DECIMAL(10, 2),CAST((SUM(a.month_get_amount_FDCOM001) + SUM(a.month_get_amount_FDCOM003)) AS FLOAT) / SUM(ft.target_quota) * 100) as percentage";
                T_SQL = T_SQL + " FROM A  left join Feat_target ft on ft.del_tag='0'  and ft.U_num=A.plan_num  and ft.group_id=A.group_id  and ft.target_ym=LEFT(@reportDate_n,6)";
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
                return StatusCode(500, resultClass);
            }
        }
        /// <summary>
        /// 日報表_(202106版)_下載 Feat_daily_report_v2021_Excel/Feat_daily_report_v2021.asp
        /// </summary>
        [HttpPost("Feat_daily_report_v2021_Excel")]
        public IActionResult Feat_daily_report_v2021_Excel(string reportDate_n)
        {
            try
            {
                FeatDailyReport_excel_Total modelFooter = new FeatDailyReport_excel_Total();
                ADOData _adoData = new ADOData();
                #region SQL_抓取所有分公司資料
                var parameters_comp = new List<SqlParameter>();
                var T_SQL_COMP = "select item_D_code,item_D_name,COUNT(*) as peoplecount from view_User_group ug";
                T_SQL_COMP = T_SQL_COMP + " join view_User_sales leader on leader.U_num = ug.group_M_code";
                T_SQL_COMP = T_SQL_COMP + " left join Item_list on Item_list.item_D_code=leader.U_BC and item_M_code='branch_company' and item_D_type='Y'";
                T_SQL_COMP = T_SQL_COMP + " group by item_D_code,item_D_name,item_sort order by item_sort";
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
                        var T_SQL = "with A ( leader_name,plan_name,plan_num,group_id,group_M_id,group_M_title ,U_PFT_sort,U_PFT_name ,day_incase_num_FDCOM001,month_incase_num_FDCOM001 ";
                        T_SQL = T_SQL + " ,day_incase_num_FDCOM003,month_incase_num_FDCOM003 ,day_incase_num_FDCOM004,month_incase_num_FDCOM004 ,day_incase_num_FDCOM005,month_incase_num_FDCOM005 ";
                        T_SQL = T_SQL + " ,day_get_amount_num,day_get_amount,month_pass_num,month_get_amount_num ,month_get_amount_FDCOM001,month_get_amount_FDCOM003,month_get_amount_FDCOM004,month_get_amount_FDCOM005,month_pass_amount_FDCOM001,month_pre_amount_FDCOM001,month_pass_amount_FDCOM003,month_pass_amount_FDCOM004,month_pass_amount_FDCOM005,advance_payment_AE ) as ( SELECT isnull(ug.group_M_name,'未分組') leader_name";
                        T_SQL = T_SQL + " ,ug.group_D_name,ug.group_D_code, ug.group_id, ug.group_M_id,ug.group_M_title ,sa.U_PFT_sort,sa.U_PFT_name";
                        //新鑫 日進件數 
                        T_SQL = T_SQL + " ,sum(case when 'FDCOM001'=fund_company and convert(varchar, Send_amount_date, 112) = @reportDate_n then 1 else 0 end ) as day_incase_num_FDCOM001";
                        //新鑫 月進件數
                        T_SQL = T_SQL + " ,sum(case when 'FDCOM001'=fund_company and left(convert(varchar, Send_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, Send_amount_date, 112) <= @reportDate_n then 1 else 0 end ) as month_incase_num_FDCOM001";
                        //國&#23791; 日進件數
                        T_SQL = T_SQL + " ,sum(case when 'FDCOM003'=fund_company and convert(varchar, Send_amount_date, 112) = @reportDate_n then 1 else 0 end ) as day_incase_num_FDCOM003";
                        //國&#23791; 月進件數
                        T_SQL = T_SQL + " ,sum(case when 'FDCOM003'=fund_company and left(convert(varchar, Send_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, Send_amount_date, 112) <= @reportDate_n then 1 else 0 end ) as month_incase_num_FDCOM003";
                        //和潤 日進件數
                        T_SQL = T_SQL + " ,sum(case when 'FDCOM004'=fund_company and convert(varchar, Send_amount_date, 112) = @reportDate_n then 1 else 0 end ) as day_incase_num_FDCOM004";
                        //和潤 月進件數
                        T_SQL = T_SQL + " ,sum(case when 'FDCOM004'=fund_company and left(convert(varchar, Send_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, Send_amount_date, 112) <= @reportDate_n then 1 else 0 end ) as month_incase_num_FDCOM004";
                        //福斯 日進件數
                        T_SQL = T_SQL + " ,sum(case when 'FDCOM005'=fund_company and convert(varchar, Send_amount_date, 112) = @reportDate_n then 1 else 0 end ) as day_incase_num_FDCOM005";
                        //福斯 月進件數
                        T_SQL = T_SQL + " ,sum(case when 'FDCOM005'=fund_company and left(convert(varchar, Send_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, Send_amount_date, 112) <= @reportDate_n then 1 else 0 end ) as month_incase_num_FDCOM005";
                        //日撥款數
                        T_SQL = T_SQL + " ,sum(case when convert(varchar, get_amount_date, 112) = @reportDate_n  then 1 else 0 end ) as day_get_amount_num";
                        //日撥款額
                        T_SQL = T_SQL + " ,sum(case when convert(varchar, get_amount_date, 112) = @reportDate_n  then get_amount else 0 end ) as day_get_amount";
                        //月核准數 
                        T_SQL = T_SQL + " ,sum(case when left(convert(varchar, Send_result_date, 112),6)=LEFT(@reportDate_n,6) AND convert(varchar, Send_result_date, 112) <= @reportDate_n  AND Send_result_type in ('SRT002','SRT005') then 1 else 0 end ) as month_pass_num";
                        //月撥款數
                        T_SQL = T_SQL + " ,sum(case when left(convert(varchar, get_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, get_amount_date, 112) <= @reportDate_n  then 1 else 0 end ) as month_get_amount_num";
                        //新鑫 月撥款額
                        T_SQL = T_SQL + " ,sum(case when 'FDCOM001'=fund_company and left(convert(varchar, get_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, get_amount_date, 112) <= @reportDate_n  then get_amount else 0 end ) as month_get_amount_FDCOM001";
                        //國&#23791; 月撥款額
                        T_SQL = T_SQL + " ,sum(case when 'FDCOM003'=fund_company and left(convert(varchar, get_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, get_amount_date, 112) <= @reportDate_n  then get_amount else 0 end ) as month_get_amount_FDCOM003";
                        //和潤 月撥款額
                        T_SQL = T_SQL + " ,sum(case when 'FDCOM004'=fund_company and left(convert(varchar, get_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, get_amount_date, 112) <= @reportDate_n  then get_amount else 0 end ) as month_get_amount_FDCOM004";
                        //福斯 月撥款額
                        T_SQL = T_SQL + " ,sum(case when 'FDCOM005'=fund_company and left(convert(varchar, get_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, get_amount_date, 112) <= @reportDate_n  then get_amount else 0 end ) as month_get_amount_FDCOM005";
                        //新鑫 已核未撥
                        T_SQL = T_SQL + " ,sum(case when 'FDCOM001'=fund_company and left(convert(varchar, Send_result_date, 112),6) in (LEFT(@reportDate_n,6),@reportDate_b) AND convert(varchar, Send_result_date, 112) <= @reportDate_n  AND Send_result_type = 'SRT002' AND isnull(check_amount_type,'') NOT IN ('CKAT003') AND isnull(get_amount_type,'') NOT IN ('GTAT002','GTAT003') then pass_amount else 0 end ) as month_pass_amount_FDCOM001";
                        //新鑫 預核額度
                        T_SQL = T_SQL + " ,sum(case when 'FDCOM001'=fund_company and left(convert(varchar, Send_result_date, 112),6) in (LEFT(@reportDate_n,6),@reportDate_b) AND convert(varchar, Send_result_date, 112) <= @reportDate_n  AND Send_result_type = 'SRT005' then pass_amount else 0 end ) as month_pre_amount_FDCOM001";
                        //國&#23791; 已核未撥
                        T_SQL = T_SQL + " ,sum(case when 'FDCOM003'=fund_company and left(convert(varchar, Send_result_date, 112),6) in (LEFT(@reportDate_n,6),@reportDate_b) AND convert(varchar, Send_result_date, 112) <= @reportDate_n  AND Send_result_type = 'SRT002' AND isnull(check_amount_type,'') NOT IN ('CKAT003') AND isnull(get_amount_type,'') NOT IN ('GTAT002','GTAT003') then pass_amount else 0 end ) as month_pass_amount_FDCOM003";
                        //和潤 已核未撥
                        T_SQL = T_SQL + " ,sum(case when 'FDCOM004'=fund_company and left(convert(varchar, Send_result_date, 112),6) in (LEFT(@reportDate_n,6),@reportDate_b) AND convert(varchar, Send_result_date, 112) <= @reportDate_n  AND Send_result_type = 'SRT002' AND isnull(check_amount_type,'') NOT IN ('CKAT003') AND isnull(get_amount_type,'') NOT IN ('GTAT002','GTAT003') then pass_amount else 0 end ) as month_pass_amount_FDCOM004";
                        //福斯 已核未撥
                        T_SQL = T_SQL + " ,sum(case when 'FDCOM005'=fund_company and left(convert(varchar, Send_result_date, 112),6) in (LEFT(@reportDate_n,6),@reportDate_b) AND convert(varchar, Send_result_date, 112) <= @reportDate_n  AND Send_result_type = 'SRT002' AND isnull(check_amount_type,'') NOT IN ('CKAT003') AND isnull(get_amount_type,'') NOT IN ('GTAT002','GTAT003') then pass_amount else 0 end ) as month_pass_amount_FDCOM005";
                        //國&#23791; 代墊款(萬) 
                        T_SQL = T_SQL + " ,sum(case when 'FDCOM003'=fund_company and left(convert(varchar, get_amount_date, 112),6) = LEFT(@reportDate_n,6) AND convert(varchar, get_amount_date, 112) <= @reportDate_n  then advance_payment_AE else 0 end ) as advance_payment_AE";
                        T_SQL = T_SQL + " FROM view_User_group ug join view_User_sales leader on leader.U_num = ug.group_M_code AND leader.U_BC = @U_BC ";
                        T_SQL = T_SQL + " join view_User_sales sa on sa.U_num = ug.group_D_code ";
                        T_SQL = T_SQL + " left join viewFeats f on ug.group_D_code = f.plan_num AND ( left(convert(varchar, f.Send_amount_date, 112),6) = LEFT(@reportDate_n,6) OR left(convert(varchar, f.Send_result_date, 112),6) in (LEFT(@reportDate_n,6),@reportDate_b) OR left(convert(varchar, f.get_amount_date, 112),6) = LEFT(@reportDate_n,6) ) ";
                        T_SQL = T_SQL + " where 1=1 and @reportDate_n between ug.group_M_start_day and ug.group_M_end_day and @reportDate_n between ug.group_D_start_day and ug.group_D_end_day  group by isnull(ug.group_M_name,'未分組'),ug.group_D_name,ug.group_D_code, ug.group_id, ug.group_M_id,ug.group_M_title,sa.U_PFT_sort,sa.U_PFT_name)";
                        T_SQL = T_SQL + " select SUM(a.day_incase_num_FDCOM001) as day_incase_num_FDCOM001_total,SUM(a.month_incase_num_FDCOM001) as month_incase_num_FDCOM001_total";
                        T_SQL = T_SQL + " ,SUM(a.day_incase_num_FDCOM003) as day_incase_num_FDCOM003_total,SUM(a.month_incase_num_FDCOM003) as month_incase_num_FDCOM003_total";
                        T_SQL = T_SQL + " ,SUM(a.day_get_amount_num) as day_get_amount_num_total,SUM(a.day_get_amount) as day_get_amount_total";
                        T_SQL = T_SQL + " ,SUM(a.month_pass_num) as month_pass_num_total,SUM(a.month_get_amount_num) as month_get_amount_num_total";
                        T_SQL = T_SQL + " ,SUM(a.month_get_amount_FDCOM001) as month_get_amount_FDCOM001_total,SUM(a.month_get_amount_FDCOM003) as month_get_amount_FDCOM003_total";
                        T_SQL = T_SQL + " ,SUM(a.month_pass_amount_FDCOM001) as month_pass_amount_FDCOM001_total,SUM(a.month_pre_amount_FDCOM001) as month_pre_amount_FDCOM001_total";
                        T_SQL = T_SQL + " ,SUM(a.month_pass_amount_FDCOM003) as month_pass_amount_FDCOM003_total,SUM(a.advance_payment_AE) as advance_payment_AE_total";
                        T_SQL = T_SQL + " ,ISNULL(SUM(ft.target_quota),0) as target_quota_total";
                        T_SQL = T_SQL + " ,ISNULL(CONVERT(DECIMAL(10, 2),CAST((SUM(a.month_get_amount_FDCOM001) + SUM(a.month_get_amount_FDCOM003)) AS FLOAT) / SUM(ft.target_quota) * 100),0) as percentage";
                        T_SQL = T_SQL + " FROM A  left join Feat_target ft on ft.del_tag='0'  and ft.U_num=A.plan_num  and ft.group_id=A.group_id  and ft.target_ym=LEFT(@reportDate_n,6)";
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
                return StatusCode(500);
            }
        }
        #endregion

        #region 案件狀態表 SendCaseStatus.asp
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
                    T_SQL = T_SQL + " where fund_Company=@Company";
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
                        var T_SQL_Q = "SELECT item_D_code,COUNT(fund_company) AS unApprCount FROM House_sendcase H";
                        T_SQL_Q = T_SQL_Q + " LEFT JOIN (SELECT item_D_code,item_D_name FROM Item_list WHERE item_M_code = 'Send_result_type' AND item_D_type = 'Y') ST";
                        T_SQL_Q = T_SQL_Q + " ON H.Send_result_type = ST.item_D_code";
                        T_SQL_Q = T_SQL_Q + " WHERE Send_amount_date >= @Send_Date_S AND Send_amount_date <= @Send_Date_E";
                        T_SQL_Q = T_SQL_Q + " AND Send_result_type <> 'SRT002' AND sendcase_handle_type = 'Y' AND fund_company = @Company";
                        T_SQL_Q = T_SQL_Q + " GROUP BY item_D_code";
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
                        var T_SQL_GU = "select item_D_code,count(fund_company) as GUDCount from House_sendcase H";
                        T_SQL_GU = T_SQL_GU + " left join (select item_D_code,item_D_name  from Item_list";
                        T_SQL_GU = T_SQL_GU + " where item_M_code = 'check_amount_type' AND item_D_type='Y') ST on H.check_amount_type=ST.item_D_code";
                        T_SQL_GU = T_SQL_GU + " where (Send_amount_date >= @Send_Date_S AND Send_amount_date <= @Send_Date_E ) and sendcase_handle_type='Y'";
                        T_SQL_GU = T_SQL_GU + " AND Send_result_type= 'SRT002' and HS_id not in ( select  H.HS_id from House_sendcase H";
                        T_SQL_GU = T_SQL_GU + " where (Send_amount_date >= @Send_Date_S AND Send_amount_date <= @Send_Date_E ) and sendcase_handle_type='Y'";
                        T_SQL_GU = T_SQL_GU + " AND isnull(H.Send_amount,'')<>'' AND get_amount_type in ( 'GTAT001','GTAT002','GTAT003')";
                        T_SQL_GU = T_SQL_GU + " ) AND  fund_company= @Company group by item_D_name,item_D_code  order by item_D_name";
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
                var T_SQL = "SELECT ISNULL(CS_PID, '') AS CS_ID,ISNULL(CS_register_address, '') AS addr,ISNULL(Loan_rate, '') AS Loan_rate";
                T_SQL = T_SQL + " ,REPLACE(REPLACE(CONVERT(varchar(10), Send_amount_date, 126), '-', '/'),YEAR(Send_amount_date)";
                T_SQL = T_SQL + " ,YEAR(Send_amount_date) - 1911) AS Send_amount_date,H.HS_id,H.pass_amount,H.Send_amount";
                T_SQL = T_SQL + " ,House_pre_project.project_title,House_apply.CS_name,House_apply.CS_MTEL1,House_apply.CS_MTEL2";
                T_SQL = T_SQL + " ,House_apply.CS_introducer,User_M.U_name AS plan_name,User_M.U_BC_name";
                T_SQL = T_SQL + " ,ISNULL((SELECT item_D_name FROM Item_list WHERE item_M_code = 'Send_result_type'";
                T_SQL = T_SQL + " AND item_D_type = 'Y' AND item_D_code = H.Send_result_type AND show_tag = '0'";
                T_SQL = T_SQL + " AND del_tag = '0'), '--') AS show_Send_result_type";
                T_SQL = T_SQL + " ,ISNULL((SELECT item_D_name FROM Item_list WHERE item_M_code = 'check_amount_type'";
                T_SQL = T_SQL + " AND item_D_type = 'Y' AND item_D_code = H.check_amount_type";
                T_SQL = T_SQL + " AND show_tag = '0'AND del_tag = '0'), '--') AS show_check_amount_type";
                T_SQL = T_SQL + " ,ISNULL((SELECT item_D_name FROM Item_list WHERE item_M_code = 'get_amount_type'";
                T_SQL = T_SQL + " AND item_D_type = 'Y' AND item_D_code = H.get_amount_type";
                T_SQL = T_SQL + " AND show_tag = '0' AND del_tag = '0'), '--') AS show_get_amount_type,";
                T_SQL = T_SQL + " (SELECT item_D_name FROM Item_list WHERE item_M_code = 'appraise_company' AND item_D_type = 'Y'";
                T_SQL = T_SQL + " AND item_D_code = H.appraise_company AND show_tag = '0' AND del_tag = '0') AS show_appraise_company,";
                T_SQL = T_SQL + " (SELECT item_D_name FROM Item_list WHERE item_M_code = 'fund_company'";
                T_SQL = T_SQL + " AND item_D_type = 'Y' AND item_D_code = H.fund_company AND show_tag = '0' AND del_tag = '0') AS show_fund_company,";
                T_SQL = T_SQL + " (SELECT item_D_name FROM Item_list WHERE item_M_code = 'project_title' AND item_D_type = 'Y'";
                T_SQL = T_SQL + " AND item_D_code = House_pre_project.project_title AND show_tag = '0'";
                T_SQL = T_SQL + " AND del_tag = '0') AS show_project_title,Users.U_name AS fin_name,";
                T_SQL = T_SQL + " CONVERT(varchar, House_pre_project.fin_date, 111) AS fin_date FROM House_sendcase H";
                T_SQL = T_SQL + " LEFT JOIN House_apply ON House_apply.HA_id = H.HA_id AND House_apply.del_tag = '0'";
                T_SQL = T_SQL + " LEFT JOIN House_pre_project ON House_pre_project.HP_project_id = H.HP_project_id AND House_pre_project.del_tag = '0'";
                T_SQL = T_SQL + " LEFT JOIN view_User_sales User_M ON User_M.U_num = House_apply.plan_num";
                T_SQL = T_SQL + " LEFT JOIN view_user_sales Users ON house_pre_project.fin_user = Users.U_num";
                T_SQL = T_SQL + " WHERE H.del_tag = '0' AND H.sendcase_handle_type = 'Y' AND ISNULL(H.Send_amount, '') <> ''";
                T_SQL = T_SQL + " AND (Send_amount_date >= @Send_Date_S AND Send_amount_date <= @Send_Date_E)";
                T_SQL = T_SQL + " AND fund_Company=@Company";
                switch (model.status)
                {
                    case "Appr":
                        T_SQL = T_SQL + " AND Send_result_type= 'SRT002'";
                        break;
                    case "unAppr":
                        T_SQL = T_SQL + " AND Send_result_type<> 'SRT002'";
                        break;
                    case "Pay":
                        T_SQL = T_SQL + " AND get_amount_type= 'GTAT002'";
                        break;
                    case "unPay":
                        T_SQL = T_SQL + " AND get_amount_type = 'GTAT003'";
                        break;
                    case "WPay":
                        T_SQL = T_SQL + " AND get_amount_type = 'GTAT001'";
                        break;
                    case "GU":
                        T_SQL = T_SQL + " AND Send_result_type= 'SRT002' and HS_id not in (";
                        T_SQL = T_SQL + " select  H.HS_id from House_sendcase H";
                        T_SQL = T_SQL + " where (Send_amount_date >= @Send_Date_S AND Send_amount_date <= @Send_Date_E";
                        T_SQL = T_SQL + " ) and sendcase_handle_type='Y'　AND isnull(H.Send_amount,'')<>''";
                        T_SQL = T_SQL + " AND get_amount_type in ( 'GTAT001','GTAT002','GTAT003'))";
                        break;
                    default:
                        break;
                }
                T_SQL = T_SQL + " ORDER BY Send_amount_date, H.HS_id DESC";
                parameters.Add(new SqlParameter("@Send_Date_S", model.Send_Date_S));
                parameters.Add(new SqlParameter("@Send_Date_E", model.Send_Date_E));
                parameters.Add(new SqlParameter("@Company", model.Company));
                #endregion
                DataTable dtResult=_adoData.ExecuteQuery(T_SQL, parameters);    
                if(dtResult.Rows.Count > 0)
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
                    T_SQL = T_SQL + " where fund_Company=@Company";
                    parameters.Add(new SqlParameter("@Company", model.Company));
                }
                parameters.Add(new SqlParameter("@Send_Date_S", model.Send_Date_S));
                parameters.Add(new SqlParameter("@Send_Date_E", model.Send_Date_E));
                #endregion
                DataTable dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
                if(dtResult.Rows.Count > 0)
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
                return StatusCode(500);
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
                var T_SQL = "SELECT ISNULL(CS_PID, '') AS CS_ID,ISNULL(CS_register_address, '') AS addr,ISNULL(Loan_rate, '') AS Loan_rate";
                T_SQL = T_SQL + " ,REPLACE(REPLACE(CONVERT(varchar(10), Send_amount_date, 126), '-', '/'),YEAR(Send_amount_date)";
                T_SQL = T_SQL + " ,YEAR(Send_amount_date) - 1911) AS Send_amount_date,H.HS_id,H.pass_amount,H.Send_amount";
                T_SQL = T_SQL + " ,House_pre_project.project_title,House_apply.CS_name,House_apply.CS_MTEL1,House_apply.CS_MTEL2";
                T_SQL = T_SQL + " ,House_apply.CS_introducer,User_M.U_name AS plan_name,User_M.U_BC_name";
                T_SQL = T_SQL + " ,ISNULL((SELECT item_D_name FROM Item_list WHERE item_M_code = 'Send_result_type'";
                T_SQL = T_SQL + " AND item_D_type = 'Y' AND item_D_code = H.Send_result_type AND show_tag = '0'";
                T_SQL = T_SQL + " AND del_tag = '0'), '--') AS show_Send_result_type";
                T_SQL = T_SQL + " ,ISNULL((SELECT item_D_name FROM Item_list WHERE item_M_code = 'check_amount_type'";
                T_SQL = T_SQL + " AND item_D_type = 'Y' AND item_D_code = H.check_amount_type";
                T_SQL = T_SQL + " AND show_tag = '0'AND del_tag = '0'), '--') AS show_check_amount_type";
                T_SQL = T_SQL + " ,ISNULL((SELECT item_D_name FROM Item_list WHERE item_M_code = 'get_amount_type'";
                T_SQL = T_SQL + " AND item_D_type = 'Y' AND item_D_code = H.get_amount_type";
                T_SQL = T_SQL + " AND show_tag = '0' AND del_tag = '0'), '--') AS show_get_amount_type,";
                T_SQL = T_SQL + " (SELECT item_D_name FROM Item_list WHERE item_M_code = 'appraise_company' AND item_D_type = 'Y'";
                T_SQL = T_SQL + " AND item_D_code = H.appraise_company AND show_tag = '0' AND del_tag = '0') AS show_appraise_company,";
                T_SQL = T_SQL + " (SELECT item_D_name FROM Item_list WHERE item_M_code = 'fund_company'";
                T_SQL = T_SQL + " AND item_D_type = 'Y' AND item_D_code = H.fund_company AND show_tag = '0' AND del_tag = '0') AS show_fund_company,";
                T_SQL = T_SQL + " (SELECT item_D_name FROM Item_list WHERE item_M_code = 'project_title' AND item_D_type = 'Y'";
                T_SQL = T_SQL + " AND item_D_code = House_pre_project.project_title AND show_tag = '0'";
                T_SQL = T_SQL + " AND del_tag = '0') AS show_project_title,Users.U_name AS fin_name,";
                T_SQL = T_SQL + " CONVERT(varchar, House_pre_project.fin_date, 111) AS fin_date FROM House_sendcase H";
                T_SQL = T_SQL + " LEFT JOIN House_apply ON House_apply.HA_id = H.HA_id AND House_apply.del_tag = '0'";
                T_SQL = T_SQL + " LEFT JOIN House_pre_project ON House_pre_project.HP_project_id = H.HP_project_id AND House_pre_project.del_tag = '0'";
                T_SQL = T_SQL + " LEFT JOIN view_User_sales User_M ON User_M.U_num = House_apply.plan_num";
                T_SQL = T_SQL + " LEFT JOIN view_user_sales Users ON house_pre_project.fin_user = Users.U_num";
                T_SQL = T_SQL + " WHERE H.del_tag = '0' AND H.sendcase_handle_type = 'Y' AND ISNULL(H.Send_amount, '') <> ''";
                T_SQL = T_SQL + " AND (Send_amount_date >= @Send_Date_S AND Send_amount_date <= @Send_Date_E)";
                T_SQL = T_SQL + " AND fund_Company=@Company";
                switch (model.status)
                {
                    case "Appr":
                        T_SQL = T_SQL + " AND Send_result_type= 'SRT002'";
                        break;
                    case "unAppr":
                        T_SQL = T_SQL + " AND Send_result_type<> 'SRT002'";
                        break;
                    case "Pay":
                        T_SQL = T_SQL + " AND get_amount_type= 'GTAT002'";
                        break;
                    case "unPay":
                        T_SQL = T_SQL + " AND get_amount_type = 'GTAT003'";
                        break;
                    case "WPay":
                        T_SQL = T_SQL + " AND get_amount_type = 'GTAT001'";
                        break;
                    case "GU":
                        T_SQL = T_SQL + " AND Send_result_type= 'SRT002' and HS_id not in (";
                        T_SQL = T_SQL + " select  H.HS_id from House_sendcase H";
                        T_SQL = T_SQL + " where (Send_amount_date >= @Send_Date_S AND Send_amount_date <= @Send_Date_E";
                        T_SQL = T_SQL + " ) and sendcase_handle_type='Y'　AND isnull(H.Send_amount,'')<>''";
                        T_SQL = T_SQL + " AND get_amount_type in ( 'GTAT001','GTAT002','GTAT003'))";
                        break;
                    default:
                        break;
                }
                T_SQL = T_SQL + " ORDER BY Send_amount_date, H.HS_id DESC";
                parameters.Add(new SqlParameter("@Send_Date_S", model.Send_Date_S));
                parameters.Add(new SqlParameter("@Send_Date_E", model.Send_Date_E));
                parameters.Add(new SqlParameter("@Company", model.Company));
                #endregion
                DataTable dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
                if(dtResult.Rows.Count > 0)
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
                return StatusCode(500);
            }
        }
        #endregion

        #region 年度業績還比表
        /// <summary>
        /// 提供查詢年月 GetSendcaseYYYMM/Performance_Plot.asp
        /// </summary>
        [HttpGet("GetSendcaseYYYMM")]
        public ActionResult<ResultClass<string>> GetSendcaseYYYMM()
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = "select distinct convert(varchar(4),(convert(varchar(4),get_amount_date,126)-1911))+'-'+convert(varchar(2),month(get_amount_date)) yyyMM";
                T_SQL = T_SQL + " ,convert(varchar(7),get_amount_date, 126) yyyymm from House_sendcase";
                T_SQL = T_SQL + " where year(get_amount_date) > year(DATEADD(year,-2,SYSDATETIME()))";
                T_SQL = T_SQL + " order by convert(varchar(7),get_amount_date, 126) desc";
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
                return StatusCode(500, resultClass);
            }
        }
        /// <summary>
        /// GetPerfByYYYY_PerformancePlot/_Ajaxhandler.asp
        /// </summary>
        /// <param name="this_YY">2024-09</param>
        /// <param name="last_YY">2023</param>
        [HttpGet("GetPerfByYYYY_PerformancePlot")]
        public ActionResult<ResultClass<string>> GetPerfByYYYY_PerformancePlot(string this_YY, string last_YY)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = "select sum( CAST(House_sendcase.get_amount AS DECIMAL))amount, month( get_amount_date)  mm";
                T_SQL = T_SQL + " ,Year(get_amount_date) yyyy from House_sendcase";
                T_SQL = T_SQL + " LEFT JOIN House_apply on House_apply.HA_id = House_sendcase.HA_id where get_amount_type='GTAT002' ";
                T_SQL = T_SQL + " and House_sendcase.del_tag = '0'  AND House_apply.del_tag='0'";
                T_SQL = T_SQL + " AND isnull(House_sendcase.get_amount,'') <>'' AND convert(varchar(7), (get_amount_date), 126)   between @last_YY and  @this_YY";
                T_SQL = T_SQL + " group by Year(get_amount_date), month( get_amount_date)";
                T_SQL = T_SQL + " order by Year(get_amount_date) asc, month( get_amount_date) asc";
                parameters.Add(new SqlParameter("@this_YY", this_YY));
                parameters.Add(new SqlParameter("@last_YY", last_YY + "-01"));
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
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                return StatusCode(500, resultClass);
            }
        }
        /// <summary>
        /// GetBCPerfByYYYYMM_PerformancePlot/_Ajaxhandler.asp
        /// </summary>
        /// <param name="this_YYMM">2024-09</param>
        /// <param name="last_YY">2023</param>
        [HttpGet("GetBCPerfByYYYYMM_PerformancePlot")]
        public ActionResult<ResultClass<string>> GetBCPerfByYYYYMM_PerformancePlot(string this_YYMM,string last_YY)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = "SELECT sum(CAST(H.get_amount AS DECIMAL))amount,Year(get_amount_date) yyyy,month(get_amount_date) mm,U_BC,BC_NA";
                T_SQL = T_SQL + " FROM House_sendcase H";
                T_SQL = T_SQL + " LEFT JOIN House_apply A ON A.HA_id = H.HA_id";
                T_SQL = T_SQL + " LEFT JOIN (select U.U_num,U.U_BC,UB.item_D_name BC_NA from User_M u";
                T_SQL = T_SQL + " LEFT JOIN Item_list ub ON ub.item_M_code='branch_company' AND ub.item_D_type='Y' AND ub.item_D_code=u.U_BC";
                T_SQL = T_SQL + " ) U on A.plan_num=U.U_num";
                T_SQL = T_SQL + " WHERE H.del_tag = '0' AND A.del_tag='0'  AND U_BC not in ('BC0700','BC0800')";
                T_SQL = T_SQL + " AND isnull(H.get_amount, '') <>''";
                T_SQL = T_SQL + " AND convert(varchar(7), (get_amount_date), 126) between @last_YYMM and @this_YYMM";
                T_SQL = T_SQL + " GROUP BY Year(get_amount_date),month(get_amount_date),U_BC,BC_NA";
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
                return StatusCode(500, resultClass);
            }
        }
        /// <summary>
        /// GetTotRateByYYYY_PerformancePlot/_Ajaxhandler.asp
        /// </summary>
        /// <param name="U_BC">BC0100</param>
        /// <param name="YYYYMM">2024-09</param>
        [HttpGet("GetTotRateByYYYY_PerformancePlot")]
        public ActionResult<ResultClass<string>> GetTotRateByYYYY_PerformancePlot(string? U_BC,string YYYYMM)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = "select ThisAmtTot,LastAmtTot,CAST(ROUND((ThisAmtTot-LastAmtTot)/LastAmtTot *100,2) AS decimal(5,2)) Rate from (";
                T_SQL = T_SQL + " select sum(ThisAmt)ThisAmtTot ,sum(LastAmt) LastAmtTot from ( select  ThisAmt ,LastAmt from (";
                T_SQL = T_SQL + " select sum(CAST(House_sendcase.get_amount AS DECIMAL))ThisAmt,Year(get_amount_date) yyyy,U_BC from House_sendcase";
                T_SQL = T_SQL + " LEFT JOIN House_apply on House_apply.HA_id = House_sendcase.HA_id";
                T_SQL = T_SQL + " LEFT JOIN (select U_num ,U_BC,U_name FROM User_M ) User_M ON User_M.U_num = House_apply.plan_num";
                T_SQL = T_SQL + " where House_sendcase.del_tag = '0' AND U_BC not in ('BC0700','BC0800') AND House_apply.del_tag='0'";
                T_SQL = T_SQL + " AND isnull(House_sendcase.get_amount,'') <> '' AND convert(varchar(7),(get_amount_date),126)";
                T_SQL = T_SQL + " between convert(varchar(4), @YYYYMM+'-01', 126)+'-01' and  @YYYYMM";
                if(!string.IsNullOrEmpty(U_BC))
                {
                    T_SQL = T_SQL + " And U_BC=@U_BC";
                }
                T_SQL = T_SQL + " group by Year(get_amount_date),U_BC ) T";
                T_SQL = T_SQL + " left join (";
                T_SQL = T_SQL + " select sum(CAST(House_sendcase.get_amount AS DECIMAL))LastAmt,Year(get_amount_date) yyyy,U_BC from House_sendcase";
                T_SQL = T_SQL + " LEFT JOIN House_apply on House_apply.HA_id = House_sendcase.HA_id";
                T_SQL = T_SQL + " LEFT JOIN (select U_num ,U_BC,U_name  FROM User_M  ) User_M ON User_M.U_num = House_apply.plan_num";
                T_SQL = T_SQL + " where House_sendcase.del_tag = '0' AND U_BC not in ('BC0700','BC0800') AND House_apply.del_tag='0'";
                T_SQL = T_SQL + " AND isnull(House_sendcase.get_amount,'') <> '' AND convert(varchar(7),(get_amount_date),126)";
                T_SQL = T_SQL + " between convert(varchar(4),  DATEADD(year,-1,@YYYYMM+'-01'), 126)+'-01'";
                T_SQL = T_SQL + " and convert(varchar(7),DATEADD(year,-1,@YYYYMM+'-01'), 126)";
                if (!string.IsNullOrEmpty(U_BC))
                {
                    T_SQL = T_SQL + " And U_BC=@U_BC";
                    parameters.Add(new SqlParameter("@U_BC", U_BC));
                }
                T_SQL = T_SQL + " group by Year(get_amount_date),U_BC ) L on T.U_BC=L.U_BC ) A ) T";
                parameters.Add(new SqlParameter("@YYYYMM", YYYYMM));
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
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                return StatusCode(500, resultClass);
            }
        }
        /// <summary>
        /// GetTargetPerfByYYYY_PerformancePlot/_Ajaxhandler.asp
        /// </summary>
        /// <param name="this_YY">2024</param>
        /// <param name="ThisTot">249855</param>
        /// <param name="U_BC">BC0100</param>
        [HttpGet("GetTargetPerfByYYYY_PerformancePlot")]
        public ActionResult<ResultClass<string>> GetTargetPerfByYYYY_PerformancePlot(string this_YY,int ThisTot,string? U_BC)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = "select *,CAST(ROUND(CAST(@ThisTot AS decimal(8,0)) / CAST(Target AS decimal(8,0)), 2)*100 AS decimal(3,0)) TargetRate,";
                T_SQL = T_SQL + " CAST(ROUND(CAST(Target-@ThisTot AS decimal(8,0)) / CAST(Target AS decimal(8,0)), 2)*100 AS decimal(3,0)) UnTargetRate";
                T_SQL = T_SQL + " from ( SELECT KeyVal yyyy,count(ColumnVal) BC_Count, sum(convert(decimal, ColumnVal)) Target FROM LogTable";
                T_SQL = T_SQL + " WHERE TableNA = 'Performance_Plot' and KeyVal=@this_YY";
                if (!string.IsNullOrEmpty(U_BC))
                {
                    T_SQL = T_SQL + " and ColumnNA=@U_BC";
                    parameters.Add(new SqlParameter("@U_BC", U_BC));
                }
                T_SQL = T_SQL + " group by KeyVal ) A";
                parameters.Add(new SqlParameter("@this_YY", this_YY));
                parameters.Add(new SqlParameter("@ThisTot", ThisTot));
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
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                return StatusCode(500, resultClass);
            }
        }
        /// <summary>
        /// GetSalesByYYYYMM_PerformancePlot/_Ajaxhandler.asp
        /// </summary>
        /// <param name="YYYYMM">2024-09</param>
        /// <param name="U_BC">BC0100</param>
        [HttpGet("GetSalesByYYYYMM_PerformancePlot")]
        public ActionResult<ResultClass<string>> GetSalesByYYYYMM_PerformancePlot(string YYYYMM, string? U_BC)
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
                    T_SQL = T_SQL + " select @YYYYMM yyyymm,E.U_BC,U_BC_Name,count(*) UserCount,convert(varchar(3),(year(@YYYYMM+'-01')-1911))+'-'+convert(varchar(2),month(@YYYYMM+'-01')) yyymm";
                    T_SQL = T_SQL + " from [dbo].[fun_GetWorkUsers] (@YYYYMM) E";
                    T_SQL = T_SQL + " Left Join  dbo.view_Sales_BC B on E.U_BC=B.U_BC where B.U_BC is not null and E.U_BC=@U_BC";
                    T_SQL = T_SQL + " and U_PFT in ('PFT010','PFT015','PFT020','PFT030','PFT050','PFT060','PFT300')  and E.U_BC=@U_BC group by  E.U_BC,U_BC_Name";
                    T_SQL = T_SQL + " union all";
                    T_SQL = T_SQL + " select convert(varchar(7),  DATEADD(year,-1,@YYYYMM+'-01'), 126) yyyymm, E.U_BC,U_BC_Name ,count(*) UserCount";
                    T_SQL = T_SQL + " ,convert(varchar(3), (year(DATEADD(year,-1,@YYYYMM+'-01'))-1911))+'-'+convert(varchar(2), month(DATEADD(year,-1,@YYYYMM+'-01'))) yyymm";
                    T_SQL = T_SQL + " from [dbo].[fun_GetWorkUsers] (convert(varchar(7),  DATEADD(year,-1,@YYYYMM+'-01'), 126)) E";
                    T_SQL = T_SQL + " Left Join  dbo.view_Sales_BC B on E.U_BC=B.U_BC where B.U_BC is not null";
                    T_SQL = T_SQL + " and U_PFT in ('PFT010','PFT015','PFT020','PFT030','PFT050','PFT060','PFT300')  and E.U_BC=@U_BC group by  E.U_BC,U_BC_Name";
                    parameters.Add(new SqlParameter("@U_BC", U_BC));
                }
                else
                {
                    T_SQL = T_SQL + " select  @YYYYMM yyyymm,count(*) UserCount,convert(varchar(3), (year(@YYYYMM+'-01')-1911))+'-'+convert(varchar(2), month(@YYYYMM+'-01')) yyymm";
                    T_SQL = T_SQL + " from [dbo].[fun_GetWorkUsers] (@YYYYMM) E";
                    T_SQL = T_SQL + " Left Join  dbo.view_Sales_BC B on E.U_BC=B.U_BC where B.U_BC is not null";
                    T_SQL = T_SQL + " and U_PFT in ('PFT010','PFT015','PFT020','PFT030','PFT050','PFT060','PFT300')";
                    T_SQL = T_SQL + " union all";
                    T_SQL = T_SQL + " select convert(varchar(7),DATEADD(year,-1,@YYYYMM+'-01'), 126) yyyy,count(*)UserCount";
                    T_SQL = T_SQL + " ,convert(varchar(3), (year(DATEADD(year,-1,@YYYYMM+'-01'))-1911))+'-'+convert(varchar(2), month(DATEADD(year,-1,@YYYYMM+'-01'))) yyymm";
                    T_SQL = T_SQL + " from [dbo].[fun_GetWorkUsers] (convert(varchar(7), DATEADD(year,-1,@YYYYMM+'-01'), 126)) E";
                    T_SQL = T_SQL + " Left Join  dbo.view_Sales_BC B on E.U_BC=B.U_BC where B.U_BC is not null";
                    T_SQL = T_SQL + " and U_PFT in ('PFT010','PFT015','PFT020','PFT030','PFT050','PFT060','PFT300')";
                }
                parameters.Add(new SqlParameter("@YYYYMM", YYYYMM));
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
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                return StatusCode(500, resultClass);
            }
        }
        /// <summary>
        /// GetBCPerfByYYYY_PerformancePlot/_Ajaxhandler.asp
        /// </summary>
        /// <param name="YYYY">2024</param>
        /// <param name="YYYYMM">2024-09</param>
        [HttpGet("GetBCPerfByYYYY_PerformancePlot")]
        public ActionResult<ResultClass<string>> GetBCPerfByYYYY_PerformancePlot(string YYYY, string YYYYMM)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = "SELECT sum(CAST(H.get_amount AS DECIMAL))amount, Year(get_amount_date) yyyy,U_BC,BC_NA,item_sort FROM House_sendcase H";
                T_SQL = T_SQL + " LEFT JOIN House_apply A ON A.HA_id = H.HA_id";
                T_SQL = T_SQL + " LEFT JOIN (select U.U_num,U.U_BC,UB.item_D_name BC_NA,item_sort from User_M u";
                T_SQL = T_SQL + " LEFT JOIN Item_list ub ON ub.item_M_code='branch_company' AND ub.item_D_type='Y' AND ub.item_D_code=u.U_BC";
                T_SQL = T_SQL + " ) U on A.plan_num=U.U_num";
                T_SQL = T_SQL + " WHERE H.del_tag = '0' AND A.del_tag='0' AND U_BC not in ('BC0700','BC0800') AND isnull(H.get_amount, '') <> ''";
                T_SQL = T_SQL + " AND convert(varchar(7), (get_amount_date), 126) between @YYYY+'-01' and @YYYYMM";
                T_SQL = T_SQL + " GROUP BY Year(get_amount_date) ,U_BC,BC_NA,item_sort";
                T_SQL = T_SQL + " ORDER BY item_sort ASC";
                parameters.Add(new SqlParameter("@YYYY", YYYY));
                parameters.Add(new SqlParameter("@YYYYMM", YYYYMM));
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
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                return StatusCode(500, resultClass);
            }
        }
        /// <summary>
        /// GetPerfByYYYY_BC_PerformancePlot/_Ajaxhandler.asp
        /// </summary>
        /// <param name="this_YY">2024-09</param>
        /// <param name="last_YY">2023</param>
        /// <param name="U_BC">BC0100</param>
        [HttpGet("GetPerfByYYYY_BC_PerformancePlot")]
        public ActionResult<ResultClass<string>> GetPerfByYYYY_BC_PerformancePlot(string this_YY,string last_YY,string U_BC)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = "select M.*,InCount,RateCount,CAST(ROUND(CAST(InCount AS decimal(8,2))/CAST(RateCount AS decimal(8,2)),2)*100 AS decimal(3,0)) TRate";
                T_SQL = T_SQL + " from (SELECT sum(CAST(H.get_amount AS DECIMAL)) amount,convert(varchar(7),get_amount_date,126)yyyymm,month(get_amount_date) mm,Year(get_amount_date) yyyy";
                T_SQL = T_SQL + " ,U_BC,BC_NA FROM House_sendcase H";
                T_SQL = T_SQL + " LEFT JOIN House_apply A ON A.HA_id = H.HA_id";
                T_SQL = T_SQL + " LEFT JOIN (select U.U_num,U.U_BC,UB.item_D_name BC_NA from User_M u";
                T_SQL = T_SQL + " LEFT JOIN Item_list ub ON ub.item_M_code='branch_company' AND ub.item_D_type='Y' AND ub.item_D_code=u.U_BC";
                T_SQL = T_SQL + " ) U on A.plan_num=U.U_num";
                T_SQL = T_SQL + " WHERE H.del_tag = '0'  AND A.del_tag='0' AND isnull(H.get_amount, '') <> ''";
                T_SQL = T_SQL + " AND convert(varchar(7),(get_amount_date),126) between @last_YY+'-01' and @this_YY";
                T_SQL = T_SQL + " GROUP BY convert(varchar(7),get_amount_date,126),Year(get_amount_date),month(get_amount_date),U_BC,BC_NA) M";
                T_SQL = T_SQL + " Left Join (select U_BC,yyyymm,sum(RateCount)RateCount from fun_GetRateCount(@last_YY+'-01',@this_YY)";
                T_SQL = T_SQL + " group by U_BC,yyyymm ) R on M.U_BC=R.U_BC and M.yyyymm=R.yyyymm";
                T_SQL = T_SQL + " Left Join (select U_BC,yyyymm,sum(InCount)InCount from fun_GetInCount(@last_YY+'-01',@this_YY)";
                T_SQL = T_SQL + " group by U_BC,yyyymm ) I on M.U_BC=I.U_BC and M.yyyymm=I.yyyymm";
                T_SQL = T_SQL + " where  M.U_BC=@U_BC ORDER BY M.yyyymm ASC";
                parameters.Add(new SqlParameter("@last_YY", last_YY));
                parameters.Add(new SqlParameter("@this_YY", this_YY));
                parameters.Add(new SqlParameter("@U_BC", U_BC));
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
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                return StatusCode(500, resultClass);
            }
        }
        /// <summary>
        /// SetTargetPerfByYYYY_BC_PerformancePlot/_Ajaxhandler.asp
        /// </summary>
        /// <param name="YYYY">2024</param>
        /// <param name="U_BC">BC0100</param>
        /// <param name="isSet">Y</param>
        [HttpGet("SetTargetPerfByYYYY_BC_PerformancePlot")]
        public ActionResult<ResultClass<string>> SetTargetPerfByYYYY_BC_PerformancePlot(int YYYY,string U_BC,string isSet)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                if (isSet == "N")
                    YYYY = YYYY - 1;

                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = "select D.*,isnull(M.Target,'0')[Target] from (SELECT item_D_code U_BC,UB.item_D_name U_BC_Name,item_sort";
                T_SQL = T_SQL + " FROM Item_list ub";
                T_SQL = T_SQL + " WHERE ub.item_M_code='branch_company' AND ub.item_D_type='Y' ) D";
                T_SQL = T_SQL + " Left join (SELECT KeyVal yyyy,convert(decimal,ColumnVal) Target,ColumnNA U_BC";
                T_SQL = T_SQL + " FROM LogTable L  WHERE TableNA = 'Performance_Plot' and KeyVal=@YYYY) M";
                T_SQL = T_SQL + " on m.U_BC=D.U_BC where D.U_BC not in ('BC0700','BC0800')";
                T_SQL = T_SQL + " and  D.U_BC=@U_BC order by item_sort";
                parameters.Add(new SqlParameter("@YYYY", YYYY));
                parameters.Add(new SqlParameter("@U_BC", U_BC));
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
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                return StatusCode(500, resultClass);
            }
        }
        /// <summary>
        /// 異動年度目標SaveBC_Target_PerformancePlot/_Ajaxhandler.asp
        /// </summary>
        [HttpPost("SaveBC_Target_PerformancePlot")]
        public ActionResult<ResultClass<string>> SaveBC_Target_PerformancePlot(string YYYY, string U_BC, Target_YYYY model)
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
                    T_SQL = T_SQL + "('Performance_Plot',@KeyVal,@ColumnNA,@ColumnVal,@LogID,GETDATE())";
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
                return StatusCode(500, resultClass);
            }
        }
        /// <summary>
        /// GetSalesListByYYYYMM_PerformancePlot/_Ajaxhandler.asp
        /// </summary>
        /// <param name="YYYYMM">2024-09</param>
        /// <param name="U_BC">BC0100</param>
        [HttpGet("GetSalesListByYYYYMM_PerformancePlot")]
        public ActionResult<ResultClass<string>> GetSalesListByYYYYMM_PerformancePlot(string U_BC, string YYYYMM)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = "select E.u_name,PFT_NA,U_BC_Name,convert(varchar(3),(year(U_arrive_date)-1911))+'-'+convert(varchar(2),month(U_arrive_date))+'-'+convert(varchar(2),Day(U_arrive_date)) U_arrive_date,";
                T_SQL = T_SQL + " convert(varchar(3),(year(U_leave_date)-1911))+'-'+convert(varchar(2),month(U_leave_date))+'-'+convert(varchar(2),Day(U_leave_date)) U_leave_date,";
                T_SQL = T_SQL + " case when U_leave_date is null then 'Y' else 'N' end isWork,";
                T_SQL = T_SQL + " DATEDIFF(MONTH,E.U_arrive_date,case when U_leave_date is null then GETDATE() else U_leave_date end)/12 AS JobYY,";
                T_SQL = T_SQL + " DATEDIFF(day,E.U_arrive_date,case when U_leave_date is null then GETDATE() else U_leave_date end ) % 365/30 AS JobMM";
                T_SQL = T_SQL + " from [dbo].[fun_GetWorkUsers] (@YYYYMM) E";
                T_SQL = T_SQL + " Left Join  dbo.view_Sales_BC B on E.U_BC=B.U_BC";
                T_SQL = T_SQL + " Left Join (select item_D_name PFT_NA,item_D_code U_PFT,item_sort from item_list where item_D_type='Y' and item_M_code ='professional_title') I";
                T_SQL = T_SQL + " on E.U_PFT=I.U_PFT";
                T_SQL = T_SQL + " where B.U_BC is not null and E.U_PFT in ('PFT010','PFT015','PFT020','PFT030','PFT050','PFT060','PFT300')";
                T_SQL = T_SQL + " and E.U_BC=@U_BC order by E.U_BC, item_sort,E.U_arrive_date asc";
                parameters.Add(new SqlParameter("@YYYYMM", YYYYMM));
                parameters.Add(new SqlParameter("@U_BC", U_BC));
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
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                return StatusCode(500, resultClass);
            }
        }
        #endregion

        #region 部門年度業績還比表 Performance_Plot_ByBC.asp

        #endregion

    }
}
