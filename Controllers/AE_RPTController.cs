using KF_WebAPI.BaseClass;
using KF_WebAPI.BaseClass.AE;
using KF_WebAPI.FunctionHandler;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
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
        /// 提供業務名單資料(當show_more_type='Y') GetTeamUsers/select_team_more.asp
        /// </summary>
        /// <returns></returns>
        [HttpGet("GetTeamUsers")]
        public ActionResult<ResultClass<string>> GetTeamUsers()
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
                var frBCs = U_Check_BC_txt.Split(',').Distinct().ToList();
                var parameters_BC = frBCs.Select((k, i) => $"@BC_{i}").ToList();
                T_SQL = T_SQL + $" AND um.U_BC IN ({string.Join(", ", parameters_BC)})";
                for (int i = 0; i < frBCs.Count; i++)
                {
                    parameters.Add(new SqlParameter($"@BC_{i}", frBCs[i]));
                }
                T_SQL = T_SQL + " ORDER BY bc.item_sort,pft.item_sort";
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
        /// 提供可查詢業務人員之權限(當show_more_type='N'時 迴圈呼叫Feat_user_list) Get_UG_list/Feat_user_list.asp&_fn.asp
        /// </summary>
        [HttpGet("Get_UG_list")]
        public ActionResult<ResultClass<string>> Get_UG_list()
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
        public ActionResult<ResultClass<string>> Feat_user_list(Feat_user_list_req model)
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
        /// 業績報表_業務合計計算 Feat_user_list/Feat_user_list.asp
        /// </summary>
        /// <param name="TotalCheckAmount">總業績</param>
        /// <param name="TotalGetAmount">業績金額</param>
        /// <returns></returns>
        [HttpPost("Feat_user_Totail")]
        public ActionResult<ResultClass<string>> Feat_user_Totail(int TotalCheckAmount, int TotalGetAmount)
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
        /// 提供組長名單資料(當show_more_type='Y') GetTeamLeaders/select_leader_more.asp
        /// </summary>
        [HttpGet("GetTeamLeaders")]
        public ActionResult<ResultClass<string>> GetTeamLeaders()
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
                var frBCs = U_Check_BC_txt.Split(',').Distinct().ToList();
                var parameters_BC = frBCs.Select((k, i) => $"@BC_{i}").ToList();
                T_SQL = T_SQL + $" AND um.U_BC IN ({string.Join(", ", parameters_BC)})";
                for (int i = 0; i < frBCs.Count; i++)
                {
                    parameters.Add(new SqlParameter($"@BC_{i}", frBCs[i]));
                }
                T_SQL = T_SQL + " ORDER BY bc.item_sort,pft.item_sort";
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
        [HttpPost("Feat_leader_list")]
        public ActionResult<ResultClass<string>> Feat_leader_list(Feat_leader_list_req model)
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
        /// 業績目標設定_提供組長名單 GetMonthQuotaLeader
        /// </summary>
        [HttpGet("GetMonthQuotaLeader")]
        public ActionResult<ResultClass<string>> GetMonthQuotaLeader()
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
                var frBCs = U_Check_BC_txt.Split(',').Distinct().ToList();
                var parameters_BC = frBCs.Select((k, i) => $"@BC_{i}").ToList();
                T_SQL = T_SQL + $" AND um.U_BC IN ({string.Join(", ", parameters_BC)})";
                for (int i = 0; i < frBCs.Count; i++)
                {
                    parameters.Add(new SqlParameter($"@BC_{i}", frBCs[i]));
                }
                T_SQL = T_SQL + " ORDER BY bc.item_sort,pft.item_sort";
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
        /// 業績目標設定_讀取 GetMonthQuotaList/Month_quota_editor.asp
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
                    var T_SQL = "select group_M_name,group_D_name,group_D_code,sa.U_PFT_sort,sa.U_PFT_name";
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
                    DataTable dtResult=_adoData.ExecuteQuery(T_SQL, parameters);
                    var modelList = dtResult.AsEnumerable().Select(row => new MonthQuota_res 
                    {
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
        //業績目標設定_修改/新增 Month_quota_editor.asp
        [HttpPost("UpdMonthQuotaList")]
        public ActionResult<ResultClass<string>> UpdMonthQuotaList()
        {
            //修改Feat_target資料 若是沒資料就新增
            return Ok();
        }
        //放款公司選項
        //日報表_顯示個人(2020版)_讀取
        //日報表_(202210版)_讀取
        //日報表_(202210版)_匯出
        #endregion

        #region 業績報表_日報表(202106合計版)

        #endregion
    }
}
