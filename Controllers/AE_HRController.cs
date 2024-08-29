using KF_WebAPI.BaseClass;
using KF_WebAPI.BaseClass.AE;
using KF_WebAPI.FunctionHandler;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Extensions.Configuration;
using Microsoft.IdentityModel.Tokens;
using Newtonsoft.Json;
using System;
using System.Data;
using System.Data.SqlClient;
using System.IdentityModel.Tokens.Jwt;
using System.Net;
using System.Reflection;
using System.Security.Claims;
using System.Text;

namespace KF_WebAPI.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class AE_HRController : ControllerBase
    {
        #region 請假單
        /// <summary>
        /// 請假單列表查詢-Flow_rest/Flow_rest_list.asp
        /// </summary>
        /// <returns></returns>
        [HttpPost("Flow_Rest_LQuery")]
        public ActionResult<ResultClass<string>> Flow_Rest_LQuery(Flow_rest_req model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            var User_Num = HttpContext.Session.GetString("UserID");
            var roleNum = HttpContext.Session.GetString("Role_num");

            try
            {
                var parameters = new List<SqlParameter>();
                var T_SQL = "select fr.FR_id,um1.U_name AS FR_U_name,it9.item_D_name AS FR_kind_show,fr.FR_date_begin,";
                T_SQL = T_SQL + " fr.FR_date_end,fr.FR_total_hour,it1.item_D_name AS FR_step_01_type_name,";
                T_SQL = T_SQL + " it2.item_D_name AS FR_step_02_type_name,it3.item_D_name AS FR_step_03_type_name,";
                T_SQL = T_SQL + " it4.item_D_name AS FR_step_HR_type_name,it5.item_D_name AS FR_step_01_sign_name,";
                T_SQL = T_SQL + " it6.item_D_name AS FR_step_02_sign_name,it7.item_D_name AS FR_step_03_sign_name,";
                T_SQL = T_SQL + " it8.item_D_name AS FR_step_HR_sign_name,it10.item_D_name AS FR_sign_type_name,fr.FR_note";
                T_SQL = T_SQL + " from Flow_rest fr";
                T_SQL = T_SQL + " LEFT JOIN User_M um1 ON um1.U_num = fr.FR_U_num";
                T_SQL = T_SQL + " LEFT JOIN Item_list it1 ON it1.item_M_code = 'Flow_step_type' AND it1.item_D_type = 'Y' AND it1.item_D_code = fr.FR_step_01_type AND it1.del_tag = '0'";
                T_SQL = T_SQL + " LEFT JOIN Item_list it2 ON it2.item_M_code = 'Flow_step_type' AND it2.item_D_type = 'Y' AND it2.item_D_code = fr.FR_step_02_type AND it2.del_tag = '0'";
                T_SQL = T_SQL + " LEFT JOIN Item_list it3 ON it3.item_M_code = 'Flow_step_type' AND it3.item_D_type = 'Y' AND it3.item_D_code = fr.FR_step_03_type AND it3.del_tag = '0'";
                T_SQL = T_SQL + " LEFT JOIN Item_list it4 ON it4.item_M_code = 'Flow_step_type' AND it4.item_D_type = 'Y' AND it4.item_D_code = fr.FR_step_HR_type AND it4.del_tag = '0'";
                T_SQL = T_SQL + " LEFT JOIN Item_list it5 ON it5.item_M_code = 'Flow_sign_type' AND it5.item_D_type = 'Y' AND it5.item_D_code = fr.FR_step_01_sign AND it5.del_tag = '0'";
                T_SQL = T_SQL + " LEFT JOIN Item_list it6 ON it6.item_M_code = 'Flow_sign_type' AND it6.item_D_type = 'Y' AND it6.item_D_code = fr.FR_step_02_sign AND it6.del_tag = '0'";
                T_SQL = T_SQL + " LEFT JOIN Item_list it7 ON it7.item_M_code = 'Flow_sign_type' AND it7.item_D_type = 'Y' AND it7.item_D_code = fr.FR_step_03_sign AND it7.del_tag = '0'";
                T_SQL = T_SQL + " LEFT JOIN Item_list it8 ON it8.item_M_code = 'Flow_sign_type' AND it8.item_D_type = 'Y' AND it8.item_D_code = fr.FR_step_HR_sign AND it8.del_tag = '0'";
                T_SQL = T_SQL + " LEFT JOIN Item_list it9 ON it9.item_M_code = 'FR_kind' AND it9.item_D_type = 'Y' AND it9.item_D_code = fr.FR_kind AND it9.del_tag = '0'";
                T_SQL = T_SQL + " LEFT JOIN Item_list it10 ON it10.item_M_code = 'Flow_sign_type' AND it10.item_D_type = 'Y' AND it10.item_D_code = fr.FR_sign_type AND it10.del_tag = '0'";
                T_SQL = T_SQL + " where fr.del_tag = '0' AND fr.cancel_date IS NULL";
                #region 權限判定
                //判斷是否 管理主管/人事助理/開發者 
                var validRoles = new HashSet<string> { "1005", "1006", "1007", "1001" };
                if (!validRoles.Contains(roleNum))
                {
                    T_SQL = T_SQL + " AND ( fr.FR_U_num = @U_num";
                    T_SQL = T_SQL + " OR (fr.FR_step_01_num = @U_num AND fr.FR_step_now = '1')";
                    T_SQL = T_SQL + " OR (fr.FR_step_02_num = @U_num AND fr.FR_step_now = '2')";
                    T_SQL = T_SQL + " OR (fr.FR_step_03_num = @U_num AND fr.FR_step_now = '3')";
                    T_SQL = T_SQL + " )";

                    parameters.Add(new SqlParameter("@U_num", User_Num));
                }
                #endregion
                #region 查詢條件
                //請假起迄查詢
                if (model.FR_date_begin.HasValue && model.FR_date_end.HasValue)
                {
                    T_SQL = T_SQL + " AND ((fr.FR_date_begin >= @FR_date_begin AND fr.FR_date_begin <= @FR_date_end) ";
                    T_SQL = T_SQL + " OR (fr.FR_date_end >= @FR_date_begin AND fr.FR_date_end <= @FR_date_end)";
                    T_SQL = T_SQL + " OR (fr.FR_date_begin <= @FR_date_begin AND fr.FR_date_end >= @FR_date_end))";

                    parameters.Add(new SqlParameter("@FR_date_begin", model.FR_date_begin));
                    parameters.Add(new SqlParameter("@FR_date_end", model.FR_date_end));
                }
                //部門區域查詢
                if (!string.IsNullOrEmpty(model.U_BC))
                {
                    T_SQL = T_SQL + "AND um1.U_BC=@U_BC";
                    parameters.Add(new SqlParameter("@U_BC", model.U_BC));
                }
                //簽核結果查詢
                if (!string.IsNullOrEmpty(model.FR_sign_type))
                {
                    T_SQL = T_SQL + "AND fr.FR_sign_type = @FR_sign_type";
                    parameters.Add(new SqlParameter("@FR_sign_type", model.FR_sign_type));
                }
                //外出,忘打卡等查詢
                if (!string.IsNullOrEmpty(model.FR_kind))
                {
                    // 分割 FR_kind，處理每個項目
                    var frKinds = model.FR_kind.Split(',').Distinct().ToList();

                    // 生成參數化的查詢部分
                    var parameterNames = frKinds.Select((k, i) => $"@FR_kind_{i}").ToList();
                    T_SQL = T_SQL + $" AND fr.FR_kind IN ({string.Join(", ", parameterNames)})";

                    // 添加參數到 SqlParameter 集合中
                    for (int i = 0; i < frKinds.Count; i++)
                    {
                        parameters.Add(new SqlParameter($"@FR_kind_{i}", frKinds[i]));
                    }
                }
                #endregion
                T_SQL = T_SQL + " ORDER BY fr.FR_date_begin DESC,fr.FR_id";

                ADOData _adoData = new ADOData();
                DataTable dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
                if (dtResult.Rows.Count > 0)
                {
                    resultClass.ResultCode = "000";
                    DataTable pageData = FuncHandler.GetPage(dtResult, model.page, 100);
                    resultClass.objResult = JsonConvert.SerializeObject(pageData);
                }
                else
                {
                    resultClass.ResultCode = "201";
                    resultClass.ResultMsg = "查無資料";
                }
                return Ok(resultClass);
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                return StatusCode(500, resultClass); // 返回 500 錯誤碼
            }
        }
        /// <summary>
        /// 請假單單筆查詢-Flow_rest/Flow_rest_V201803_detail.asp
        /// </summary>
        /// <returns></returns>
        [HttpPost("Flow_Rest_SQuery")]
        public ActionResult<ResultClass<string>> Flow_Rest_SQuery(string Fr_Id)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                if (!string.IsNullOrEmpty(Fr_Id))
                {
                    var parameters = new List<SqlParameter>();
                    var T_SQL = "SELECT FR.*,U1.U_name AS FR_U_name,U2.U_name AS FR_step_01_name,U3.U_name AS FR_step_02_name,";
                    T_SQL = T_SQL + " U4.U_name AS FR_step_03_name,I1.item_D_name AS FR_step_01_type_name,I2.item_D_name AS FR_step_02_type_name,";
                    T_SQL = T_SQL + " I3.item_D_name AS FR_step_03_type_name,I4.item_D_name AS FR_step_HR_type_name,";
                    T_SQL = T_SQL + " I5.item_D_name AS FR_step_01_sign_name,I6.item_D_name AS FR_step_02_sign_name,";
                    T_SQL = T_SQL + " I7.item_D_name AS FR_step_03_sign_name,I8.item_D_name AS FR_step_HR_sign_name,";
                    T_SQL = T_SQL + " I5.item_D_color AS FR_step_01_color,I6.item_D_color AS FR_step_02_color,";
                    T_SQL = T_SQL + " I7.item_D_color AS FR_step_03_color,I8.item_D_color AS FR_step_HR_color,I9.item_D_name AS FR_kind_show";
                    T_SQL = T_SQL + " FROM Flow_rest FR";
                    T_SQL = T_SQL + " LEFT JOIN User_M U1 ON U1.U_num = FR.FR_U_num AND U1.del_tag = '0'";
                    T_SQL = T_SQL + " LEFT JOIN User_M U2 ON U2.U_num = FR.FR_step_01_num AND U2.del_tag = '0'";
                    T_SQL = T_SQL + " LEFT JOIN User_M U3 ON U3.U_num = FR.FR_step_02_num AND U3.del_tag = '0'";
                    T_SQL = T_SQL + " LEFT JOIN User_M U4 ON U4.U_num = FR.FR_step_03_num AND U4.del_tag = '0'";
                    T_SQL = T_SQL + " LEFT JOIN Item_list I1 ON I1.item_M_code = 'Flow_step_type' AND I1.item_D_type = 'Y' AND I1.item_D_code = FR.FR_step_01_type AND I1.del_tag = '0'";
                    T_SQL = T_SQL + " LEFT JOIN Item_list I2 ON I2.item_M_code = 'Flow_step_type' AND I2.item_D_type = 'Y' AND I2.item_D_code = FR.FR_step_02_type AND I2.del_tag = '0'";
                    T_SQL = T_SQL + " LEFT JOIN Item_list I3 ON I3.item_M_code = 'Flow_step_type' AND I3.item_D_type = 'Y' AND I3.item_D_code = FR.FR_step_03_type AND I3.del_tag = '0'";
                    T_SQL = T_SQL + " LEFT JOIN Item_list I4 ON I4.item_M_code = 'Flow_step_type' AND I4.item_D_type = 'Y' AND I4.item_D_code = FR.FR_step_HR_type AND I4.del_tag = '0'";
                    T_SQL = T_SQL + " LEFT JOIN Item_list I5 ON I5.item_M_code = 'Flow_sign_type' AND I5.item_D_type = 'Y' AND I5.item_D_code = FR.FR_step_01_sign AND I5.del_tag = '0'";
                    T_SQL = T_SQL + " LEFT JOIN Item_list I6 ON I6.item_M_code = 'Flow_sign_type' AND I6.item_D_type = 'Y' AND I6.item_D_code = FR.FR_step_02_sign AND I6.del_tag = '0'";
                    T_SQL = T_SQL + " LEFT JOIN Item_list I7 ON I7.item_M_code = 'Flow_sign_type' AND I7.item_D_type = 'Y' AND I7.item_D_code = FR.FR_step_03_sign AND I7.del_tag = '0'";
                    T_SQL = T_SQL + " LEFT JOIN Item_list I8 ON I8.item_M_code = 'Flow_sign_type' AND I8.item_D_type = 'Y' AND I8.item_D_code = FR.FR_step_HR_sign AND I8.del_tag = '0'";
                    T_SQL = T_SQL + " LEFT JOIN Item_list I9 ON I9.item_M_code = 'FR_kind' AND I9.item_D_type = 'Y' AND I9.item_D_code = FR.FR_kind AND I9.del_tag = '0'";
                    T_SQL = T_SQL + " WHERE FR.del_tag = '0' AND FR.FR_id = @FR_id";
                    parameters.Add(new SqlParameter("@FR_id", Fr_Id));

                    ADOData _adoData = new ADOData();
                    DataTable dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
                    if (dtResult.Rows.Count > 0)
                    {
                        resultClass.ResultCode = "000";
                        resultClass.objResult = JsonConvert.SerializeObject(dtResult);
                    }
                    else
                    {
                        resultClass.ResultCode = "201";
                        resultClass.ResultMsg = "查無資料";
                    }
                }
                return Ok(resultClass);
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                return StatusCode(500, resultClass); // 返回 500 錯誤碼
            }
            
        }
        /// <summary>
        /// 代理人相關資訊 Flow_Rest_leader_Query/Flow_rest_V201803_add.asp
        /// </summary>
        /// <returns></returns>
        [HttpGet("Flow_Rest_leader_Query")]
        public ActionResult<ResultClass<string>> Flow_Rest_leader_Query()
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            var User_Num = HttpContext.Session.GetString("UserID");

            try
            {
                var T_SQL = "SELECT m.U_num AS Member_Num,m.U_name AS Member_Name,m.U_name AS Member_Name,";
                T_SQL = T_SQL + " a.U_num AS Agent_Num,a.U_name AS Agent_Name,l1.U_num AS Leader_1_Num, l1.U_name AS Leader_1_Name,";
                T_SQL = T_SQL + " l2.U_num AS Leader_2_Num,l2.U_name AS Leader_2_Name,l3.U_num AS Leader_3_Num,l3.U_name AS Leader_3_Name";
                T_SQL = T_SQL + " FROM User_M m";
                T_SQL = T_SQL + " LEFT JOIN User_M a ON m.U_agent_num = a.U_num";
                T_SQL = T_SQL + " LEFT JOIN User_M l1 ON m.U_leader_1_num = l1.U_num";
                T_SQL = T_SQL + " LEFT JOIN User_M l2 ON m.U_leader_2_num = l2.U_num";
                T_SQL = T_SQL + " LEFT JOIN User_M l3 ON m.U_leader_3_num = l3.U_num";
                T_SQL = T_SQL + " WHERE m.del_tag = '0' AND m.U_num = @U_num";
                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@U_num", User_Num));

                ADOData _adoData = new ADOData();
                DataTable dtResult = _adoData.ExecuteQuery(T_SQL, parameters);

                if (dtResult.Rows.Count > 0)
                {
                    resultClass.ResultCode = "000";
                    resultClass.objResult = JsonConvert.SerializeObject(dtResult);
                }
                else
                {
                    resultClass.ResultCode = "201";
                    resultClass.ResultMsg = "查無資料";
                }
                return Ok(resultClass);
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                return StatusCode(500, resultClass); // 返回 500 錯誤碼
            }
        }
        /// <summary>
        /// 請假類別 Flow_Rest_kind/Flow_rest_V201803_add.asp
        /// </summary>
        /// <returns></returns>
        [HttpGet("Flow_Rest_kind")]
        public ActionResult<ResultClass<string>> Flow_Rest_kind()
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            try
            {
                var T_SQL = "select  item_D_code,item_D_name from Item_list where item_M_code = 'FR_kind' AND item_D_type='Y' AND show_tag='0' AND del_tag='0' order by item_sort";
                ADOData _adoData = new ADOData();
                DataTable dtResult = _adoData.ExecuteSQuery(T_SQL);
                if (dtResult.Rows.Count > 0)
                {
                    resultClass.ResultCode = "000";
                    resultClass.objResult = JsonConvert.SerializeObject(dtResult);
                }
                else
                {
                    resultClass.ResultCode = "201";
                    resultClass.ResultMsg = "查無資料";
                }
                return Ok(resultClass);
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                return StatusCode(500, resultClass); // 返回 500 錯誤碼
            }
        }
        /// <summary>
        /// 員工名單抓取 Flow_Rest/select_user_one.asp
        /// </summary>
        /// <returns></returns>
        [HttpGet("GetUserMenu")]
        public ActionResult<ResultClass<string>> GetUserMenu() 
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                var T_SQL = "SELECT bc.item_D_name AS U_BC_name,um.U_num,um.U_name,pft.item_D_name AS U_PFT_name";
                T_SQL = T_SQL + " FROM User_M um";
                T_SQL = T_SQL + " LEFT JOIN Item_list bc ON bc.item_M_code = 'branch_company'  AND bc.item_D_code = um.U_BC AND bc.item_D_type = 'Y' AND bc.show_tag = '0' AND bc.del_tag = '0'";
                T_SQL = T_SQL + " LEFT JOIN Item_list pft ON pft.item_M_code = 'professional_title' AND pft.item_D_code = um.U_PFT AND pft.item_D_type = 'Y' AND pft.show_tag = '0' AND pft.del_tag = '0'";
                T_SQL = T_SQL + " WHERE um.del_tag = '0' AND bc.item_D_name is not null AND U_num <> 'AA999'";
                T_SQL = T_SQL + " ORDER BY bc.item_sort,pft.item_sort;";

                ADOData _adoData = new ADOData();
                DataTable dtResult = _adoData.ExecuteSQuery(T_SQL);

                if (dtResult.Rows.Count > 0)
                {
                    resultClass.ResultCode = "000";
                    resultClass.objResult = JsonConvert.SerializeObject(dtResult);
                }
                else
                {
                    resultClass.ResultCode = "201";
                    resultClass.ResultMsg = "查無資料";
                }

                return Ok(resultClass);
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                return StatusCode(500, resultClass); // 返回 500 錯誤碼
            }
        }

        /// <summary>
        /// 請假單單筆新增-Flow_rest/Flow_rest_V201803_add.asp
        /// </summary>
        /// <returns></returns>
        [HttpPost("Flow_Rest_SIns")]
        public ActionResult<ResultClass<string>> Flow_Rest_SIns(Flow_rest model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            var User_Num = HttpContext.Session.GetString("UserID");
            var roleNum = HttpContext.Session.GetString("Role_num");

            try
            {
                //單位主管跟直屬主管同一人時 直屬主管無須簽核
                return Ok(resultClass);
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                return StatusCode(500, resultClass); // 返回 500 錯誤碼
            }
        }
        /// <summary>
        /// 請假單單筆異動-Flow_rest/Flow_rest_list.asp
        /// </summary>
        /// <returns></returns>
        [HttpPost("Flow_Rest_SUpd")]
        public ActionResult<ResultClass<string>> Flow_Rest_SUpd(Flow_rest model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            var User_Num = HttpContext.Session.GetString("UserID");
            var roleNum = HttpContext.Session.GetString("Role_num");

            try
            {
                return Ok(resultClass);
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                return StatusCode(500, resultClass); // 返回 500 錯誤碼
            }
        }
        /// <summary>
        /// 請假單附加檔案查詢-Flow_rest/Flow_rest_list.asp
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [HttpPost("Flow_Rest_UpLoad_Query")]
        public ActionResult<ResultClass<string>> Flow_Rest_UpLoad_Query(Flow_rest model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            var User_Num = HttpContext.Session.GetString("UserID");
            var roleNum = HttpContext.Session.GetString("Role_num");

            try
            {
                return Ok(resultClass);
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                return StatusCode(500, resultClass); // 返回 500 錯誤碼
            }
        }
        
        #endregion
    }

}
