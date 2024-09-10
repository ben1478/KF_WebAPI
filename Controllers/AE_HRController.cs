using KF_WebAPI.BaseClass;
using KF_WebAPI.BaseClass.AE;
using KF_WebAPI.FunctionHandler;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Extensions.Configuration;
using Microsoft.IdentityModel.Tokens;
using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics.Metrics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IdentityModel.Tokens.Jwt;
using System.Net;
using System.Reflection;
using System.Reflection.PortableExecutable;
using System.Security.Claims;
using System.Text;
using System.Text.RegularExpressions;

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
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = "select fr.FR_id,um1.U_name AS FR_U_name,it9.item_D_name AS FR_kind_show,fr.FR_date_begin,";
                T_SQL = T_SQL + " fr.FR_date_end,fr.FR_total_hour,it1.item_D_name AS FR_step_01_type_name,";
                T_SQL = T_SQL + " it2.item_D_name AS FR_step_02_type_name,it3.item_D_name AS FR_step_03_type_name,";
                T_SQL = T_SQL + " it4.item_D_name AS FR_step_HR_type_name,it5.item_D_name AS FR_step_01_sign_name,";
                T_SQL = T_SQL + " it6.item_D_name AS FR_step_02_sign_name,it7.item_D_name AS FR_step_03_sign_name,";
                T_SQL = T_SQL + " it8.item_D_name AS FR_step_HR_sign_name,it10.item_D_name AS FR_sign_type_name,fr.FR_note,fr.FR_cancel,";
                T_SQL = T_SQL + " ISNULL( (SELECT COUNT(*) FROM ASP_UpLoad WHERE cknum = fr.FR_cknum AND del_tag = '0'), '0') AS upload_num";
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
                T_SQL = T_SQL + " where fr.del_tag = '0'";
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
                //請假人員
                if (!string.IsNullOrEmpty(model.Rest_Num))
                {
                    T_SQL = T_SQL + " AND fr.FR_U_num = @Rest_Num";
                    parameters.Add(new SqlParameter("@Rest_Num", model.Rest_Num));
                }
                #endregion
                T_SQL = T_SQL + " ORDER BY fr.FR_date_begin DESC,fr.FR_id";
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
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = "SELECT FR.*,U1.U_name AS FR_U_name,U2.U_name AS FR_step_01_name,U3.U_name AS FR_step_02_name,";
                T_SQL = T_SQL + " U4.U_name AS FR_step_03_name,I1.item_D_name AS FR_step_01_type_name,I2.item_D_name AS FR_step_02_type_name,";
                T_SQL = T_SQL + " I3.item_D_name AS FR_step_03_type_name,I4.item_D_name AS FR_step_HR_type_name,";
                T_SQL = T_SQL + " I5.item_D_name AS FR_step_01_sign_name,I6.item_D_name AS FR_step_02_sign_name,";
                T_SQL = T_SQL + " I7.item_D_name AS FR_step_03_sign_name,I8.item_D_name AS FR_step_HR_sign_name,";
                T_SQL = T_SQL + " I5.item_D_color AS FR_step_01_color,I6.item_D_color AS FR_step_02_color,";
                T_SQL = T_SQL + " I7.item_D_color AS FR_step_03_color,I8.item_D_color AS FR_step_HR_color,I9.item_D_name AS FR_kind_show,";
                T_SQL = T_SQL + " ISNULL( (SELECT COUNT(*) FROM ASP_UpLoad WHERE cknum = FR.FR_cknum AND del_tag = '0'), '0') AS upload_num";
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
                return StatusCode(500, resultClass); // 返回 500 錯誤碼
            }
            
        }
        /// <summary>
        /// 代理人相關資訊 Flow_Rest_leader_Query/Flow_rest_V201803_add.asp
        /// </summary>
        /// <returns></returns>
        [HttpGet("Flow_Rest_leader_Query")]
        public ActionResult<ResultClass<string>> Flow_Rest_leader_Query(string U_num)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = "SELECT m.U_num AS Member_Num,m.U_name AS Member_Name,m.U_name AS Member_Name,";
                T_SQL = T_SQL + " a.U_num AS Agent_Num,a.U_name AS Agent_Name,l1.U_num AS Leader_1_Num, l1.U_name AS Leader_1_Name,";
                T_SQL = T_SQL + " l2.U_num AS Leader_2_Num,l2.U_name AS Leader_2_Name,l3.U_num AS Leader_3_Num,l3.U_name AS Leader_3_Name";
                T_SQL = T_SQL + " FROM User_M m";
                T_SQL = T_SQL + " LEFT JOIN User_M a ON m.U_agent_num = a.U_num";
                T_SQL = T_SQL + " LEFT JOIN User_M l1 ON m.U_leader_1_num = l1.U_num";
                T_SQL = T_SQL + " LEFT JOIN User_M l2 ON m.U_leader_2_num = l2.U_num";
                T_SQL = T_SQL + " LEFT JOIN User_M l3 ON m.U_leader_3_num = l3.U_num";
                T_SQL = T_SQL + " WHERE m.del_tag = '0' AND m.U_num = @U_num";
                parameters.Add(new SqlParameter("@U_num", U_num));
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
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = "select  item_D_code,item_D_name from Item_list where item_M_code = 'FR_kind' AND item_D_type='Y' AND show_tag='0' AND del_tag='0' order by item_sort";
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
                return StatusCode(500, resultClass); // 返回 500 錯誤碼
            }
        }
        /// <summary>
        /// 抓取特休資料 Flow_Rest_kind/Flow_rest_V201803_add.asp
        /// <summary>
        [HttpGet("GetHDay")]
        public ActionResult<ResultClass<string>> GetHDay()
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            var User_Num = HttpContext.Session.GetString("UserID");
            DateTime BeginDate = new DateTime(DateTime.Today.Year, DateTime.Today.Month, DateTime.Today.Day, 0, 0, 0, 0);
            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = "select H_day_total,H_begin,H_end from User_Hday where del_tag = '0' and U_num=@U_num and H_begin <= @BeginDate order by H_begin desc";
                parameters.Add(new SqlParameter("@U_num", User_Num));
                parameters.Add(new SqlParameter("@BeginDate", BeginDate));
                #endregion

                DataTable dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
                //抓取特休資料
                if (dtResult.Rows.Count > 0)
                {
                    User_Hday_res model = new User_Hday_res();
                    if (dtResult.Rows.Count == 1)
                    {
                        DataRow row = dtResult.Rows[0];
                        model.H_day_total_now = Convert.ToInt32(row["H_day_total"]);
                        model.H_begin_now = Convert.ToDateTime(row["H_begin"]);
                        model.H_end_now = Convert.ToDateTime(row["H_end"]);

                        #region SQL
                        var parameters1 = new List<SqlParameter>();
                        var T_SQL_1 = "select sum(case when fr.FR_kind='FRK005' AND FR_date_begin <= @FR_date_end_now AND FR_date_end >= @FR_date_begin_now then FR_total_hour else 0 end) as 'FR_kind_FRK005'";
                        T_SQL_1 = T_SQL_1 + " FROM User_M u";
                        T_SQL_1 = T_SQL_1 + " left join  Flow_rest fr on fr.FR_U_num=u.U_num AND fr.del_tag='0' AND fr.FR_sign_type in ('FSIGN001','FSIGN002') AND fr.FR_kind='FRK005' AND fr.FR_cancel='N'";
                        T_SQL_1 = T_SQL_1 + " where u.del_tag='0' AND u.U_num=@U_num";
                        parameters1.Add(new SqlParameter("@U_num", User_Num));
                        parameters1.Add(new SqlParameter("@FR_date_begin_now", model.H_begin_now));
                        parameters1.Add(new SqlParameter("@FR_date_end_now", model.H_end_now));
                        #endregion

                        DataTable dtResult1 = _adoData.ExecuteQuery(T_SQL_1, parameters1);
                        DataRow row1 = dtResult1.Rows[0];
                        model.FR_kind_FRK005_now= Convert.ToInt32(row1["FR_kind_FRK005"]);

                        resultClass.ResultCode = "000";
                        resultClass.objResult = JsonConvert.SerializeObject(model);
                        

                    }
                    else
                    {
                        DataRow row = dtResult.Rows[0];
                        DataRow row_last = dtResult.Rows[1];

                        model.H_day_total_now = Convert.ToInt32(row["H_day_total"]);
                        model.H_begin_now = Convert.ToDateTime(row["H_begin"]);
                        model.H_end_now = Convert.ToDateTime(row["H_end"]);
                        model.H_day_total_last = Convert.ToInt32(row_last["H_day_total"]);
                        model.H_begin_last = Convert.ToDateTime(row_last["H_begin"]);
                        model.H_end_last = Convert.ToDateTime(row_last["H_end"]);

                        #region SQL
                        var parameters1 = new List<SqlParameter>();
                        var T_SQL_1 = "select sum(case when fr.FR_kind='FRK005' AND FR_date_begin <= @FR_date_end_now AND FR_date_end >= @FR_date_begin_now then FR_total_hour else 0 end) as 'FR_kind_FRK005'";
                        T_SQL_1 = T_SQL_1 + " ,sum(case when fr.FR_kind='FRK005' AND FR_date_begin <= @FR_date_end_last AND FR_date_end >= @FR_date_begin_last then FR_total_hour else 0 end) as 'FR_kind_FRK005_year'";
                        T_SQL_1 = T_SQL_1 + " FROM User_M u";
                        T_SQL_1 = T_SQL_1 + " left join  Flow_rest fr on fr.FR_U_num=u.U_num AND fr.del_tag='0' AND fr.FR_sign_type in ('FSIGN001','FSIGN002') AND fr.FR_kind='FRK005' AND fr.FR_cancel='N'";
                        T_SQL_1 = T_SQL_1 + " where u.del_tag='0' AND u.U_num=@U_num";
                        parameters1.Add(new SqlParameter("@U_num", User_Num));
                        parameters1.Add(new SqlParameter("@FR_date_begin_now", model.H_begin_now));
                        parameters1.Add(new SqlParameter("@FR_date_end_now", model.H_end_now));
                        parameters1.Add(new SqlParameter("@FR_date_begin_last", model.H_begin_last));
                        parameters1.Add(new SqlParameter("@FR_date_end_last", model.H_end_last));
                        #endregion

                        DataTable dtResult1 = _adoData.ExecuteQuery(T_SQL_1, parameters1);
                        DataRow row1 = dtResult1.Rows[0];
                        model.FR_kind_FRK005_now = Convert.ToInt32(row1["FR_kind_FRK005"]);
                        model.FR_kind_FRK005_last = Convert.ToInt32(row1["FR_kind_FRK005_year"]);

                        resultClass.ResultCode = "000";
                        resultClass.objResult = JsonConvert.SerializeObject(model);                       
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
        /// 員工名單抓取 Flow_Rest/select_user_one.asp
        /// </summary>
        /// <returns></returns>
        [HttpGet("GetUserMenu")]
        public ActionResult<ResultClass<string>> GetUserMenu() 
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = "SELECT bc.item_D_name AS U_BC_name,um.U_num,um.U_name,pft.item_D_name AS U_PFT_name";
                T_SQL = T_SQL + " FROM User_M um";
                T_SQL = T_SQL + " LEFT JOIN Item_list bc ON bc.item_M_code = 'branch_company'  AND bc.item_D_code = um.U_BC AND bc.item_D_type = 'Y' AND bc.show_tag = '0' AND bc.del_tag = '0'";
                T_SQL = T_SQL + " LEFT JOIN Item_list pft ON pft.item_M_code = 'professional_title' AND pft.item_D_code = um.U_PFT AND pft.item_D_type = 'Y' AND pft.show_tag = '0' AND pft.del_tag = '0'";
                T_SQL = T_SQL + " WHERE um.del_tag = '0' AND bc.item_D_name is not null AND U_num <> 'AA999'";
                T_SQL = T_SQL + " ORDER BY bc.item_sort,pft.item_sort;";
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
            
            try
            {
                ADOData _adoData = new ADOData();
                #region 特休天數檢查
                if (model.FR_kind == "FRK005")
                {
                    DateTime BeginDate = new DateTime(DateTime.Today.Year, DateTime.Today.Month, DateTime.Today.Day, 0, 0, 0, 0);

                    #region SQL
                    var parameters_hay = new List<SqlParameter>();
                    var T_SQL_Hay = "select H_day_total,H_begin,H_end from User_Hday where del_tag = '0' and U_num=@U_num and H_begin <= @BeginDate order by H_begin desc";
                    parameters_hay.Add(new SqlParameter("@U_num", User_Num));
                    parameters_hay.Add(new SqlParameter("@BeginDate", BeginDate));
                    #endregion

                    DataTable dtResultHay = _adoData.ExecuteQuery(T_SQL_Hay, parameters_hay);
                    DataRow rowhay = dtResultHay.Rows[0];

                    #region SQL
                    var parameters_hay_imp = new List<SqlParameter>();
                    var T_SQL_Hay_Imp = "select sum(case when fr.FR_kind='FRK005' AND FR_date_begin <= @FR_date_end_now AND FR_date_end >= @FR_date_begin_now then FR_total_hour else 0 end) as 'FR_kind_FRK005'";
                    T_SQL_Hay_Imp = T_SQL_Hay_Imp + " FROM User_M u";
                    T_SQL_Hay_Imp = T_SQL_Hay_Imp + " left join  Flow_rest fr on fr.FR_U_num=u.U_num AND fr.del_tag='0' AND fr.FR_sign_type in ('FSIGN001','FSIGN002') AND fr.FR_kind='FRK005' AND fr.FR_cancel='N'";
                    T_SQL_Hay_Imp = T_SQL_Hay_Imp + " where u.del_tag='0' AND u.U_num=@U_num";
                    parameters_hay_imp.Add(new SqlParameter("@U_num", User_Num));
                    parameters_hay_imp.Add(new SqlParameter("@FR_date_begin_now", Convert.ToDateTime(rowhay["H_begin"])));
                    parameters_hay_imp.Add(new SqlParameter("@FR_date_end_now", Convert.ToDateTime(rowhay["H_end"])));
                    #endregion

                    DataTable dtResultHayImp = _adoData.ExecuteQuery(T_SQL_Hay_Imp, parameters_hay_imp);
                    DataRow rowhayimp = dtResultHayImp.Rows[0];

                    var hay = Convert.ToInt32(rowhay["H_day_total"]) * 8;//可休時數
                    var implement_hay = Convert.ToInt32(rowhayimp["FR_kind_FRK005"]);//已休時數
                    decimal FR_total_hour = model.FR_total_hour;
                    if ((hay - implement_hay) < FR_total_hour)
                    {
                        resultClass.ResultCode = "400";
                        resultClass.ResultMsg = "特休時數不夠";
                        return BadRequest(resultClass);
                    }
                }

                #endregion

                var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();
                var Fun = new FuncHandler();
                string CheckNum = Fun.GetCheckNum();

                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = "INSERT INTO Flow_Rest(FR_cknum,FR_version,FR_kind,FR_date_begin,FR_date_end,FR_total_hour,FR_note,FR_sign_type,";
                T_SQL = T_SQL + " FR_cancel,FR_U_num,FR_step_now,FR_step_01_type,FR_step_01_num,FR_step_01_sign,FR_step_02_type,FR_step_02_num,";
                T_SQL = T_SQL + " FR_step_02_sign,FR_step_03_type,FR_step_03_num,FR_step_03_sign,";
                T_SQL = T_SQL + " FR_step_04_type,FR_step_04_sign,FR_step_05_type,FR_step_05_sign,";
                T_SQL = T_SQL + " FR_step_HR_type,FR_step_HR_num,FR_step_HR_sign,add_date,add_num,add_ip,del_tag )";
                T_SQL = T_SQL + " VALUES ( @FR_cknum,@FR_version,@FR_kind,@FR_date_begin,@FR_date_end,@FR_total_hour,@FR_note,@FR_sign_type,";
                T_SQL = T_SQL + " @FR_cancel,@FR_U_num,@FR_step_now,@FR_step_01_type,@FR_step_01_num,@FR_step_01_sign,@FR_step_02_type,";
                T_SQL = T_SQL + " @FR_step_02_num,@FR_step_02_sign,@FR_step_03_type,@FR_step_03_num,@FR_step_03_sign,@FR_step_04_type,";
                T_SQL = T_SQL + " @FR_step_04_sign,@FR_step_05_type,@FR_step_05_sign,@FR_step_HR_type,@FR_step_HR_num,@FR_step_HR_sign,";
                T_SQL = T_SQL + " @add_date,@add_num,@add_ip,@del_tag )";
                parameters.Add(new SqlParameter("@FR_cknum", CheckNum));
                parameters.Add(new SqlParameter("@FR_version", "V201803"));
                parameters.Add(new SqlParameter("@FR_kind", model.FR_kind));
                parameters.Add(new SqlParameter("@FR_date_begin", model.FR_date_begin));
                parameters.Add(new SqlParameter("@FR_date_end", model.FR_date_end));
                parameters.Add(new SqlParameter("@FR_total_hour", model.FR_total_hour));
                parameters.Add(new SqlParameter("@FR_note", model.FR_note));
                parameters.Add(new SqlParameter("@FR_sign_type", "FSIGN001"));
                parameters.Add(new SqlParameter("@FR_cancel", "N"));
                parameters.Add(new SqlParameter("@FR_U_num", User_Num));
                //公出 FRK016 忘打卡 FRK017 無須代理人
                if (model.FR_kind == "FRK016" || model.FR_kind == "FRK017")
                {
                    parameters.Add(new SqlParameter("@FR_step_now", "2"));
                    parameters.Add(new SqlParameter("@FR_step_01_type", "FSTEP003"));
                    parameters.Add(new SqlParameter("@FR_step_01_num", DBNull.Value));
                    parameters.Add(new SqlParameter("@FR_step_HR_type", "FSTEP003"));
                }
                else
                {
                    parameters.Add(new SqlParameter("@FR_step_now", "1"));
                    parameters.Add(new SqlParameter("@FR_step_01_type", "FSTEP001"));
                    parameters.Add(new SqlParameter("@FR_step_01_num", model.FR_step_01_num));
                    parameters.Add(new SqlParameter("@FR_step_HR_type", "FSTEP001"));
                }
                parameters.Add(new SqlParameter("@FR_step_01_sign", "FSIGN001"));
                parameters.Add(new SqlParameter("@FR_step_02_type", "FSTEP001"));
                parameters.Add(new SqlParameter("@FR_step_02_num", model.FR_step_02_num));
                parameters.Add(new SqlParameter("@FR_step_02_sign", "FSIGN001"));
                //超過3天要直屬主管簽核
                if (model.FR_total_hour >= 24)
                {
                    parameters.Add(new SqlParameter("@FR_step_03_num", model.FR_step_03_num));
                    parameters.Add(new SqlParameter("@FR_step_03_type", "FSTEP001"));
                }
                else
                {
                    parameters.Add(new SqlParameter("@FR_step_03_num", DBNull.Value));
                    parameters.Add(new SqlParameter("@FR_step_03_type", "FSTEP003"));
                }
                parameters.Add(new SqlParameter("@FR_step_03_sign", "FSIGN001"));
                parameters.Add(new SqlParameter("@FR_step_04_type", "FSTEP003"));
                parameters.Add(new SqlParameter("@FR_step_04_sign", "FSIGN001"));
                parameters.Add(new SqlParameter("@FR_step_05_type", "FSTEP003"));
                parameters.Add(new SqlParameter("@FR_step_05_sign", "FSIGN001"));
                parameters.Add(new SqlParameter("@FR_step_HR_num", ""));
                parameters.Add(new SqlParameter("@FR_step_HR_sign", "FSIGN001"));
                parameters.Add(new SqlParameter("@add_date", DateTime.Now));
                parameters.Add(new SqlParameter("@add_num", User_Num));
                parameters.Add(new SqlParameter("@add_ip", clientIp));
                parameters.Add(new SqlParameter("@del_tag", "0"));
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

            try
            {
                var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = "Update Flow_rest set FR_kind=@FR_kind,FR_date_begin=@FR_date_begin,FR_date_end=@FR_date_end,";
                T_SQL = T_SQL + " FR_total_hour=@FR_total_hour,FR_note=@FR_note,FR_step_01_num=@FR_step_01_num,FR_step_02_num=@FR_step_02_num,";
                T_SQL = T_SQL + " FR_step_03_num=@FR_step_03_num,edit_date=GETDATE(),edit_num=@edit_num,edit_ip=@IP";
                T_SQL = T_SQL + " where FR_id=@FR_id";
                parameters.Add(new SqlParameter("@FR_kind", model.FR_kind));
                parameters.Add(new SqlParameter("@FR_date_begin", model.FR_date_begin));
                parameters.Add(new SqlParameter("@FR_date_end", model.FR_date_end));
                parameters.Add(new SqlParameter("@FR_total_hour", model.FR_total_hour));
                parameters.Add(new SqlParameter("@FR_note", model.FR_note));
                if (model.FR_step_01_num == null)
                {
                    parameters.Add(new SqlParameter("@FR_step_01_num", DBNull.Value));
                }
                else
                {
                    parameters.Add(new SqlParameter("@FR_step_01_num", model.FR_step_01_num));
                }
                if (model.FR_step_02_num == null)
                {
                    parameters.Add(new SqlParameter("@FR_step_02_num", DBNull.Value));
                }
                else
                {
                    parameters.Add(new SqlParameter("@FR_step_02_num", model.FR_step_02_num));
                }
                if (model.FR_step_03_num == null)
                {
                    parameters.Add(new SqlParameter("@FR_step_03_num", DBNull.Value));
                }
                else
                {
                    parameters.Add(new SqlParameter("@FR_step_03_num", model.FR_step_03_num));
                }
                parameters.Add(new SqlParameter("@edit_num", User_Num));
                parameters.Add(new SqlParameter("@IP", clientIp));
                parameters.Add(new SqlParameter("@FR_id", model.FR_id));
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
                return StatusCode(500, resultClass); // 返回 500 錯誤碼
            }
        }
        /// <summary>
        /// 請假單抽單 Flow_Rest_Cacnel/Flow_rest_list.asp
        /// </summary>
        /// <returns></returns>
        [HttpPost("Flow_Rest_Cacnel")]
        public ActionResult<ResultClass<string>> Flow_Rest_Cacnel(string fr_id)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            var User_Num = HttpContext.Session.GetString("UserID");

            try
            {
                var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();
                var Fun = new FuncHandler();

                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = "Update Flow_rest set FR_cancel='Y',cancel_date=GETDATE(),cancel_num=@User_Num,cancel_ip=@IP where FR_id=@FR_id";
                parameters.Add(new SqlParameter("@User_Num", User_Num));
                parameters.Add(new SqlParameter("@IP", clientIp));
                parameters.Add(new SqlParameter("@FR_id", fr_id));
                #endregion

                int result = _adoData.ExecuteNonQuery(T_SQL, parameters);
                if (result == 0)
                {
                    resultClass.ResultCode = "400";
                    resultClass.ResultMsg = "取消失敗";
                    return BadRequest(resultClass);
                }
                else
                {
                    resultClass.ResultCode = "000";
                    resultClass.ResultMsg = "取消成功";
                    return Ok(resultClass);
                }
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                return StatusCode(500, resultClass); // 返回 500 錯誤碼
            }
        }
        /// <summary>
        /// 請假單刪除 Flow_Rest_Del/Flow_rest_list.asp
        /// </summary>
        /// <returns></returns>
        [HttpPost("Flow_Rest_Del")]
        public ActionResult<ResultClass<string>> Flow_Rest_Del(string fr_id)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            var User_Num = HttpContext.Session.GetString("UserID");

            try
            {
                
                var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = "Update Flow_rest set del_tag='1',del_date=GETDATE(),del_num=@User_Num,cancel_ip=@IP where FR_id=@FR_id";
                parameters.Add(new SqlParameter("@User_Num", User_Num));
                parameters.Add(new SqlParameter("@IP", clientIp));
                parameters.Add(new SqlParameter("@FR_id", fr_id));
                #endregion

                int result = _adoData.ExecuteNonQuery(T_SQL, parameters);

                if (result == 0)
                {
                    resultClass.ResultCode = "400";
                    resultClass.ResultMsg = "刪除失敗";
                    return BadRequest(resultClass);
                }
                else
                {
                    resultClass.ResultCode = "000";
                    resultClass.ResultMsg = "刪除成功";
                    return Ok(resultClass);
                }
                
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                return StatusCode(500, resultClass); // 返回 500 錯誤碼
            }
        }
        /// <summary>
        /// 請假單審核 Flow_Rest_Sign/Flow_rest_V201803_signDB.asp
        /// </summary>
        /// <returns></returns>
        [HttpPost("Flow_Rest_Sign")]
        public ActionResult<ResultClass<string>> Flow_Rest_Sign(Flow_rest model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = "Update Flow_rest set ";
                switch (model.FR_step_now)
                {
                    //代理人
                    case "1":
                        T_SQL = T_SQL + " FR_step_01_type='FSTEP002',FR_step_01_sign=@FR_step_01_sign,FR_step_01_date=Getdate()";
                        T_SQL = T_SQL + " ,FR_step_01_note=@FR_step_01_note";
                        if (model.FR_step_01_sign == "FSIGN002")
                        {
                            T_SQL = T_SQL + " ,FR_step_now='2'";
                        }
                        else
                        {
                            T_SQL = T_SQL + " ,FR_sign_type='FSIGN003',FR_step_now='0'";
                        }
                        parameters.Add(new SqlParameter("@FR_step_01_sign", model.FR_step_01_sign));
                        parameters.Add(new SqlParameter("@FR_step_01_note", model.FR_step_01_note));
                        break;
                    //直屬主管
                    case "2":
                        T_SQL = T_SQL + " FR_step_02_type='FSTEP002',FR_step_02_sign=@FR_step_02_sign,FR_step_02_date=Getdate()";
                        T_SQL = T_SQL + " ,FR_step_02_note=@FR_step_02_note";
                        //第三關屬於免簽核(3天內的假單)的直接到人資 另公出 FRK016 忘打卡 FRK017可直接結案 無須到人資
                        if (model.FR_step_02_sign == "FSIGN002")
                        {
                            if (model.FR_step_03_type == "FSTEP001")
                            {
                                T_SQL = T_SQL + " ,FR_step_now='3'";
                            }
                            else
                            {
                                if (model.FR_kind == "FRK016" || model.FR_kind == "FRK017")
                                {
                                    T_SQL = T_SQL + " ,FR_sign_type='FSIGN002',FR_step_now='0'";
                                }
                                else
                                {
                                    T_SQL = T_SQL + " ,FR_step_now='9'";
                                }
                            }
                        }
                        else
                        {
                            T_SQL = T_SQL + " ,FR_sign_type='FSIGN003',FR_step_now='0'";
                        }
                        parameters.Add(new SqlParameter("@FR_step_02_sign", model.FR_step_02_sign));
                        parameters.Add(new SqlParameter("@FR_step_02_note", model.FR_step_02_note));
                        break;
                    //單位主管
                    case "3":
                        T_SQL = T_SQL + " FR_step_03_type='FSTEP002',FR_step_03_sign=@FR_step_03_sign,FR_step_03_date=Getdate()";
                        T_SQL = T_SQL + " ,FR_step_03_note=@FR_step_03_note";
                        if (model.FR_step_03_sign == "FSIGN002")
                        {
                            T_SQL = T_SQL + " ,FR_step_now='9'";
                        }
                        else
                        {
                            T_SQL = T_SQL + " ,FR_sign_type='FSIGN003',FR_step_now='0'";
                        }
                        parameters.Add(new SqlParameter("@FR_step_03_sign", model.FR_step_03_sign));
                        parameters.Add(new SqlParameter("@FR_step_03_note", model.FR_step_03_note));
                        break;
                    //人資
                    case "9":
                        T_SQL = T_SQL + " FR_step_HR_type='FSTEP002',FR_step_HR_sign=@FR_step_HR_sign,FR_step_HR_date=Getdate()";
                        T_SQL = T_SQL + " ,FR_step_HR_note=@FR_step_HR_note,FR_step_now='0',FR_sign_type=@FR_step_HR_sign";
                        parameters.Add(new SqlParameter("@FR_step_HR_sign", model.FR_step_HR_sign));
                        parameters.Add(new SqlParameter("@FR_step_HR_note", model.FR_step_HR_note));
                        break;
                }

                T_SQL = T_SQL + " Where FR_id=@FR_id";
                parameters.Add(new SqlParameter("@FR_id", model.FR_id));
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
                return StatusCode(500, resultClass); // 返回 500 錯誤碼
            }
        }
        #endregion

        #region 請假單查詢
        /// <summary>
        /// 請假單查詢 Flow_rest_list_search/Flow_rest_list_search.asp
        /// </summary>
        /// <returns></returns>
        [HttpPost("Flow_rest_list_search")]
        public ActionResult<ResultClass<string>> Flow_rest_list_search(Flow_rest_req model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            var User_Num = HttpContext.Session.GetString("UserID");
            var roleNum = HttpContext.Session.GetString("Role_num");

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = "select fr.FR_id,um1.U_name AS FR_U_name,it9.item_D_name AS FR_kind_show,fr.FR_date_begin,";
                T_SQL = T_SQL + " fr.FR_date_end,fr.FR_total_hour,it1.item_D_name AS FR_step_01_type_name,";
                T_SQL = T_SQL + " it2.item_D_name AS FR_step_02_type_name,it3.item_D_name AS FR_step_03_type_name,";
                T_SQL = T_SQL + " it4.item_D_name AS FR_step_HR_type_name,it5.item_D_name AS FR_step_01_sign_name,";
                T_SQL = T_SQL + " it6.item_D_name AS FR_step_02_sign_name,it7.item_D_name AS FR_step_03_sign_name,";
                T_SQL = T_SQL + " it8.item_D_name AS FR_step_HR_sign_name,it10.item_D_name AS FR_sign_type_name,fr.FR_note,fr.FR_cancel,";
                T_SQL = T_SQL + " ISNULL( (SELECT COUNT(*) FROM ASP_UpLoad WHERE cknum = fr.FR_cknum AND del_tag = '0'), '0') AS upload_num";
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
                T_SQL = T_SQL + " where fr.del_tag = '0'";
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
                //請假人員
                if (!string.IsNullOrEmpty(model.Rest_Num))
                {
                    T_SQL = T_SQL + " AND fr.FR_U_num = @Rest_Num";
                    parameters.Add(new SqlParameter("@Rest_Num", model.Rest_Num));
                }
                #endregion
                T_SQL = T_SQL + " ORDER BY fr.FR_date_begin DESC,fr.FR_id";
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
                return StatusCode(500, resultClass); // 返回 500 錯誤碼
            }
        }

        #endregion

        #region 出勤紀錄上傳
        /// <summary>
        /// 出勤紀錄上傳 Attendance_Upload/attendance_upload.asp&Upload_Exel.asp
        /// </summary>
        [HttpPost("Attendance_Upload")]
        public ActionResult<ResultClass<string>> Attendance_Upload(IFormFile file)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            if (!file.FileName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                resultClass.ResultMsg = "檔案格式錯誤。請上傳 .xlsx 格式的檔案。";
                return BadRequest(resultClass);
            }

            try
            {
                ADOData _adoData = new ADOData();

                #region 1.檔案存入
                string _storagePath = @"C:\UploadedFiles";
                var filePath = Path.Combine(_storagePath, file.FileName);
                using (var stream = new FileStream(filePath, FileMode.Create))
                {
                    file.CopyTo(stream);
                }
                #endregion

                #region 2.讀取檔案
                var attendances = new List<Attendance>();
                var yyyymm = Path.GetFileNameWithoutExtension(file.FileName).Replace("-", "");
                var inputdate = DateTime.Now;
                // 使用 EPPlus 讀取 Excel 文件
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    foreach (var worksheet in package.Workbook.Worksheets)
                    {
                        var user_apart = "";
                        var user_name = "";
                        var userID = "";
                        var user_num = worksheet.Name.Split(" ")[1];
                        for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                        {
                            var dateCellValue = worksheet.Cells[row, 3].Text;
                            if (dateCellValue == "時數")
                            {
                                break; // 退出循環
                            }

                            user_apart = worksheet.Cells[2, 1].Text;
                            user_name = worksheet.Cells[2, 2].Text;
                            if (row == 2)
                            {
                                //抓取員編&比對員工難字對應表item_M_code='SpecName'&比對出勤人員 item_M_code='NonUser'
                                #region SQL
                                var parameters = new List<SqlParameter>();
                                var T_SQL = "SELECT u_num, u_name FROM user_M Where del_tag='0' and u_name=@u_name";
                                parameters.Add(new SqlParameter("@u_name", user_name));
                                #endregion
                                DataTable dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
                                if (dtResult.Rows.Count > 0)
                                {
                                    DataRow nser_row = dtResult.Rows[0];
                                    userID = nser_row["u_num"].ToString();
                                    user_name = nser_row["u_name"].ToString();
                                }
                                else
                                {
                                    #region SQL
                                    var parameters_sp = new List<SqlParameter>();
                                    var T_SQL_SP = "SELECT u_num, u_name FROM user_M M JOIN ( SELECT item_D_code, item_D_name";
                                    T_SQL_SP = T_SQL_SP + " FROM dbo.Item_list WHERE item_M_code = 'SpecName' AND item_D_type = 'Y' AND item_D_name = @name_sp";
                                    T_SQL_SP = T_SQL_SP + " ) I ON M.u_name = item_D_code";
                                    parameters_sp.Add(new SqlParameter("@name_sp", worksheet.Cells[2, 2].Text));
                                    #endregion
                                    DataTable dtResult_sp = _adoData.ExecuteQuery(T_SQL_SP, parameters_sp);
                                    if (dtResult_sp.Rows.Count > 0)
                                    {
                                        DataRow nser_row_sp = dtResult_sp.Rows[0];
                                        userID = nser_row_sp["u_num"].ToString();
                                        user_name = nser_row_sp["u_name"].ToString();
                                    }
                                    else
                                    {
                                        #region SQL
                                        var parameters_nop = new List<SqlParameter>();
                                        var T_SQL_NOP = "SELECT item_D_code, item_D_name FROM dbo.Item_list";
                                        T_SQL_NOP = T_SQL_NOP + " WHERE item_M_code = 'NonUser' AND item_D_type = 'Y' AND item_D_code = @name_nop";
                                        parameters_nop.Add(new SqlParameter("@name_nop", worksheet.Cells[2, 2].Text));
                                        #endregion
                                        DataTable dtResult_nop = _adoData.ExecuteQuery(T_SQL_NOP, parameters_nop);
                                        if (dtResult_nop.Rows.Count > 0)
                                        {
                                            DataRow nser_row_nop = dtResult_nop.Rows[0];
                                            userID = "";
                                            user_name = nser_row_nop["item_D_code"].ToString();
                                        }
                                        else
                                        {
                                            resultClass.ResultMsg = "沒有對應的員工姓名,該人員可能已離職,如未離職,請維護[開發人員設定]各式選項管理:員工姓名難字";
                                            return BadRequest(resultClass);
                                        }
                                    }
                                }
                            }
                            var attendance_date = worksheet.Cells[row, 3].Text.Split(' ')[0];
                            var work_time = worksheet.Cells[row, 10].Text;
                            var getoffwork_time = worksheet.Cells[row, 11].Text;

                            var attendance = new Attendance
                            {
                                user_name = user_name,
                                userID = userID,
                                yyyymm = yyyymm,
                                user_apart = user_apart,
                                attendance_date = attendance_date,
                                work_time = work_time,
                                getoffwork_time = getoffwork_time,
                                inputdate = inputdate,
                                user_num = user_num
                            };
                            attendances.Add(attendance);
                        }
                    }
                }
                #endregion

                #region 3.資料處理
                #region SQL
                var T_SQL_DEL = "delete attendance where yyyymm=@yyyymm";
                var parameters_DEL = new List<SqlParameter>();
                parameters_DEL.Add(new SqlParameter("@yyyymm", yyyymm));
                #endregion
                int Result = _adoData.ExecuteNonQuery(T_SQL_DEL, parameters_DEL);
                if (Result == 0)
                {
                    resultClass.ResultMsg = "資料刪除失敗";
                    return BadRequest(resultClass);
                }
                foreach (var item in attendances)
                {
                    #region SQL
                    var T_SQL_IN = "Insert into attendance (user_name,userID,yyyymm,user_apart,attendance_date,work_time,getoffwork_time,inputdate,user_num)";
                    T_SQL_IN = T_SQL_IN + "Values (@user_name,@userID,@yyyymm,@user_apart,@attendance_date,@work_time,@getoffwork_time,@inputdate,@user_num)";
                    var parameters_IN = new List<SqlParameter>();
                    parameters_IN.Add(new SqlParameter("@user_name", item.user_name));
                    parameters_IN.Add(new SqlParameter("@userID", item.userID));
                    parameters_IN.Add(new SqlParameter("@yyyymm", item.yyyymm));
                    parameters_IN.Add(new SqlParameter("@user_apart", item.user_apart));
                    parameters_IN.Add(new SqlParameter("@attendance_date", item.attendance_date));
                    parameters_IN.Add(new SqlParameter("@work_time", item.work_time));
                    parameters_IN.Add(new SqlParameter("@getoffwork_time", item.getoffwork_time));
                    parameters_IN.Add(new SqlParameter("@inputdate", item.inputdate));
                    parameters_IN.Add(new SqlParameter("@user_num", item.user_num));
                    #endregion
                    int ResultIn = _adoData.ExecuteNonQuery(T_SQL_IN, parameters_IN);
                    if (ResultIn == 0)
                    {
                        resultClass.ResultMsg = "資料新增失敗";
                        return BadRequest(resultClass);
                    }
                }
                #endregion
                resultClass.ResultMsg = "資料新增成功";
                return Ok(resultClass);
            }
            catch (Exception ex)
            {
                resultClass.ResultMsg = ex.Message;
                resultClass.ResultCode = "500";
                return StatusCode(500, resultClass); 
            }
        }
        #endregion

        #region 個人出勤紀錄
        /// <summary>
        /// 個人出勤紀錄年月範圍 Attendance_Query/Attendance_report.asp?Self=Y
        /// </summary>
        [HttpGet("Attendance_YYYYMM")]
        public ActionResult<ResultClass<string>> Attendance_YYYYMM()
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = "select distinct yyyymm from [attendance] order by yyyymm desc";
                #endregion
                DataTable dtResult = _adoData.ExecuteSQuery(T_SQL);
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
                resultClass.ResultMsg = ex.Message;
                resultClass.ResultCode = "500";
                return StatusCode(500, resultClass);
            }
        }

        /// <summary>
        /// 個人出勤紀錄 Attendance_Query/Attendance_report.asp?Self=Y
        /// </summary>
        [HttpPost("Attendance_SQuery")]
        public ActionResult<ResultClass<string>> Attendance_SQuery(Attendance_req model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            var User_Num = HttpContext.Session.GetString("UserID");

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = "SELECT U.U_Na,[userID],U_name [user_name],@yyyy + ad.[attendance_date] attendance_date,[work_time],";
                T_SQL = T_SQL + " case When isnull([work_time], '') = '' then 0 When [work_time] > '09:00' then DATEDIFF(MINUTE, '09:00', [work_time]) else 0 end Late,";
                T_SQL = T_SQL + " case When isnull([work_time], '') = '' then '未刷卡' When [work_time] > '09:00' then case when isnull(U_num_NL, 'N') = 'N' then '遲到'  else '' end else '' end work_status,";
                T_SQL = T_SQL + " [getoffwork_time],";
                T_SQL = T_SQL + " case When isnull([getoffwork_time], '') = '' then 0 When [getoffwork_time] < '18:00' then DATEDIFF(MINUTE, [getoffwork_time], '18:00') else 0 end early,";
                T_SQL = T_SQL + " case When isnull([getoffwork_time], '') = '' then '未刷卡' When [getoffwork_time] < '18:00' then case when isnull(U_num_NL, 'N') = 'N' then '早退' else '' end else '' end offwork_status,";
                T_SQL = T_SQL + " U_BC,isnull(RestCount, 0) RestCount";
                T_SQL = T_SQL + " FROM attendance ad";
                T_SQL = T_SQL + " left join ( SELECT U_PFT,U_BC,U_num,U_name,I.item_D_name U_Na from User_M U";
                T_SQL = T_SQL + " left join ( SELECT [item_D_code],[item_D_name] FROM Item_list WHERE item_M_code = 'branch_company'";
                T_SQL = T_SQL + " AND item_D_type = 'Y' and del_tag = '0'  ) I on U.u_bc = I.item_D_code";
                T_SQL = T_SQL + " where del_tag = '0' ) U on ad.userID = U.U_num";
                T_SQL = T_SQL + " Left Join ( SELECT [item_D_code] U_num_NL FROM Item_list where item_M_code = 'NonLate'";
                T_SQL = T_SQL + " and [item_M_type] = 'N' ) NL on U.U_num = NL.U_num_NL";
                T_SQL = T_SQL + " left join ( select FR_U_num,convert(varchar, FR_date_begin, 111) FR_date_S,";
                T_SQL = T_SQL + " convert(varchar, FR_date_end, 111) FR_date_E,count(FR_U_num) RestCount";
                T_SQL = T_SQL + " from Flow_rest where  del_tag = '0' and FR_cancel <> 'Y' group by";
                T_SQL = T_SQL + " FR_U_num,convert(varchar, FR_date_begin, 111),convert(varchar, FR_date_end, 111)";
                T_SQL = T_SQL + " ) R on ad.userID = R.FR_U_num and @yyyy + ad.[attendance_date] between R.FR_date_S";
                T_SQL = T_SQL + " and FR_date_E";
                T_SQL = T_SQL + " where convert( varchar,convert(datetime, @yyyy + [attendance_date]),111";
                T_SQL = T_SQL + "  ) not in ( SELECT convert(varchar, convert(datetime, [HDate]), 111) FROM Holidays )";
                switch (model.AttStatus)
                {
                    case 1:
                        T_SQL = T_SQL + " and (work_time>'09:00' or getoffwork_time<'18:00')";
                        break;
                    case 2:
                        T_SQL = T_SQL + " and (work_time>'09:00' or getoffwork_time<'18:00') and isnull(RestCount, 0) = 0";
                        break;
                    case 3:
                        T_SQL = T_SQL + " and (work_time>'09:00' or getoffwork_time<'18:00') and isnull(RestCount, 0) <> 0";
                        break;
                    case 4:
                        T_SQL = T_SQL + " and work_time>'09:00'";
                        break;
                    case 5:
                        T_SQL = T_SQL + "and getoffwork_time<'18:00'";
                        break;
                }
                T_SQL = T_SQL + " AND userID = @userID AND yyyymm = @yyyymm order by";
                T_SQL = T_SQL + " u_BC,userID,attendance_date";
                var YYYY = (model.yyyymm).Substring(0, 4) + "/";
                parameters.Add(new SqlParameter("@yyyy", YYYY));
                parameters.Add(new SqlParameter("@userID", User_Num));
                parameters.Add(new SqlParameter("@yyyymm", model.yyyymm));
                #endregion
                DataTable dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
                if (dtResult.Rows.Count > 0)
                {
                    var modellist_d = dtResult.AsEnumerable().Select(row => new Attendance_res
                    {
                        U_Na = row.Field<string>("U_Na"),
                        userID = row.Field<string>("userID"),
                        user_name = row.Field<string>("user_name"),
                        attendance_date = row.Field<string>("attendance_date"),
                        work_time = row.Field<string>("work_time"),
                        Late = row.Field<int>("Late"),
                        work_status = row.Field<string>("work_status"),
                        getoffwork_time = row.Field<string>("getoffwork_time"),
                        early = row.Field<int>("early"),
                        offwork_status = row.Field<string>("offwork_status"),
                        U_BC = row.Field<string>("U_BC"),
                        RestCount = row.Field<int>("RestCount")
                    }).ToList();

                    var modellist_c= modellist_d.Where(s=>s.RestCount != 0).ToList();
                    foreach (var item in modellist_c)
                    {
                        #region SQL
                        var T_SQL_d = "SELECT '假別:' + il1.item_D_name + ';' + CONVERT(VARCHAR, fr.FR_date_begin, 111) + '~' +";
                        T_SQL_d = T_SQL_d + " CONVERT(VARCHAR, fr.FR_date_end, 111) + '狀態:' + ";
                        T_SQL_d = T_SQL_d + " CASE WHEN fr.FR_step_now = 1 THEN '代理人-' + COALESCE(um1.U_name, '')";
                        T_SQL_d = T_SQL_d + " WHEN fr.FR_step_now = 2 THEN '直屬主管-' + COALESCE(um2.U_name, '')";
                        T_SQL_d = T_SQL_d + " WHEN fr.FR_step_now = 3 THEN '單位主管-' + COALESCE(um3.U_name, '')";
                        T_SQL_d = T_SQL_d + " WHEN fr.FR_step_now = 9 THEN '人資-'";
                        T_SQL_d = T_SQL_d + " WHEN fr.FR_step_now = 0 THEN '' END + COALESCE(il2.item_D_name, '') AS FR_sign_type_name_desc,";
                        T_SQL_d = T_SQL_d + " fr.FR_total_hour";
                        T_SQL_d = T_SQL_d + " FROM Flow_rest fr";
                        T_SQL_d = T_SQL_d + " LEFT JOIN User_M um1 ON fr.FR_step_01_num = um1.u_num";
                        T_SQL_d = T_SQL_d + " LEFT JOIN User_M um2 ON fr.FR_step_02_num = um2.u_num";
                        T_SQL_d = T_SQL_d + " LEFT JOIN User_M um3 ON fr.FR_step_03_num = um3.u_num";
                        T_SQL_d = T_SQL_d + " LEFT JOIN Item_list il1 ON fr.FR_kind = il1.item_D_code";
                        T_SQL_d = T_SQL_d + " AND il1.item_M_code = 'FR_kind' AND il1.item_D_type = 'Y' AND il1.del_tag = '0'";
                        T_SQL_d = T_SQL_d + " LEFT JOIN Item_list il2 ON fr.FR_sign_type = il2.item_D_code ";
                        T_SQL_d = T_SQL_d + " AND il2.item_M_code = 'Flow_sign_type' AND il2.item_D_type = 'Y' AND il2.del_tag = '0'";
                        T_SQL_d = T_SQL_d + " WHERE fr.del_tag = '0' AND fr.FR_cancel <> 'Y' AND fr.FR_U_num = @FR_U_num AND CONVERT(VARCHAR, fr.FR_date_begin, 111) = @Date";

                        var parameters_d = new List<SqlParameter>();
                        parameters_d.Add(new SqlParameter("@FR_U_num", User_Num));
                        parameters_d.Add(new SqlParameter("@Date", item.attendance_date));
                        #endregion
                        DataTable dtResult_c = _adoData.ExecuteQuery(T_SQL_d, parameters_d);
                        if(dtResult_c.Rows.Count > 0)
                        {
                            DataRow row_c = dtResult_c.Rows[0];
                            item.FR_sign_type_name_desc = row_c["FR_sign_type_name_desc"].ToString();
                            item.FR_total_hour= (decimal)row_c["FR_total_hour"];
                        }
                    }
                    resultClass.ResultCode = "000";
                    resultClass.objResult = JsonConvert.SerializeObject(modellist_d);
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
                return StatusCode(500, resultClass); // 返回 500 錯誤碼
            }
        }

        /// <summary>
        /// 個人出勤紀錄Excel下載 Attendance_SExcel/Attendance_report.asp? Self = Y
        /// </ summary >
        [HttpPost("Attendance_SExcel")]
        public IActionResult Attendance_SExcel(Attendance_req model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            var User_Num = HttpContext.Session.GetString("UserID");

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = "SELECT U.U_Na,[userID],U_name [user_name],@yyyy + ad.[attendance_date] attendance_date,[work_time],";
                T_SQL = T_SQL + " case When isnull([work_time], '') = '' then 0 When [work_time] > '09:00' then DATEDIFF(MINUTE, '09:00', [work_time]) else 0 end Late,";
                T_SQL = T_SQL + " case When isnull([work_time], '') = '' then '未刷卡' When [work_time] > '09:00' then case when isnull(U_num_NL, 'N') = 'N' then '遲到'  else '' end else '' end work_status,";
                T_SQL = T_SQL + " [getoffwork_time],";
                T_SQL = T_SQL + " case When isnull([getoffwork_time], '') = '' then 0 When [getoffwork_time] < '18:00' then DATEDIFF(MINUTE, [getoffwork_time], '18:00') else 0 end early,";
                T_SQL = T_SQL + " case When isnull([getoffwork_time], '') = '' then '未刷卡' When [getoffwork_time] < '18:00' then case when isnull(U_num_NL, 'N') = 'N' then '早退' else '' end else '' end offwork_status,";
                T_SQL = T_SQL + " U_BC,isnull(RestCount, 0) RestCount";
                T_SQL = T_SQL + " FROM attendance ad";
                T_SQL = T_SQL + " left join ( SELECT U_PFT,U_BC,U_num,U_name,I.item_D_name U_Na from User_M U";
                T_SQL = T_SQL + " left join ( SELECT [item_D_code],[item_D_name] FROM Item_list WHERE item_M_code = 'branch_company'";
                T_SQL = T_SQL + " AND item_D_type = 'Y' and del_tag = '0'  ) I on U.u_bc = I.item_D_code";
                T_SQL = T_SQL + " where del_tag = '0' ) U on ad.userID = U.U_num";
                T_SQL = T_SQL + " Left Join ( SELECT [item_D_code] U_num_NL FROM Item_list where item_M_code = 'NonLate'";
                T_SQL = T_SQL + " and [item_M_type] = 'N' ) NL on U.U_num = NL.U_num_NL";
                T_SQL = T_SQL + " left join ( select FR_U_num,convert(varchar, FR_date_begin, 111) FR_date_S,";
                T_SQL = T_SQL + " convert(varchar, FR_date_end, 111) FR_date_E,count(FR_U_num) RestCount";
                T_SQL = T_SQL + " from Flow_rest where  del_tag = '0' and FR_cancel <> 'Y' group by";
                T_SQL = T_SQL + " FR_U_num,convert(varchar, FR_date_begin, 111),convert(varchar, FR_date_end, 111)";
                T_SQL = T_SQL + " ) R on ad.userID = R.FR_U_num and @yyyy + ad.[attendance_date] between R.FR_date_S";
                T_SQL = T_SQL + " and FR_date_E";
                T_SQL = T_SQL + " where convert( varchar,convert(datetime, @yyyy + [attendance_date]),111";
                T_SQL = T_SQL + "  ) not in ( SELECT convert(varchar, convert(datetime, [HDate]), 111) FROM Holidays )";
                switch (model.AttStatus)
                {
                    case 1:
                        T_SQL = T_SQL + " and (work_time>'09:00' or getoffwork_time<'18:00')";
                        break;
                    case 2:
                        T_SQL = T_SQL + " and (work_time>'09:00' or getoffwork_time<'18:00') and isnull(RestCount, 0) = 0";
                        break;
                    case 3:
                        T_SQL = T_SQL + " and (work_time>'09:00' or getoffwork_time<'18:00') and isnull(RestCount, 0) <> 0";
                        break;
                    case 4:
                        T_SQL = T_SQL + " and work_time>'09:00'";
                        break;
                    case 5:
                        T_SQL = T_SQL + "and getoffwork_time<'18:00'";
                        break;
                }
                T_SQL = T_SQL + " AND userID = @userID AND yyyymm = @yyyymm order by";
                T_SQL = T_SQL + " u_BC,userID,attendance_date";

                var parameters = new List<SqlParameter>();
                var YYYY = (model.yyyymm).Substring(0, 4) + "/";
                parameters.Add(new SqlParameter("@yyyy", YYYY));
                parameters.Add(new SqlParameter("@userID", User_Num));
                parameters.Add(new SqlParameter("@yyyymm", model.yyyymm));
                #endregion

                DataTable dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
                if (dtResult.Rows.Count > 0)
                {
                    var modellist_m = dtResult.AsEnumerable().Select(row => new Attendance_res
                    {
                        U_Na = row.Field<string>("U_Na"),
                        userID = row.Field<string>("userID"),
                        user_name = row.Field<string>("user_name"),
                        attendance_date = row.Field<string>("attendance_date"),
                        work_time = row.Field<string>("work_time"),
                        Late = row.Field<int>("Late"),
                        work_status = row.Field<string>("work_status"),
                        getoffwork_time = row.Field<string>("getoffwork_time"),
                        early = row.Field<int>("early"),
                        offwork_status = row.Field<string>("offwork_status"),
                        U_BC = row.Field<string>("U_BC"),
                        RestCount = row.Field<int>("RestCount")
                    }).ToList();

                    var modellist_d = modellist_m.Where(p => p.RestCount != 0).ToList();
                    foreach (var item in modellist_d)
                    {
                        #region SQL
                        var T_SQL_d = "SELECT '假別:' + il1.item_D_name + ';' + CONVERT(VARCHAR, fr.FR_date_begin, 111) + '~' +";
                        T_SQL_d = T_SQL_d + " CONVERT(VARCHAR, fr.FR_date_end, 111) + '狀態:' + ";
                        T_SQL_d = T_SQL_d + " CASE WHEN fr.FR_step_now = 1 THEN '代理人-' + COALESCE(um1.U_name, '')";
                        T_SQL_d = T_SQL_d + " WHEN fr.FR_step_now = 2 THEN '直屬主管-' + COALESCE(um2.U_name, '')";
                        T_SQL_d = T_SQL_d + " WHEN fr.FR_step_now = 3 THEN '單位主管-' + COALESCE(um3.U_name, '')";
                        T_SQL_d = T_SQL_d + " WHEN fr.FR_step_now = 9 THEN '人資-'";
                        T_SQL_d = T_SQL_d + " WHEN fr.FR_step_now = 0 THEN '' END + COALESCE(il2.item_D_name, '') AS FR_sign_type_name_desc,";
                        T_SQL_d = T_SQL_d + " fr.FR_total_hour";
                        T_SQL_d = T_SQL_d + " FROM Flow_rest fr";
                        T_SQL_d = T_SQL_d + " LEFT JOIN User_M um1 ON fr.FR_step_01_num = um1.u_num";
                        T_SQL_d = T_SQL_d + " LEFT JOIN User_M um2 ON fr.FR_step_02_num = um2.u_num";
                        T_SQL_d = T_SQL_d + " LEFT JOIN User_M um3 ON fr.FR_step_03_num = um3.u_num";
                        T_SQL_d = T_SQL_d + " LEFT JOIN Item_list il1 ON fr.FR_kind = il1.item_D_code";
                        T_SQL_d = T_SQL_d + " AND il1.item_M_code = 'FR_kind' AND il1.item_D_type = 'Y' AND il1.del_tag = '0'";
                        T_SQL_d = T_SQL_d + " LEFT JOIN Item_list il2 ON fr.FR_sign_type = il2.item_D_code ";
                        T_SQL_d = T_SQL_d + " AND il2.item_M_code = 'Flow_sign_type' AND il2.item_D_type = 'Y' AND il2.del_tag = '0'";
                        T_SQL_d = T_SQL_d + " WHERE fr.del_tag = '0' AND fr.FR_cancel <> 'Y' AND fr.FR_U_num = @FR_U_num AND CONVERT(VARCHAR, fr.FR_date_begin, 111) = @Date";

                        var parameters_d = new List<SqlParameter>();
                        parameters_d.Add(new SqlParameter("@FR_U_num", User_Num));
                        parameters_d.Add(new SqlParameter("@Date", item.attendance_date));
                        #endregion

                        DataTable dtResult_c = _adoData.ExecuteQuery(T_SQL_d, parameters_d);
                        if (dtResult_c.Rows.Count > 0)
                        {
                            DataRow row_c = dtResult_c.Rows[0];
                            item.FR_sign_type_name_desc = row_c["FR_sign_type_name_desc"].ToString();
                            item.FR_total_hour = (decimal)row_c["FR_total_hour"];
                        }
                    }

                    var excellist = modellist_m.Select(p => new Attendance_res_excel
                    {
                        U_Na = p.U_Na,
                        userID = p.userID,
                        user_name = p.user_name,
                        attendance_date = p.attendance_date,
                        work_time = p.work_time,
                        Late = p.Late,
                        work_status = p.work_status,
                        getoffwork_time = p.getoffwork_time,
                        early = p.early,
                        offwork_status = p.offwork_status,
                        FR_sign_type_name_desc = p.FR_sign_type_name_desc,
                        FR_total_hour = p.FR_total_hour
                    }).ToList();

                    var Attendance_SExcel_Headers = new Dictionary<string, string>
                    {
                        { "U_Na", "公司別" },
                        { "userID", "員編" },
                        { "user_name", "姓名" },
                        { "attendance_date", "日期" },
                        { "work_time", "上班刷卡" },
                        { "work_status", "狀態" },
                        { "Late", "遲到" },
                        { "getoffwork_time", "下班刷卡" },
                        { "offwork_status", "狀態" },
                        { "early", "早退" },
                        { "FR_sign_type_name_desc", "請假資訊" },
                        { "FR_total_hour", "請假時數" }
                    };

                    var fileBytes = FuncHandler.ExportToExcel(excellist, Attendance_SExcel_Headers);
                    var fileName = "請假單報表"+ DateTime.Now.ToString("yyyyMMddHHmm")+".xlsx";
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

        #region 出勤紀錄查詢
        /// <summary>
        /// 出勤紀錄查詢 Attendance_Query/Attendance_report.asp
        /// </summary>
        [HttpPost("Attendance_Query")]
        public ActionResult<ResultClass<string>> Attendance_Query(Attendance_req model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            var User_Num = HttpContext.Session.GetString("UserID");

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = "SELECT U.U_Na,[userID],U_name [user_name],@yyyy + ad.[attendance_date] attendance_date,[work_time],";
                T_SQL = T_SQL + " case When isnull([work_time], '') = '' then 0 When [work_time] > '09:00' then DATEDIFF(MINUTE, '09:00', [work_time]) else 0 end Late,";
                T_SQL = T_SQL + " case When isnull([work_time], '') = '' then '未刷卡' When [work_time] > '09:00' then case when isnull(U_num_NL, 'N') = 'N' then '遲到'  else '' end else '' end work_status,";
                T_SQL = T_SQL + " [getoffwork_time],";
                T_SQL = T_SQL + " case When isnull([getoffwork_time], '') = '' then 0 When [getoffwork_time] < '18:00' then DATEDIFF(MINUTE, [getoffwork_time], '18:00') else 0 end early,";
                T_SQL = T_SQL + " case When isnull([getoffwork_time], '') = '' then '未刷卡' When [getoffwork_time] < '18:00' then case when isnull(U_num_NL, 'N') = 'N' then '早退' else '' end else '' end offwork_status,";
                T_SQL = T_SQL + " U_BC,isnull(RestCount, 0) RestCount";
                T_SQL = T_SQL + " FROM attendance ad";
                T_SQL = T_SQL + " left join ( SELECT U_PFT,U_BC,U_num,U_name,I.item_D_name U_Na from User_M U";
                T_SQL = T_SQL + " left join ( SELECT [item_D_code],[item_D_name] FROM Item_list WHERE item_M_code = 'branch_company'";
                T_SQL = T_SQL + " AND item_D_type = 'Y' and del_tag = '0'  ) I on U.u_bc = I.item_D_code";
                T_SQL = T_SQL + " where del_tag = '0' ) U on ad.userID = U.U_num";
                T_SQL = T_SQL + " Left Join ( SELECT [item_D_code] U_num_NL FROM Item_list where item_M_code = 'NonLate'";
                T_SQL = T_SQL + " and [item_M_type] = 'N' ) NL on U.U_num = NL.U_num_NL";
                T_SQL = T_SQL + " left join ( select FR_U_num,convert(varchar, FR_date_begin, 111) FR_date_S,";
                T_SQL = T_SQL + " convert(varchar, FR_date_end, 111) FR_date_E,count(FR_U_num) RestCount";
                T_SQL = T_SQL + " from Flow_rest where  del_tag = '0' and FR_cancel <> 'Y' group by";
                T_SQL = T_SQL + " FR_U_num,convert(varchar, FR_date_begin, 111),convert(varchar, FR_date_end, 111)";
                T_SQL = T_SQL + " ) R on ad.userID = R.FR_U_num and @yyyy + ad.[attendance_date] between R.FR_date_S";
                T_SQL = T_SQL + " and FR_date_E";
                T_SQL = T_SQL + " where convert( varchar,convert(datetime, @yyyy + [attendance_date]),111";
                T_SQL = T_SQL + "  ) not in ( SELECT convert(varchar, convert(datetime, [HDate]), 111) FROM Holidays )";
                switch (model.AttStatus)
                {
                    case 1:
                        T_SQL = T_SQL + " and (work_time>'09:00' or getoffwork_time<'18:00')";
                        break;
                    case 2:
                        T_SQL = T_SQL + " and (work_time>'09:00' or getoffwork_time<'18:00') and isnull(RestCount, 0) = 0";
                        break;
                    case 3:
                        T_SQL = T_SQL + " and (work_time>'09:00' or getoffwork_time<'18:00') and isnull(RestCount, 0) <> 0";
                        break;
                    case 4:
                        T_SQL = T_SQL + " and work_time>'09:00'";
                        break;
                    case 5:
                        T_SQL = T_SQL + "and getoffwork_time<'18:00'";
                        break;
                }
                if (!string.IsNullOrEmpty(model.U_num))
                {
                    T_SQL = T_SQL + " AND userID = @userID";
                    parameters.Add(new SqlParameter("@userID", model.U_num));
                }
                if (!string.IsNullOrEmpty(model.U_BU))
                {
                    T_SQL = T_SQL + " AND u_bc=@U_BU";
                    parameters.Add(new SqlParameter("@U_BU", model.U_BU));
                }
                if (!string.IsNullOrEmpty(model.U_name))
                {
                    T_SQL = T_SQL + " AND user_name=@U_name";
                    parameters.Add(new SqlParameter("@U_name", model.U_name));
                }
                T_SQL = T_SQL + " AND yyyymm = @yyyymm order by";
                T_SQL = T_SQL + " u_BC,userID,attendance_date";
                var YYYY = (model.yyyymm).Substring(0, 4) + "/";
                parameters.Add(new SqlParameter("@yyyy", YYYY));
                parameters.Add(new SqlParameter("@yyyymm", model.yyyymm));
                #endregion
                DataTable dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
                if (dtResult.Rows.Count > 0)
                {
                    var modellist_d = dtResult.AsEnumerable().Select(row => new Attendance_res
                    {
                        U_Na = row.Field<string>("U_Na"),
                        userID = row.Field<string>("userID"),
                        user_name = row.Field<string>("user_name"),
                        attendance_date = row.Field<string>("attendance_date"),
                        work_time = row.Field<string>("work_time"),
                        Late = row.Field<int>("Late"),
                        work_status = row.Field<string>("work_status"),
                        getoffwork_time = row.Field<string>("getoffwork_time"),
                        early = row.Field<int>("early"),
                        offwork_status = row.Field<string>("offwork_status"),
                        U_BC = row.Field<string>("U_BC"),
                        RestCount = row.Field<int>("RestCount")
                    }).ToList();

                    var modellist_c = modellist_d.Where(s => s.RestCount != 0).ToList();
                    foreach (var item in modellist_c)
                    {
                        #region SQL
                        var T_SQL_d = "SELECT '假別:' + il1.item_D_name + ';' + CONVERT(VARCHAR, fr.FR_date_begin, 111) + '~' +";
                        T_SQL_d = T_SQL_d + " CONVERT(VARCHAR, fr.FR_date_end, 111) + '狀態:' + ";
                        T_SQL_d = T_SQL_d + " CASE WHEN fr.FR_step_now = 1 THEN '代理人-' + COALESCE(um1.U_name, '')";
                        T_SQL_d = T_SQL_d + " WHEN fr.FR_step_now = 2 THEN '直屬主管-' + COALESCE(um2.U_name, '')";
                        T_SQL_d = T_SQL_d + " WHEN fr.FR_step_now = 3 THEN '單位主管-' + COALESCE(um3.U_name, '')";
                        T_SQL_d = T_SQL_d + " WHEN fr.FR_step_now = 9 THEN '人資-'";
                        T_SQL_d = T_SQL_d + " WHEN fr.FR_step_now = 0 THEN '' END + COALESCE(il2.item_D_name, '') AS FR_sign_type_name_desc,";
                        T_SQL_d = T_SQL_d + " fr.FR_total_hour";
                        T_SQL_d = T_SQL_d + " FROM Flow_rest fr";
                        T_SQL_d = T_SQL_d + " LEFT JOIN User_M um1 ON fr.FR_step_01_num = um1.u_num";
                        T_SQL_d = T_SQL_d + " LEFT JOIN User_M um2 ON fr.FR_step_02_num = um2.u_num";
                        T_SQL_d = T_SQL_d + " LEFT JOIN User_M um3 ON fr.FR_step_03_num = um3.u_num";
                        T_SQL_d = T_SQL_d + " LEFT JOIN Item_list il1 ON fr.FR_kind = il1.item_D_code";
                        T_SQL_d = T_SQL_d + " AND il1.item_M_code = 'FR_kind' AND il1.item_D_type = 'Y' AND il1.del_tag = '0'";
                        T_SQL_d = T_SQL_d + " LEFT JOIN Item_list il2 ON fr.FR_sign_type = il2.item_D_code ";
                        T_SQL_d = T_SQL_d + " AND il2.item_M_code = 'Flow_sign_type' AND il2.item_D_type = 'Y' AND il2.del_tag = '0'";
                        T_SQL_d = T_SQL_d + " WHERE fr.del_tag = '0' AND fr.FR_cancel <> 'Y' AND fr.FR_U_num = @FR_U_num AND CONVERT(VARCHAR, fr.FR_date_begin, 111) = @Date";

                        var parameters_d = new List<SqlParameter>();
                        parameters_d.Add(new SqlParameter("@FR_U_num", item.userID));
                        parameters_d.Add(new SqlParameter("@Date", item.attendance_date));
                        #endregion
                        DataTable dtResult_c = _adoData.ExecuteQuery(T_SQL_d, parameters_d);
                        if (dtResult_c.Rows.Count > 0)
                        {
                            DataRow row_c = dtResult_c.Rows[0];
                            item.FR_sign_type_name_desc = row_c["FR_sign_type_name_desc"].ToString();
                            item.FR_total_hour = (decimal)row_c["FR_total_hour"];
                        }
                    }
                    resultClass.ResultCode = "000";
                    resultClass.objResult = JsonConvert.SerializeObject(modellist_d);
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
                return StatusCode(500, resultClass); // 返回 500 錯誤碼
            }
        }

        /// <summary>
        /// 出勤紀錄查詢Excel下載 Attendance_Excel/Attendance_report.asp
        /// </ summary >
        [HttpPost("Attendance_Excel")]
        public IActionResult Attendance_Excel(Attendance_req model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            var User_Num = HttpContext.Session.GetString("UserID");

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = "SELECT U.U_Na,[userID],U_name [user_name],@yyyy + ad.[attendance_date] attendance_date,[work_time],";
                T_SQL = T_SQL + " case When isnull([work_time], '') = '' then 0 When [work_time] > '09:00' then DATEDIFF(MINUTE, '09:00', [work_time]) else 0 end Late,";
                T_SQL = T_SQL + " case When isnull([work_time], '') = '' then '未刷卡' When [work_time] > '09:00' then case when isnull(U_num_NL, 'N') = 'N' then '遲到'  else '' end else '' end work_status,";
                T_SQL = T_SQL + " [getoffwork_time],";
                T_SQL = T_SQL + " case When isnull([getoffwork_time], '') = '' then 0 When [getoffwork_time] < '18:00' then DATEDIFF(MINUTE, [getoffwork_time], '18:00') else 0 end early,";
                T_SQL = T_SQL + " case When isnull([getoffwork_time], '') = '' then '未刷卡' When [getoffwork_time] < '18:00' then case when isnull(U_num_NL, 'N') = 'N' then '早退' else '' end else '' end offwork_status,";
                T_SQL = T_SQL + " U_BC,isnull(RestCount, 0) RestCount";
                T_SQL = T_SQL + " FROM attendance ad";
                T_SQL = T_SQL + " left join ( SELECT U_PFT,U_BC,U_num,U_name,I.item_D_name U_Na from User_M U";
                T_SQL = T_SQL + " left join ( SELECT [item_D_code],[item_D_name] FROM Item_list WHERE item_M_code = 'branch_company'";
                T_SQL = T_SQL + " AND item_D_type = 'Y' and del_tag = '0'  ) I on U.u_bc = I.item_D_code";
                T_SQL = T_SQL + " where del_tag = '0' ) U on ad.userID = U.U_num";
                T_SQL = T_SQL + " Left Join ( SELECT [item_D_code] U_num_NL FROM Item_list where item_M_code = 'NonLate'";
                T_SQL = T_SQL + " and [item_M_type] = 'N' ) NL on U.U_num = NL.U_num_NL";
                T_SQL = T_SQL + " left join ( select FR_U_num,convert(varchar, FR_date_begin, 111) FR_date_S,";
                T_SQL = T_SQL + " convert(varchar, FR_date_end, 111) FR_date_E,count(FR_U_num) RestCount";
                T_SQL = T_SQL + " from Flow_rest where  del_tag = '0' and FR_cancel <> 'Y' group by";
                T_SQL = T_SQL + " FR_U_num,convert(varchar, FR_date_begin, 111),convert(varchar, FR_date_end, 111)";
                T_SQL = T_SQL + " ) R on ad.userID = R.FR_U_num and @yyyy + ad.[attendance_date] between R.FR_date_S";
                T_SQL = T_SQL + " and FR_date_E";
                T_SQL = T_SQL + " where convert( varchar,convert(datetime, @yyyy + [attendance_date]),111";
                T_SQL = T_SQL + "  ) not in ( SELECT convert(varchar, convert(datetime, [HDate]), 111) FROM Holidays )";
                switch (model.AttStatus)
                {
                    case 1:
                        T_SQL = T_SQL + " and (work_time>'09:00' or getoffwork_time<'18:00')";
                        break;
                    case 2:
                        T_SQL = T_SQL + " and (work_time>'09:00' or getoffwork_time<'18:00') and isnull(RestCount, 0) = 0";
                        break;
                    case 3:
                        T_SQL = T_SQL + " and (work_time>'09:00' or getoffwork_time<'18:00') and isnull(RestCount, 0) <> 0";
                        break;
                    case 4:
                        T_SQL = T_SQL + " and work_time>'09:00'";
                        break;
                    case 5:
                        T_SQL = T_SQL + "and getoffwork_time<'18:00'";
                        break;
                }
                if (!string.IsNullOrEmpty(model.U_num))
                {
                    T_SQL = T_SQL + " AND userID = @userID";
                    parameters.Add(new SqlParameter("@userID", model.U_num));
                }
                if (!string.IsNullOrEmpty(model.U_BU))
                {
                    T_SQL = T_SQL + " AND u_bc=@U_BU";
                    parameters.Add(new SqlParameter("@U_BU", model.U_BU));
                }
                if (!string.IsNullOrEmpty(model.U_name))
                {
                    T_SQL = T_SQL + " AND user_name=@U_name";
                    parameters.Add(new SqlParameter("@U_name", model.U_name));
                }
                T_SQL = T_SQL + " AND yyyymm = @yyyymm order by";
                T_SQL = T_SQL + " u_BC,userID,attendance_date";
                var YYYY = (model.yyyymm).Substring(0, 4) + "/";
                parameters.Add(new SqlParameter("@yyyy", YYYY));
                parameters.Add(new SqlParameter("@yyyymm", model.yyyymm));
                #endregion

                DataTable dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
                if (dtResult.Rows.Count > 0)
                {
                    var modellist_m = dtResult.AsEnumerable().Select(row => new Attendance_res
                    {
                        U_Na = row.Field<string>("U_Na"),
                        userID = row.Field<string>("userID"),
                        user_name = row.Field<string>("user_name"),
                        attendance_date = row.Field<string>("attendance_date"),
                        work_time = row.Field<string>("work_time"),
                        Late = row.Field<int>("Late"),
                        work_status = row.Field<string>("work_status"),
                        getoffwork_time = row.Field<string>("getoffwork_time"),
                        early = row.Field<int>("early"),
                        offwork_status = row.Field<string>("offwork_status"),
                        U_BC = row.Field<string>("U_BC"),
                        RestCount = row.Field<int>("RestCount")
                    }).ToList();

                    var modellist_d = modellist_m.Where(p => p.RestCount != 0).ToList();
                    foreach (var item in modellist_d)
                    {
                        #region SQL
                        var T_SQL_d = "SELECT '假別:' + il1.item_D_name + ';' + CONVERT(VARCHAR, fr.FR_date_begin, 111) + '~' +";
                        T_SQL_d = T_SQL_d + " CONVERT(VARCHAR, fr.FR_date_end, 111) + '狀態:' + ";
                        T_SQL_d = T_SQL_d + " CASE WHEN fr.FR_step_now = 1 THEN '代理人-' + COALESCE(um1.U_name, '')";
                        T_SQL_d = T_SQL_d + " WHEN fr.FR_step_now = 2 THEN '直屬主管-' + COALESCE(um2.U_name, '')";
                        T_SQL_d = T_SQL_d + " WHEN fr.FR_step_now = 3 THEN '單位主管-' + COALESCE(um3.U_name, '')";
                        T_SQL_d = T_SQL_d + " WHEN fr.FR_step_now = 9 THEN '人資-'";
                        T_SQL_d = T_SQL_d + " WHEN fr.FR_step_now = 0 THEN '' END + COALESCE(il2.item_D_name, '') AS FR_sign_type_name_desc,";
                        T_SQL_d = T_SQL_d + " fr.FR_total_hour";
                        T_SQL_d = T_SQL_d + " FROM Flow_rest fr";
                        T_SQL_d = T_SQL_d + " LEFT JOIN User_M um1 ON fr.FR_step_01_num = um1.u_num";
                        T_SQL_d = T_SQL_d + " LEFT JOIN User_M um2 ON fr.FR_step_02_num = um2.u_num";
                        T_SQL_d = T_SQL_d + " LEFT JOIN User_M um3 ON fr.FR_step_03_num = um3.u_num";
                        T_SQL_d = T_SQL_d + " LEFT JOIN Item_list il1 ON fr.FR_kind = il1.item_D_code";
                        T_SQL_d = T_SQL_d + " AND il1.item_M_code = 'FR_kind' AND il1.item_D_type = 'Y' AND il1.del_tag = '0'";
                        T_SQL_d = T_SQL_d + " LEFT JOIN Item_list il2 ON fr.FR_sign_type = il2.item_D_code ";
                        T_SQL_d = T_SQL_d + " AND il2.item_M_code = 'Flow_sign_type' AND il2.item_D_type = 'Y' AND il2.del_tag = '0'";
                        T_SQL_d = T_SQL_d + " WHERE fr.del_tag = '0' AND fr.FR_cancel <> 'Y' AND fr.FR_U_num = @FR_U_num AND CONVERT(VARCHAR, fr.FR_date_begin, 111) = @Date";

                        var parameters_d = new List<SqlParameter>();
                        parameters_d.Add(new SqlParameter("@FR_U_num", item.userID));
                        parameters_d.Add(new SqlParameter("@Date", item.attendance_date));
                        #endregion

                        DataTable dtResult_c = _adoData.ExecuteQuery(T_SQL_d, parameters_d);
                        if (dtResult_c.Rows.Count > 0)
                        {
                            DataRow row_c = dtResult_c.Rows[0];
                            item.FR_sign_type_name_desc = row_c["FR_sign_type_name_desc"].ToString();
                            item.FR_total_hour = (decimal)row_c["FR_total_hour"];
                        }
                    }

                    var excellist = modellist_m.Select(p => new Attendance_res_excel
                    {
                        U_Na = p.U_Na,
                        userID = p.userID,
                        user_name = p.user_name,
                        attendance_date = p.attendance_date,
                        work_time = p.work_time,
                        Late = p.Late,
                        work_status = p.work_status,
                        getoffwork_time = p.getoffwork_time,
                        early = p.early,
                        offwork_status = p.offwork_status,
                        FR_sign_type_name_desc = p.FR_sign_type_name_desc,
                        FR_total_hour = p.FR_total_hour
                    }).ToList();

                    var Attendance_SExcel_Headers = new Dictionary<string, string>
                    {
                        { "U_Na", "公司別" },
                        { "userID", "員編" },
                        { "user_name", "姓名" },
                        { "attendance_date", "日期" },
                        { "work_time", "上班刷卡" },
                        { "work_status", "狀態" },
                        { "Late", "遲到" },
                        { "getoffwork_time", "下班刷卡" },
                        { "offwork_status", "狀態" },
                        { "early", "早退" },
                        { "FR_sign_type_name_desc", "請假資訊" },
                        { "FR_total_hour", "請假時數" }
                    };

                    var fileBytes = FuncHandler.ExportToExcel(excellist, Attendance_SExcel_Headers);
                    var fileName = "請假單報表" + DateTime.Now.ToString("yyyyMMddHHmm") + ".xlsx";
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

        #region 使用者管理
        /// <summary>
        /// 使用者清單查詢 UserM_List/User_list.asp
        /// </summary>
        [HttpPost("User_M_LQuery")]
        public ActionResult<ResultClass<string>> User_M_LQuery(Uesr_M_req model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = "select (select item_D_name from Item_list where item_M_code = 'branch_company' AND item_D_type='Y' AND item_D_code = UM.U_BC AND del_tag='0') as U_BC_name";
                T_SQL = T_SQL + " ,(select item_D_name from Item_list where item_M_code = 'professional_title' AND item_D_type='Y' AND item_D_code = UM.U_PFT AND del_tag='0') as U_PFT_name";
                T_SQL = T_SQL + " ,(select U_name FROM User_M where U_num = UM.U_agent_num AND del_tag='0') as U_agent_name ";
                T_SQL = T_SQL + " ,U_num,UM.del_tag,(select U_name FROM User_M where U_num = UM.U_leader_1_num AND del_tag='0') as U_leader_1_name ";
                T_SQL = T_SQL + " ,(select U_name FROM User_M where U_num = UM.U_leader_2_num AND del_tag='0') as U_leader_2_name  ";
                T_SQL = T_SQL + " ,(select U_name FROM User_M where U_num = UM.U_leader_3_num AND del_tag='0') as U_leader_3_name,Rm.R_name,UM.U_Check_BC";
                T_SQL = T_SQL + " ,(select item_sort from Item_list where item_M_code = 'branch_company' AND item_D_type='Y' AND item_D_code = UM.U_BC AND del_tag='0') as U_BC_sort ";
                T_SQL = T_SQL + " ,(select item_sort from Item_list where item_M_code = 'professional_title' AND item_D_type='Y' AND item_D_code = UM.U_PFT AND del_tag='0') as U_PFT_sort ";
                T_SQL = T_SQL + " from User_M UM left join Role_M Rm on UM.Role_num = Rm.R_num where 1=1";
                if (!string.IsNullOrEmpty(model.U_Num_Name))
                {
                    T_SQL = T_SQL + " AND (U_num LIKE '%' + @U_Num_Name + '%' OR U_name LIKE '%' + @U_Num_Name + '%') ";
                    parameters.Add(new SqlParameter("@U_Num_Name", model.U_Num_Name));
                }
                if (!string.IsNullOrEmpty(model.U_BC))
                {
                    T_SQL = T_SQL + " AND U_BC='@U_BC";
                    parameters.Add(new SqlParameter("@U_BC", model.U_BC));
                }
                if (!string.IsNullOrEmpty(model.Job_Status)) 
                {
                    if(model.Job_Status == "Y")
                    {
                        T_SQL = T_SQL + " AND ISNULL(UM.U_leave_date, '') = ''";
                    }
                    else
                    {
                        T_SQL = T_SQL + " AND ISNULL(UM.U_leave_date, '') <> ''";
                    }
                }
                if (!string.IsNullOrEmpty(model.U_Role)) 
                {
                    T_SQL = T_SQL + " AND Rm.R_num=@R_num";
                }
                T_SQL = T_SQL + " order by U_type desc,U_leave_date,U_BC_sort,U_PFT_sort,U_id";
                #endregion

                DataTable dtResult = _adoData.ExecuteQuery(T_SQL,parameters);
                if (dtResult.Rows.Count > 0)
                {
                    var modellist_M = dtResult.AsEnumerable().Select(row => new User_M_res
                    {
                        U_BC_name = row.Field<string>("U_BC_name"),
                        U_PFT_name = row.Field<string>("U_PFT_name"),
                        U_agent_name = row.Field<string>("U_agent_name"),
                        U_num = row.Field<string>("U_num"),
                        del_tag = row.Field<string>("del_tag"),
                        U_leader_1_name = row.Field<string>("U_leader_1_name"),
                        U_leader_2_name = row.Field<string>("U_leader_2_name"),
                        U_leader_3_name = row.Field<string>("U_leader_3_name"),
                        R_name = row.Field<string>("R_name"),
                        U_Check_BC = row.Field<string>("U_Check_BC"),
                        U_Check_BC_Name = null,
                        U_BC_sort = row.Field<int>("U_BC_sort"),
                        U_PFT_sort = row.Field<int>("U_PFT_sort")
                    }).ToList();

                    foreach (var item in modellist_M)
                    {
                        var str_Bc = item.U_Check_BC.Split('#').Where(code => !string.IsNullOrEmpty(code)).ToArray();
                        if (str_Bc.Length == 0)
                            continue;
                        #region SQL
                        var parameters_bc = new List<SqlParameter>();
                        var T_SQL_BC = "SELECT item_D_name FROM Item_list WHERE item_M_code = 'branch_company' AND item_D_type = 'Y'";
                        var inClause = string.Join(", ", str_Bc.Select((code, index) => "@code" + index));
                        T_SQL_BC = T_SQL_BC + " AND item_D_code IN (" + inClause + ")";
                        for (int i = 0; i < str_Bc.Length; i++)
                        {
                            parameters_bc.Add(new SqlParameter("@code" + i, str_Bc[i]));
                        }
                        #endregion
                        DataTable result_bc =_adoData.ExecuteQuery(T_SQL_BC, parameters_bc);
                        var values = result_bc.AsEnumerable()
                             .Select(row => "#" + row.Field<string>("item_D_name"))
                             .ToArray();
                        item.U_Check_BC_Name = string.Join("", values);
                    }

                    resultClass.ResultCode = "000";
                    var pageData = FuncHandler.GetPagedList(modellist_M, model.page, 100);
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
                return StatusCode(500, resultClass); 
            }
        }
        /// <summary>
        /// 個人資料單筆查詢 User_M_SQuery/User_edit.asp
        /// </summary>
        [HttpPost("User_M_SQuery")]
        public ActionResult<ResultClass<string>> User_M_SQuery(string U_num)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = "select UM.* ";
                T_SQL = T_SQL + " ,(select U_name FROM User_M where U_num = UM.U_agent_num AND del_tag='0') as U_agent_name";
                T_SQL = T_SQL + " ,(select U_name FROM User_M where U_num = UM.U_leader_1_num AND del_tag='0') as U_leader_1_name";
                T_SQL = T_SQL + " ,(select U_name FROM User_M where U_num = UM.U_leader_2_num AND del_tag='0') as U_leader_2_name";
                T_SQL = T_SQL + " ,(select U_name FROM User_M where U_num = UM.U_leader_3_num AND del_tag='0') as U_leader_3_name";
                T_SQL = T_SQL + " ,Rm.R_num ,Rm.R_name";
                T_SQL = T_SQL + " from User_M UM left join Role_M Rm on Rm.R_num = UM.Role_num where 1=1 AND U_num=@U_num";
                parameters.Add(new SqlParameter("@U_num", U_num));
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
        /// 個人資料修改 User_M_SUpd/User_edit.asp
        /// </summary>
        [HttpPost("User_M_SUpd")]
        public ActionResult<ResultClass<string>> User_M_SUpd(User_M model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            var User_Num = HttpContext.Session.GetString("UserID");
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = "Update User_M set edit_num=@edit_num,edit_date=@edit_date,edit_ip=@edit_ip,U_sex=@U_sex,Marriage=@Marriage,Military=@Military";
                T_SQL = T_SQL + " ,Military_SDate=@Military_SDate,Military_EDate=@Military_EDate,Military_Exemption=@Military_Exemption,License_Car=@License_Car";
                T_SQL = T_SQL + " ,Self_Car=@Self_Car,License_Motorcycle=@License_Motorcycle,Self_Motorcycle=@Self_Motorcycle,School_SDate=@School_SDate,School_EDate=@School_EDate";
                T_SQL = T_SQL + " ,U_BC=@U_BC,U_PFT=@U_PFT,Role_num=@Role_num,U_agent_num=@U_agent_num";
                if (!string.IsNullOrEmpty(model.U_name))
                {
                    T_SQL = T_SQL + ",U_name=@U_name";
                    parameters.Add(new SqlParameter("@U_name", model.U_name));
                }
                if (!string.IsNullOrEmpty(model.U_Ename))
                {
                    T_SQL = T_SQL + ",U_Ename=@U_Ename";
                    parameters.Add(new SqlParameter("@U_Ename", model.U_Ename));
                }
                if (model.U_Birthday != null)
                {
                    T_SQL = T_SQL + ",U_Birthday=@U_Birthday";
                    parameters.Add(new SqlParameter("@U_Birthday", model.U_Birthday));
                }
                if (model.Children != null) 
                {
                    T_SQL = T_SQL + " ,Children=@Children";
                    parameters.Add(new SqlParameter("@Children", model.Children));
                }
                else
                {
                    T_SQL = T_SQL + " ,Children=@Children";
                    parameters.Add(new SqlParameter("@Children", DBNull.Value));
                }
                if (!string.IsNullOrEmpty(model.U_PID)) 
                {
                    T_SQL = T_SQL + ",U_PID=@U_PID";
                    parameters.Add(new SqlParameter("@U_PID",model.U_PID));
                }
                if (!string.IsNullOrEmpty(model.U_Tel)) 
                {
                    T_SQL = T_SQL + " ,U_Tel=@U_Tel";
                    parameters.Add(new SqlParameter("@U_Tel", model.U_Tel));
                }
                if (!string.IsNullOrEmpty(model.U_MTel)) 
                {
                    T_SQL = T_SQL + " ,U_MTel=@U_MTel";
                    parameters.Add(new SqlParameter("@U_MTel", model.U_MTel));
                }
                if (!string.IsNullOrEmpty(model.U_Email)) 
                {
                    T_SQL = T_SQL + " ,U_Email=@U_Email";
                    parameters.Add(new SqlParameter("@U_Email", model.U_Email));
                }
                if (!string.IsNullOrEmpty(model.Emergency_contact)) 
                {
                    T_SQL = T_SQL + " ,Emergency_contact=@Emergency_contact";
                    parameters.Add(new SqlParameter("@Emergency_contact", model.Emergency_contact));
                }
                if (!string.IsNullOrEmpty(model.Emergency_Tel)) 
                {
                    T_SQL = T_SQL + " ,Emergency_Tel=@Emergency_Tel";
                    parameters.Add(new SqlParameter("@Emergency_Tel", model.Emergency_Tel));
                }
                if (!string.IsNullOrEmpty(model.Emergency_MTel)) 
                {
                    T_SQL = T_SQL + " ,Emergency_MTel=@Emergency_MTel";
                    parameters.Add(new SqlParameter("@Emergency_MTel", model.Emergency_MTel));
                }
                if (!string.IsNullOrEmpty(model.School_Level)) 
                {
                    T_SQL = T_SQL + " ,School_Level=@School_Level";
                    parameters.Add(new SqlParameter("@School_Level", model.School_Level));
                }
                if (!string.IsNullOrEmpty(model.School_Name)) 
                {
                    T_SQL = T_SQL + " ,School_Name=@School_Name";
                    parameters.Add(new SqlParameter("@School_Name", model.School_Name));
                }
                if (!string.IsNullOrEmpty(model.School_Graduated)) 
                {
                    T_SQL = T_SQL + " ,School_Graduated=@School_Graduated";
                    parameters.Add(new SqlParameter("@School_Graduated", model.School_Graduated));
                }
                if (!string.IsNullOrEmpty(model.School_D_N)) 
                {
                    T_SQL = T_SQL + " ,School_D_N=@School_D_N";
                    parameters.Add(new SqlParameter("@School_D_N", model.School_D_N));
                }
                if (!string.IsNullOrEmpty(model.School_Major)) 
                {
                    T_SQL = T_SQL + " ,School_Major=@School_Major";
                    parameters.Add(new SqlParameter("@School_Major", model.School_Major));
                }
                if (!string.IsNullOrEmpty(model.U_leader_1_num)) 
                {
                    T_SQL = T_SQL + " ,U_leader_1_num=@U_leader_1_num";
                    parameters.Add(new SqlParameter("@U_leader_1_num", model.U_leader_1_num));
                }
                if (!string.IsNullOrEmpty(model.U_leader_2_num)) 
                {
                    T_SQL = T_SQL + " ,U_leader_2_num=@U_leader_2_num";
                    parameters.Add(new SqlParameter("@U_leader_2_num", model.U_leader_2_num));
                }
                if (!string.IsNullOrEmpty(model.U_Check_BC)) 
                {
                    T_SQL = T_SQL + " ,U_Check_BC=@U_Check_BC";
                    parameters.Add(new SqlParameter("@U_Check_BC", model.U_Check_BC));
                }
                if (!string.IsNullOrEmpty(model.U_address_live)) 
                {
                    T_SQL = T_SQL + " ,U_address_live=@U_address_live";
                    parameters.Add(new SqlParameter("@U_address_live", model.U_address_live));
                }
                if (model.U_arrive_date != null) 
                {
                    T_SQL = T_SQL + " ,U_arrive_date=@U_arrive_date";
                    parameters.Add(new SqlParameter("@U_arrive_date", model.U_arrive_date));
                }
                if (model.U_leave_date != null)
                {
                    T_SQL = T_SQL + " ,U_leave_date=@U_leave_date";
                    parameters.Add(new SqlParameter("@U_leave_date", model.U_leave_date));
                }
                else
                {
                    T_SQL = T_SQL + " ,U_leave_date=@U_leave_date";
                    parameters.Add(new SqlParameter("@U_leave_date", DBNull.Value));
                }
                T_SQL = T_SQL + " Where U_id=@U_id";
                parameters.Add(new SqlParameter("@edit_num", User_Num));
                parameters.Add(new SqlParameter("@edit_date", DateTime.Now));
                parameters.Add(new SqlParameter("@edit_ip", clientIp));
                parameters.Add(new SqlParameter("@U_sex", model.U_sex));
                parameters.Add(new SqlParameter("@Marriage", model.Marriage));
                parameters.Add(new SqlParameter("@Military", model.Military));
                parameters.Add(new SqlParameter("@Military_SDate", model.Military_SDate));
                parameters.Add(new SqlParameter("@Military_EDate", model.Military_EDate));
                parameters.Add(new SqlParameter("@Military_Exemption", model.Military_Exemption));
                parameters.Add(new SqlParameter("@License_Car", model.License_Car));
                parameters.Add(new SqlParameter("@Self_Car", model.Self_Car));
                parameters.Add(new SqlParameter("@License_Motorcycle", model.License_Motorcycle));
                parameters.Add(new SqlParameter("@Self_Motorcycle", model.Self_Motorcycle));
                parameters.Add(new SqlParameter("@School_SDate", model.School_SDate));
                parameters.Add(new SqlParameter("@School_EDate", model.School_EDate));
                parameters.Add(new SqlParameter("@U_BC", model.U_BC));
                parameters.Add(new SqlParameter("@U_PFT", model.U_PFT));
                parameters.Add(new SqlParameter("@Role_num", model.Role_num));
                parameters.Add(new SqlParameter("@U_agent_num", model.U_agent_num));
                parameters.Add(new SqlParameter("@U_id", model.U_id));
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
            catch (Exception)
            {
                resultClass.ResultCode = "500";
                return StatusCode(500, resultClass); 
            }
        }
        /// <summary>
        /// 個人資料新增 User_M_SUpd/User_add.asp
        /// </summary>
        [HttpPost("User_M_Ins")]
        public ActionResult<ResultClass<string>> User_M_Ins(User_M model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var User_Num = HttpContext.Session.GetString("UserID");
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

            try
            {
                ADOData _adoData = new ADOData();
                #region 檢查員編
                #region SQL
                var parameters_nu = new List<SqlParameter>();
                var T_SQL_NU = "select * from User_M where del_tag = '0' AND U_num=@U_num ";
                parameters_nu.Add(new SqlParameter("@U_num", model.U_num));
                #endregion
                int result_nu = _adoData.ExecuteNonQuery(T_SQL_NU, parameters_nu);
                if (result_nu > 0)
                {
                    resultClass.ResultCode = "400";
                    resultClass.ResultMsg = "此員編已被使用";
                    return BadRequest(resultClass);
                }
                #endregion

                #region SQL
                var T_SQL = "Insert into User_M ( add_num,add_date,add_ip,U_cknum,U_num,U_name,U_Ename,U_Birthday,U_sex,Marriage,Children,U_PID,Military,Military_SDate,";
                T_SQL = T_SQL + " Military_EDate,Military_Exemption,License_Car,Self_Car,License_Motorcycle,Self_Motorcycle,";
                T_SQL = T_SQL + " U_Tel,U_MTel,U_Email,Emergency_contact,Emergency_Tel,Emergency_MTel,School_Level,School_Name,";
                T_SQL = T_SQL + " School_SDate,School_EDate,School_Graduated,School_D_N,School_Major,U_BC,U_PFT,Role_num,";
                T_SQL = T_SQL + " U_agent_num,U_leader_1_num,U_leader_2_num,U_leader_3_num,U_Check_BC,U_address_live,U_arrive_date,U_leave_date )";
                T_SQL = T_SQL + " Values ( @add_num,@add_date,@add_ip,@U_cknum,@U_num,@U_name,@U_Ename,@U_Birthday,@U_sex,@Marriage,@Children,";
                T_SQL = T_SQL + " @U_PID,@Military,@Military_SDate,@Military_EDate,@Military_Exemption,@License_Car,@Self_Car,";
                T_SQL = T_SQL + " @License_Motorcycle,@Self_Motorcycle,@U_Tel,@U_MTel,@U_Email,@Emergency_contact,@Emergency_Tel,";
                T_SQL = T_SQL + " @Emergency_MTel,@School_Level,@School_Name,@School_SDate,@School_EDate,@School_Graduated,@School_D_N,";
                T_SQL = T_SQL + " @School_Major,@U_BC,@U_PFT,@Role_num,@U_agent_num,@U_leader_1_num,@U_leader_2_num,@U_leader_3_num,@U_Check_BC,@U_address_live,@U_arrive_date,@U_leave_date )";
                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@add_num", User_Num));
                parameters.Add(new SqlParameter("@add_date",DateTime.Now));
                parameters.Add(new SqlParameter("@add_ip", clientIp));
                var Fun = new FuncHandler();
                parameters.Add(new SqlParameter("@U_cknum", Fun.GetCheckNum()));
                parameters.Add(new SqlParameter("@U_num", model.U_num));
                parameters.Add(new SqlParameter("@U_name", model.U_name));
                if (!string.IsNullOrEmpty(model.U_Ename))
                {
                    parameters.Add(new SqlParameter("@U_Ename", model.U_Ename));
                }
                else
                {
                    parameters.Add (new SqlParameter("@U_Ename","")); 
                }
                if (model.U_Birthday != null)
                {
                    parameters.Add(new SqlParameter("@U_Birthday", model.U_Birthday));
                }
                else
                {
                    parameters.Add(new SqlParameter("@U_Birthday", DBNull.Value));
                }
                parameters.Add(new SqlParameter("@U_sex", model.U_sex));
                parameters.Add(new SqlParameter("@Marriage", model.Marriage));
                if(model.Children != null) 
                {
                    parameters.Add(new SqlParameter("@Children", model.Children));
                }
                else
                {
                    parameters.Add (new SqlParameter("@Children",DBNull.Value));
                }
                if (!string.IsNullOrEmpty(model.U_PID))
                {
                    parameters.Add(new SqlParameter("@U_PID", model.U_PID));
                }
                else
                {
                    parameters.Add(new SqlParameter("@U_PID", DBNull.Value));
                }
                parameters.Add(new SqlParameter("@Military",model.Military));
                parameters.Add(new SqlParameter("@Military_SDate", model.Military_SDate));
                parameters.Add(new SqlParameter("@Military_EDate", model.Military_EDate));
                parameters.Add(new SqlParameter("@Military_Exemption", model.Military_Exemption));
                parameters.Add(new SqlParameter("@License_Car", model.License_Car));
                parameters.Add(new SqlParameter("@Self_Car", model.Self_Car));
                parameters.Add(new SqlParameter("@License_Motorcycle", model.License_Motorcycle));
                parameters.Add(new SqlParameter("@Self_Motorcycle", model.Self_Motorcycle));
                if (!string.IsNullOrEmpty(model.U_Tel))
                {
                    parameters.Add(new SqlParameter("@U_Tel", model.U_Tel));
                }
                else
                {
                    parameters.Add(new SqlParameter("@U_Tel",""));
                }
                if (!string.IsNullOrEmpty(model.U_MTel))
                {
                    parameters.Add(new SqlParameter("@U_MTel", model.U_MTel));
                }
                else
                {
                    parameters.Add(new SqlParameter("@U_MTel", ""));
                }
                if (!string.IsNullOrEmpty(model.U_Email))
                {
                    parameters.Add(new SqlParameter("@U_Email", model.U_Email));
                }
                else
                {
                    parameters.Add (new SqlParameter("@U_Email",""));
                }
                if (!string.IsNullOrEmpty(model.Emergency_contact))
                {
                    parameters.Add(new SqlParameter("@Emergency_contact", model.Emergency_contact));
                }
                else
                { 
                    parameters.Add(new SqlParameter("@Emergency_contact", ""));
                }
                if (!string.IsNullOrEmpty(model.Emergency_Tel))
                {
                    parameters.Add(new SqlParameter("@Emergency_Tel",model.Emergency_Tel));
                }
                else
                {
                    parameters.Add(new SqlParameter("@Emergency_Tel", ""));
                }
                if (!string.IsNullOrEmpty(model.Emergency_MTel))
                {
                    parameters.Add(new SqlParameter("@Emergency_MTel", model.Emergency_MTel));
                }
                else
                {
                    parameters.Add(new SqlParameter("@Emergency_MTel", ""));
                }
                if (!string.IsNullOrEmpty(model.School_Level))
                {
                    parameters.Add(new SqlParameter("@School_Level", model.School_Level));
                }
                else
                {
                    parameters.Add(new SqlParameter("@School_Level", ""));
                }
                if (!string.IsNullOrEmpty(model.School_Name))
                {
                    parameters.Add(new SqlParameter("@School_Name",model.School_Name));
                }
                else
                {
                    parameters.Add(new SqlParameter("@School_Name", ""));
                }
                parameters.Add(new SqlParameter("@School_SDate", model.School_SDate));
                parameters.Add(new SqlParameter("@School_EDate", model.School_EDate));
                if (!string.IsNullOrEmpty(model.School_Graduated))
                {
                    parameters.Add(new SqlParameter("@School_Graduated", model.School_Graduated));
                }
                else
                {
                    parameters.Add(new SqlParameter("@School_Graduated", ""));
                }
                if (!string.IsNullOrEmpty(model.School_D_N))
                {
                    parameters.Add(new SqlParameter("@School_D_N", model.School_D_N));
                }
                else
                {
                    parameters.Add(new SqlParameter("@School_D_N", ""));
                }
                if (!string.IsNullOrEmpty(model.School_Major))
                {
                    parameters.Add(new SqlParameter("@School_Major", model.School_Major));
                }
                else
                {
                    parameters.Add(new SqlParameter("@School_Major", ""));
                }
                parameters.Add(new SqlParameter("@U_BC",model.U_BC));
                parameters.Add(new SqlParameter("@U_PFT", model.U_PFT));
                parameters.Add(new SqlParameter("@Role_num", model.Role_num));
                parameters.Add(new SqlParameter("@U_agent_num", model.U_agent_num));
                if (!string.IsNullOrEmpty(model.U_leader_1_num))
                {
                    parameters.Add(new SqlParameter("@U_leader_1_num", model.U_leader_1_num));
                }
                else
                {
                    parameters.Add(new SqlParameter("@U_leader_1_num", ""));
                }
                if (!string.IsNullOrEmpty(model.U_leader_2_num))
                {
                    parameters.Add(new SqlParameter("@U_leader_2_num", model.U_leader_2_num));
                }
                else
                {
                    parameters.Add(new SqlParameter("@U_leader_2_num", ""));
                }
                parameters.Add(new SqlParameter("@U_leader_3_num", ""));
                parameters.Add(new SqlParameter("@U_Check_BC", ""));
                if (!string.IsNullOrEmpty(model.U_address_live))
                {
                    parameters.Add(new SqlParameter("@U_address_live", model.U_address_live));
                }
                else
                {
                    parameters.Add(new SqlParameter("@U_address_live", ""));
                }
                parameters.Add(new SqlParameter("@U_arrive_date", model.U_arrive_date));
                if (model.U_leave_date != null)
                {
                    parameters.Add(new SqlParameter("@U_leave_date", model.U_leave_date));
                }
                else
                {
                    parameters.Add(new SqlParameter("@U_leave_date", DBNull.Value));
                }
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
                return StatusCode(500, resultClass);
            }
        }

        /// <summary>
        /// 密碼變更 Password_Upd/User_edit_psw.asp
        /// </summary>
        [HttpPost("Password_Upd")]
        public ActionResult<ResultClass<string>> Password_Upd(string U_num,int Psw)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var User_Num = HttpContext.Session.GetString("UserID");
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

            try
            {
                ADOData _adoData = new ADOData();
                #region 檢查使用者
                #region SQL
                var parameters_um = new List<SqlParameter>();
                var T_SQL_Um = "select * from User_M where del_tag = '0' AND U_num=@U_num";
                parameters_um.Add(new SqlParameter("@U_num", U_num));
                #endregion
                int result_um = _adoData.ExecuteNonQuery(T_SQL_Um, parameters_um);
                if (result_um == 0)
                {
                    resultClass.ResultCode = "400";
                    resultClass.ResultMsg = "查無資料";
                    return BadRequest(resultClass);
                }
                #endregion

                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = "Update User_M set U_psw=@U_psw,edit_ip=@edit_ip,edit_num=@edit_num,edit_date=@edit_date Where U_num=@U_num";
                parameters.Add(new SqlParameter("@U_psw", Psw));
                parameters.Add(new SqlParameter("@edit_ip", clientIp));
                parameters.Add(new SqlParameter("@edit_num", User_Num));
                parameters.Add(new SqlParameter("@edit_date", DateTime.Now));
                parameters.Add(new SqlParameter("@U_num", U_num));
                #endregion
                int result = _adoData.ExecuteNonQuery(T_SQL, parameters);
                if (result == 0)
                {
                    resultClass.ResultCode = "400";
                    resultClass.ResultMsg = "變更失敗";
                    return BadRequest(resultClass);
                }
                else
                {
                    resultClass.ResultCode = "000";
                    resultClass.ResultMsg = "變更成功";
                    return Ok(resultClass);
                }
            }
            catch (Exception)
            {
                resultClass.ResultCode = "500";
                return StatusCode(500, resultClass);
            }
        }

        /// <summary>
        /// 取得或新增年假管理資料 User_Hday_List/User_Hday_edit.asp
        /// </summary>
        [HttpPost("User_Hday_List")]
        public ActionResult<ResultClass<string>> User_Hday_List(string U_num,DateTime Arrive_Date)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var User_Num = HttpContext.Session.GetString("UserID");
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

            try
            {
                ADOData _adoData = new ADOData();
               
                var parameters = new List<SqlParameter>();
                var T_SQL = "Select * from User_Hday Where del_tag='0' and U_num=@U_num";
                parameters.Add(new SqlParameter("@U_num", U_num));
               
                DataTable dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
                if(dtResult.Rows.Count > 0)
                {
                    resultClass.ResultCode = "000";
                    resultClass.objResult = JsonConvert.SerializeObject(dtResult);
                    return Ok(resultClass);
                }
                else
                {
                    //無資料需新增
                    List<User_Hday> userHdaysList = new List<User_Hday>();
                    for (int i = 1; i < 33; i++)
                    {
                        User_Hday user_Hday = new User_Hday
                        {
                            tbInfo = new tbInfo()
                        };
                        user_Hday.U_num = U_num;
                        user_Hday.tbInfo.del_tag = "0";
                        user_Hday.tbInfo.add_date = DateTime.Now.ToString();
                        user_Hday.tbInfo.add_num = User_Num;
                        user_Hday.tbInfo.add_ip = clientIp;
                        user_Hday.tbInfo.edit_date = DateTime.Now.ToString();
                        user_Hday.tbInfo.edit_num = User_Num;
                        user_Hday.tbInfo.edit_ip = clientIp;
                        user_Hday.H_day_adjust = 0;
                        switch (i)
                        {
                            case 1:
                                user_Hday.H_year_count = "0";
                                user_Hday.H_begin = Arrive_Date;
                                user_Hday.H_end = user_Hday.H_begin.AddMonths(6).AddDays(-1);
                                user_Hday.H_end = new DateTime(user_Hday.H_end.Year, user_Hday.H_end.Month, user_Hday.H_end.Day, 23, 59, 59);
                                user_Hday.H_day_base = 0;
                                user_Hday.H_day_total = user_Hday.H_day_base;
                                break;
                            case 2:
                                user_Hday.H_year_count = "6";
                                user_Hday.H_begin = Arrive_Date.AddMonths(6);
                                user_Hday.H_end = user_Hday.H_begin.AddMonths(6).AddDays(-1);
                                user_Hday.H_end = new DateTime(user_Hday.H_end.Year, user_Hday.H_end.Month, user_Hday.H_end.Day, 23, 59, 59);
                                user_Hday.H_day_base = 3;
                                user_Hday.H_day_total = user_Hday.H_day_base;
                                break;
                            case 3:
                                user_Hday.H_year_count = "12";
                                user_Hday.H_begin = Arrive_Date.AddMonths(12);
                                user_Hday.H_end = user_Hday.H_begin.AddYears(1).AddDays(-1);
                                user_Hday.H_end = new DateTime(user_Hday.H_end.Year, user_Hday.H_end.Month, user_Hday.H_end.Day, 23, 59, 59);
                                user_Hday.H_day_base = 7;
                                user_Hday.H_day_total = user_Hday.H_day_base;
                                break;
                            case 4:
                                user_Hday.H_year_count = "24";
                                user_Hday.H_begin = Arrive_Date.AddMonths(24);
                                user_Hday.H_end = user_Hday.H_begin.AddYears(1).AddDays(-1);
                                user_Hday.H_end = new DateTime(user_Hday.H_end.Year, user_Hday.H_end.Month, user_Hday.H_end.Day, 23, 59, 59);
                                user_Hday.H_day_base = 10;
                                user_Hday.H_day_total = user_Hday.H_day_base;
                                break;
                            case 5:
                                user_Hday.H_year_count = "36";
                                user_Hday.H_begin = Arrive_Date.AddMonths(36);
                                user_Hday.H_end = user_Hday.H_begin.AddYears(1).AddDays(-1);
                                user_Hday.H_end = new DateTime(user_Hday.H_end.Year, user_Hday.H_end.Month, user_Hday.H_end.Day, 23, 59, 59);
                                user_Hday.H_day_base = 14;
                                user_Hday.H_day_total = user_Hday.H_day_base;
                                break;
                            case 6:
                                user_Hday.H_year_count = "48";
                                user_Hday.H_begin = Arrive_Date.AddMonths(48);
                                user_Hday.H_end = user_Hday.H_begin.AddYears(1).AddDays(-1);
                                user_Hday.H_end = new DateTime(user_Hday.H_end.Year, user_Hday.H_end.Month, user_Hday.H_end.Day, 23, 59, 59);
                                user_Hday.H_day_base = 14;
                                user_Hday.H_day_total = user_Hday.H_day_base;
                                break; 
                            case 7:
                                user_Hday.H_year_count = "60";
                                user_Hday.H_begin = Arrive_Date.AddMonths(60);
                                user_Hday.H_end = user_Hday.H_begin.AddYears(1).AddDays(-1);
                                user_Hday.H_end = new DateTime(user_Hday.H_end.Year, user_Hday.H_end.Month, user_Hday.H_end.Day, 23, 59, 59);
                                user_Hday.H_day_base = 15;
                                user_Hday.H_day_total = user_Hday.H_day_base;
                                break;
                            case 8:
                                user_Hday.H_year_count = "72";
                                user_Hday.H_begin = Arrive_Date.AddMonths(72);
                                user_Hday.H_end = user_Hday.H_begin.AddYears(1).AddDays(-1);
                                user_Hday.H_end = new DateTime(user_Hday.H_end.Year, user_Hday.H_end.Month, user_Hday.H_end.Day, 23, 59, 59);
                                user_Hday.H_day_base = 15;
                                user_Hday.H_day_total = user_Hday.H_day_base;
                                break;
                            case 9:
                                user_Hday.H_year_count = "84";
                                user_Hday.H_begin = Arrive_Date.AddMonths(84);
                                user_Hday.H_end = user_Hday.H_begin.AddYears(1).AddDays(-1);
                                user_Hday.H_end = new DateTime(user_Hday.H_end.Year, user_Hday.H_end.Month, user_Hday.H_end.Day, 23, 59, 59);
                                user_Hday.H_day_base = 15;
                                user_Hday.H_day_total = user_Hday.H_day_base;
                                break;
                            case 10:
                                user_Hday.H_year_count = "96";
                                user_Hday.H_begin = Arrive_Date.AddMonths(96);
                                user_Hday.H_end = user_Hday.H_begin.AddYears(1).AddDays(-1);
                                user_Hday.H_end = new DateTime(user_Hday.H_end.Year, user_Hday.H_end.Month, user_Hday.H_end.Day, 23, 59, 59);
                                user_Hday.H_day_base = 15;
                                user_Hday.H_day_total = user_Hday.H_day_base;
                                break;
                            case 11:
                                user_Hday.H_year_count = "108";
                                user_Hday.H_begin = Arrive_Date.AddMonths(108);
                                user_Hday.H_end = user_Hday.H_begin.AddYears(1).AddDays(-1);
                                user_Hday.H_end = new DateTime(user_Hday.H_end.Year, user_Hday.H_end.Month, user_Hday.H_end.Day, 23, 59, 59);
                                user_Hday.H_day_base = 15;
                                user_Hday.H_day_total = user_Hday.H_day_base;
                                break;
                            case 12:
                                user_Hday.H_year_count = "120";
                                user_Hday.H_begin = Arrive_Date.AddMonths(120);
                                user_Hday.H_end = user_Hday.H_begin.AddYears(1).AddDays(-1);
                                user_Hday.H_end = new DateTime(user_Hday.H_end.Year, user_Hday.H_end.Month, user_Hday.H_end.Day, 23, 59, 59);
                                user_Hday.H_day_base = 16;
                                user_Hday.H_day_total = user_Hday.H_day_base;
                                break;
                            case 13:
                                user_Hday.H_year_count = "132";
                                user_Hday.H_begin = Arrive_Date.AddMonths(132);
                                user_Hday.H_end = user_Hday.H_begin.AddYears(1).AddDays(-1);
                                user_Hday.H_end = new DateTime(user_Hday.H_end.Year, user_Hday.H_end.Month, user_Hday.H_end.Day, 23, 59, 59);
                                user_Hday.H_day_base = 17;
                                user_Hday.H_day_total = user_Hday.H_day_base;
                                break;
                            case 14:
                                user_Hday.H_year_count = "144";
                                user_Hday.H_begin = Arrive_Date.AddMonths(144);
                                user_Hday.H_end = user_Hday.H_begin.AddYears(1).AddDays(-1);
                                user_Hday.H_end = new DateTime(user_Hday.H_end.Year, user_Hday.H_end.Month, user_Hday.H_end.Day, 23, 59, 59);
                                user_Hday.H_day_base = 18;
                                user_Hday.H_day_total = user_Hday.H_day_base;
                                break;
                            case 15:
                                user_Hday.H_year_count = "156";
                                user_Hday.H_begin = Arrive_Date.AddMonths(156);
                                user_Hday.H_end = user_Hday.H_begin.AddYears(1).AddDays(-1);
                                user_Hday.H_end = new DateTime(user_Hday.H_end.Year, user_Hday.H_end.Month, user_Hday.H_end.Day, 23, 59, 59);
                                user_Hday.H_day_base = 19;
                                user_Hday.H_day_total = user_Hday.H_day_base;
                                break;
                            case 16:
                                user_Hday.H_year_count = "168";
                                user_Hday.H_begin = Arrive_Date.AddMonths(168);
                                user_Hday.H_end = user_Hday.H_begin.AddYears(1).AddDays(-1);
                                user_Hday.H_end = new DateTime(user_Hday.H_end.Year, user_Hday.H_end.Month, user_Hday.H_end.Day, 23, 59, 59);
                                user_Hday.H_day_base = 20;
                                user_Hday.H_day_total = user_Hday.H_day_base;
                                break;
                            case 17:
                                user_Hday.H_year_count = "180";
                                user_Hday.H_begin = Arrive_Date.AddMonths(180);
                                user_Hday.H_end = user_Hday.H_begin.AddYears(1).AddDays(-1);
                                user_Hday.H_end = new DateTime(user_Hday.H_end.Year, user_Hday.H_end.Month, user_Hday.H_end.Day, 23, 59, 59);
                                user_Hday.H_day_base = 21;
                                user_Hday.H_day_total = user_Hday.H_day_base;
                                break;
                            case 18:
                                user_Hday.H_year_count = "192";
                                user_Hday.H_begin = Arrive_Date.AddMonths(192);
                                user_Hday.H_end = user_Hday.H_begin.AddYears(1).AddDays(-1);
                                user_Hday.H_end = new DateTime(user_Hday.H_end.Year, user_Hday.H_end.Month, user_Hday.H_end.Day, 23, 59, 59);
                                user_Hday.H_day_base = 22;
                                user_Hday.H_day_total = user_Hday.H_day_base;
                                break;
                            case 19:
                                user_Hday.H_year_count = "204";
                                user_Hday.H_begin = Arrive_Date.AddMonths(204);
                                user_Hday.H_end = user_Hday.H_begin.AddYears(1).AddDays(-1);
                                user_Hday.H_end = new DateTime(user_Hday.H_end.Year, user_Hday.H_end.Month, user_Hday.H_end.Day, 23, 59, 59);
                                user_Hday.H_day_base = 23;
                                user_Hday.H_day_total = user_Hday.H_day_base;
                                break;
                            case 20:
                                user_Hday.H_year_count = "216";
                                user_Hday.H_begin = Arrive_Date.AddMonths(216);
                                user_Hday.H_end = user_Hday.H_begin.AddYears(1).AddDays(-1);
                                user_Hday.H_end = new DateTime(user_Hday.H_end.Year, user_Hday.H_end.Month, user_Hday.H_end.Day, 23, 59, 59);
                                user_Hday.H_day_base = 24;
                                user_Hday.H_day_total = user_Hday.H_day_base;
                                break;
                            case 21:
                                user_Hday.H_year_count = "228";
                                user_Hday.H_begin = Arrive_Date.AddMonths(228);
                                user_Hday.H_end = user_Hday.H_begin.AddYears(1).AddDays(-1);
                                user_Hday.H_end = new DateTime(user_Hday.H_end.Year, user_Hday.H_end.Month, user_Hday.H_end.Day, 23, 59, 59);
                                user_Hday.H_day_base = 25;
                                user_Hday.H_day_total = user_Hday.H_day_base;
                                break;
                            case 22:
                                user_Hday.H_year_count = "240";
                                user_Hday.H_begin = Arrive_Date.AddMonths(240);
                                user_Hday.H_end = user_Hday.H_begin.AddYears(1).AddDays(-1);
                                user_Hday.H_end = new DateTime(user_Hday.H_end.Year, user_Hday.H_end.Month, user_Hday.H_end.Day, 23, 59, 59);
                                user_Hday.H_day_base = 26;
                                user_Hday.H_day_total = user_Hday.H_day_base;
                                break;
                            case 23:
                                user_Hday.H_year_count = "252";
                                user_Hday.H_begin = Arrive_Date.AddMonths(252);
                                user_Hday.H_end = user_Hday.H_begin.AddYears(1).AddDays(-1);
                                user_Hday.H_end = new DateTime(user_Hday.H_end.Year, user_Hday.H_end.Month, user_Hday.H_end.Day, 23, 59, 59);
                                user_Hday.H_day_base = 27;
                                user_Hday.H_day_total = user_Hday.H_day_base;
                                break;
                            case 24:
                                user_Hday.H_year_count = "264";
                                user_Hday.H_begin = Arrive_Date.AddMonths(264);
                                user_Hday.H_end = user_Hday.H_begin.AddYears(1).AddDays(-1);
                                user_Hday.H_end = new DateTime(user_Hday.H_end.Year, user_Hday.H_end.Month, user_Hday.H_end.Day, 23, 59, 59);
                                user_Hday.H_day_base = 28;
                                user_Hday.H_day_total = user_Hday.H_day_base;
                                break;
                            case 25:
                                user_Hday.H_year_count = "276";
                                user_Hday.H_begin = Arrive_Date.AddMonths(276);
                                user_Hday.H_end = user_Hday.H_begin.AddYears(1).AddDays(-1);
                                user_Hday.H_end = new DateTime(user_Hday.H_end.Year, user_Hday.H_end.Month, user_Hday.H_end.Day, 23, 59, 59);
                                user_Hday.H_day_base = 29;
                                user_Hday.H_day_total = user_Hday.H_day_base;
                                break;
                            case 26:
                                user_Hday.H_year_count = "288";
                                user_Hday.H_begin = Arrive_Date.AddMonths(288);
                                user_Hday.H_end = user_Hday.H_begin.AddYears(1).AddDays(-1);
                                user_Hday.H_end = new DateTime(user_Hday.H_end.Year, user_Hday.H_end.Month, user_Hday.H_end.Day, 23, 59, 59);
                                user_Hday.H_day_base = 30;
                                user_Hday.H_day_total = user_Hday.H_day_base;
                                break;
                            case 27:
                                user_Hday.H_year_count = "300";
                                user_Hday.H_begin = Arrive_Date.AddMonths(300);
                                user_Hday.H_end = user_Hday.H_begin.AddYears(1).AddDays(-1);
                                user_Hday.H_end = new DateTime(user_Hday.H_end.Year, user_Hday.H_end.Month, user_Hday.H_end.Day, 23, 59, 59);
                                user_Hday.H_day_base = 30;
                                user_Hday.H_day_total = user_Hday.H_day_base;
                                break;
                            case 28:
                                user_Hday.H_year_count = "312";
                                user_Hday.H_begin = Arrive_Date.AddMonths(312);
                                user_Hday.H_end = user_Hday.H_begin.AddYears(1).AddDays(-1);
                                user_Hday.H_end = new DateTime(user_Hday.H_end.Year, user_Hday.H_end.Month, user_Hday.H_end.Day, 23, 59, 59);
                                user_Hday.H_day_base = 30;
                                user_Hday.H_day_total = user_Hday.H_day_base;
                                break;
                            case 29:
                                user_Hday.H_year_count = "324";
                                user_Hday.H_begin = Arrive_Date.AddMonths(324);
                                user_Hday.H_end = user_Hday.H_begin.AddYears(1).AddDays(-1);
                                user_Hday.H_end = new DateTime(user_Hday.H_end.Year, user_Hday.H_end.Month, user_Hday.H_end.Day, 23, 59, 59);
                                user_Hday.H_day_base = 30;
                                user_Hday.H_day_total = user_Hday.H_day_base;
                                break;
                            case 30:
                                user_Hday.H_year_count = "336";
                                user_Hday.H_begin = Arrive_Date.AddMonths(336);
                                user_Hday.H_end = user_Hday.H_begin.AddYears(1).AddDays(-1);
                                user_Hday.H_end = new DateTime(user_Hday.H_end.Year, user_Hday.H_end.Month, user_Hday.H_end.Day, 23, 59, 59);
                                user_Hday.H_day_base = 30;
                                user_Hday.H_day_total = user_Hday.H_day_base;
                                break;
                            case 31:
                                user_Hday.H_year_count = "348";
                                user_Hday.H_begin = Arrive_Date.AddMonths(348);
                                user_Hday.H_end = user_Hday.H_begin.AddYears(1).AddDays(-1);
                                user_Hday.H_end = new DateTime(user_Hday.H_end.Year, user_Hday.H_end.Month, user_Hday.H_end.Day, 23, 59, 59);
                                user_Hday.H_day_base = 30;
                                user_Hday.H_day_total = user_Hday.H_day_base;
                                break;
                            case 32:
                                user_Hday.H_year_count = "360";
                                user_Hday.H_begin = Arrive_Date.AddMonths(360);
                                user_Hday.H_end = user_Hday.H_begin.AddYears(1).AddDays(-1);
                                user_Hday.H_end = new DateTime(user_Hday.H_end.Year, user_Hday.H_end.Month, user_Hday.H_end.Day, 23, 59, 59);
                                user_Hday.H_day_base = 30;
                                user_Hday.H_day_total = user_Hday.H_day_base;
                                break;
                        }
                        userHdaysList.Add(user_Hday);
                    }

                    foreach (var item in userHdaysList)
                    {
                        #region SQL
                        var parameters_in = new List<SqlParameter>();
                        var T_SQL_IN = "Insert into User_Hday (U_num,H_year_count,H_begin,H_end,H_day_base,H_day_adjust,H_day_adjust_note";
                        T_SQL_IN = T_SQL_IN + " ,H_day_total,del_tag,add_date,add_num,add_ip,edit_date,edit_num,edit_ip,H_spent_count)";
                        T_SQL_IN = T_SQL_IN + "  Values ( @U_num,@H_year_count,@H_begin,@H_end,@H_day_base,@H_day_adjust";
                        T_SQL_IN = T_SQL_IN + " ,@H_day_adjust_note,@H_day_total,@del_tag,@add_date,@add_num,@add_ip,";
                        T_SQL_IN = T_SQL_IN + " @edit_date,@edit_num,@edit_ip,@H_spent_count)";
                        parameters_in.Add(new SqlParameter("@U_num", item.U_num));
                        parameters_in.Add(new SqlParameter("@H_year_count", item.H_year_count));
                        parameters_in.Add(new SqlParameter("@H_begin", item.H_begin));
                        parameters_in.Add(new SqlParameter("@H_end", item.H_end));
                        parameters_in.Add(new SqlParameter("@H_day_base", item.H_day_base));
                        parameters_in.Add(new SqlParameter("@H_day_adjust", item.H_day_adjust));
                        parameters_in.Add(new SqlParameter("@H_day_adjust_note", ""));
                        parameters_in.Add(new SqlParameter("@H_day_total", item.H_day_total));
                        parameters_in.Add(new SqlParameter("@del_tag", item.tbInfo.del_tag));
                        parameters_in.Add(new SqlParameter("@add_date", DateTime.Now));
                        parameters_in.Add(new SqlParameter("@add_num", item.tbInfo.add_num));
                        parameters_in.Add(new SqlParameter("@add_ip", item.tbInfo.add_ip));
                        parameters_in.Add(new SqlParameter("@edit_date", DateTime.Now));
                        parameters_in.Add(new SqlParameter("@edit_num", item.tbInfo.edit_num));
                        parameters_in.Add(new SqlParameter("@edit_ip", item.tbInfo.edit_ip));
                        parameters_in.Add(new SqlParameter("@H_spent_count", "0"));
                        #endregion
                        var result_in = _adoData.ExecuteNonQuery(T_SQL_IN, parameters_in);
                    }

                    var parameters_ag = new List<SqlParameter>();
                    var T_SQL_AG = "Select * from User_Hday Where del_tag='0' and U_num=@U_num";
                    parameters_ag.Add(new SqlParameter("@U_num", U_num));

                    DataTable dtResult_ag = _adoData.ExecuteQuery(T_SQL_AG, parameters_ag);
                    if(dtResult_ag.Rows.Count > 0)
                    {
                        resultClass.ResultCode = "000";
                        resultClass.objResult = JsonConvert.SerializeObject(dtResult_ag);
                        return Ok(resultClass);
                    }
                    else
                    {
                        resultClass.ResultCode = "400";
                        resultClass.ResultMsg = "執行異常";
                        return BadRequest(resultClass);
                    }
                }
            }
            catch (Exception)
            {
                resultClass.ResultCode = "500";
                return StatusCode(500, resultClass);
            }
        }

        /// <summary>
        /// 修改年假管理資料 User_Hday_Change/User_Hday_edit.asp
        /// </summary>
        [HttpPost("User_Hday_Change")]
        public ActionResult<ResultClass<string>> User_Hday_Change(List<User_Hday> Modellist)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var User_Num = HttpContext.Session.GetString("UserID");
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

            try
            {
                ADOData _adoData = new ADOData();

                foreach (var item in Modellist)
                {
                    #region SQL
                    var parameters_up = new List<SqlParameter>();
                    var T_SQL_UP = "Update User_Hday set H_begin=@H_begin,H_end=@H_end,H_day_base=@H_day_base";
                    T_SQL_UP = T_SQL_UP + " ,H_day_adjust=@H_day_adjust,H_day_adjust_note=@H_day_adjust_note,H_day_total=@H_day_total ";
                    T_SQL_UP = T_SQL_UP + " ,edit_date=@edit_date,edit_num=@edit_num,edit_ip=@edit_ip Where UH_id=@UH_id";                   
                    parameters_up.Add(new SqlParameter("@H_begin", item.H_begin));
                    parameters_up.Add(new SqlParameter("@H_end", item.H_end));
                    parameters_up.Add(new SqlParameter("@H_day_base", item.H_day_base));
                    parameters_up.Add(new SqlParameter("@H_day_adjust", item.H_day_adjust));
                    if (!string.IsNullOrEmpty(item.H_day_adjust_note))
                    {
                        parameters_up.Add(new SqlParameter("@H_day_adjust_note", item.H_day_adjust_note));
                    }
                    else
                    {
                        parameters_up.Add(new SqlParameter("@H_day_adjust_note", ""));
                    }
                    parameters_up.Add(new SqlParameter("@H_day_total", item.H_day_total));
                    parameters_up.Add(new SqlParameter("@edit_date", DateTime.Now));
                    parameters_up.Add(new SqlParameter("@edit_num", User_Num));
                    parameters_up.Add(new SqlParameter("@edit_ip", clientIp));
                    parameters_up.Add(new SqlParameter("@UH_id", item.UH_id));
                    #endregion
                    int reult_up = _adoData.ExecuteNonQuery(T_SQL_UP, parameters_up);
                    if (reult_up == 0)
                    {
                        resultClass.ResultCode = "400";
                        resultClass.ResultMsg = "變更失敗";
                        return BadRequest(resultClass);
                    }
                }

                resultClass.ResultCode = "000";
                resultClass.ResultMsg = "變更成功";
                return Ok(resultClass);
            }
            catch (Exception)
            {
                resultClass.ResultCode = "500";
                return StatusCode(500, resultClass);
            }
        }

        /// <summary>
        /// 人事異動 User_M_Shanges/User_Form.asp
        /// </summary>
        [HttpPost("User_M_Shanges")]
        public ActionResult<ResultClass<string>> User_M_Shanges(string U_id)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();

                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = "select u_name, School_Name +'-'+School_Major +'-'+SG_NA+'('+SL_NA+')' student";
                T_SQL = T_SQL + " ,convert(varchar(3), (year(U_arrive_date)-1911))+'-'+convert(varchar(2),month(U_arrive_date))+'-'+convert(varchar(2),Day(U_arrive_date)) U_arrive_date";
                T_SQL = T_SQL + " ,convert(varchar(3), (year(U_Birthday)-1911))+'-'+convert(varchar(2),month(U_Birthday))+'-'+convert(varchar(2),Day(U_Birthday)) U_Birthday ";
                T_SQL = T_SQL + " ,M.U_num, PFT_NA,B.BC_NA,N'國峯' as KFDesc FROM User_M M";
                T_SQL = T_SQL + " Left Join";
                T_SQL = T_SQL + " (select item_D_code,item_D_name SL_NA from Item_list where item_M_code = 'school_level' AND item_D_type='Y' AND show_tag='0' AND del_tag='0') S1";
                T_SQL = T_SQL + " on M.School_Level=S1.item_D_code ";
                T_SQL = T_SQL + " Left Join";
                T_SQL = T_SQL + " (select item_D_code ,item_D_name SG_NA from Item_list where item_M_code = 'School_Graduated' AND item_D_type='Y' AND show_tag='0' AND del_tag='0' )S2";
                T_SQL = T_SQL + " on M.School_Graduated=S2.item_D_code ";
                T_SQL = T_SQL + " Left Join  ";
                T_SQL = T_SQL + " (select item_D_code U_PFT ,item_D_name PFT_NA from Item_list where item_M_code = 'professional_title' AND item_D_type='Y' AND show_tag='0' AND del_tag='0' )I";
                T_SQL = T_SQL + " on M.U_PFT=I.U_PFT";
                T_SQL = T_SQL + " Left Join";
                T_SQL = T_SQL + " (select item_D_code U_BC,item_D_name BC_NA from Item_list where item_M_code = 'branch_company' AND item_D_type='Y' AND show_tag='0' AND del_tag='0' )B ";
                T_SQL = T_SQL + " on M.U_BC=B.U_BC ";
                T_SQL = T_SQL + " where M.U_id = @U_id";
                parameters.Add(new SqlParameter("@U_id", U_id));
                #endregion
                DataTable dtresult = _adoData.ExecuteQuery(T_SQL, parameters);
                if(dtresult.Rows.Count > 0)
                {
                    resultClass.ResultCode = "000";
                    resultClass.objResult = JsonConvert.SerializeObject(dtresult);
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
        #endregion

        #region 年假管理
        /// <summary>
        /// 年假列表查詢 User_Hday_LQuery/User_Hday_Remaining.asp
        /// </summary>
        [HttpPost("User_Hday_LQuery")]
        public ActionResult<ResultClass<string>> User_Hday_LQuery(User_Hday_req model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();

                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = "SELECT UM.U_id, UM.U_Num, UM.U_Name,UM.U_arrive_date,UM.U_leave_date,UM.U_cknum,";
                T_SQL = T_SQL + " ( SELECT item_D_name FROM Item_list WHERE item_M_code = 'branch_company' AND item_D_type = 'Y' AND item_D_code = UM.U_BC AND del_tag = '0') AS U_BC_name,";
                T_SQL = T_SQL + " h.H_begin AS H_begin0,h.H_end AS H_end0,h.H_day_base AS H_day_base0,h.H_day_adjust AS H_day_adjust0,";
                T_SQL = T_SQL + " ISNULL(( SELECT SUM(FR_total_hour) FROM Flow_rest FR WHERE FR.del_tag = '0' AND FR.FR_kind = 'FRK005' AND FR.FR_U_num = UM.U_Num  AND FR.FR_date_begin <= h.H_end AND FR.FR_date_end >= h.H_begin), 0) AS rest0,";
                T_SQL = T_SQL + " p.H_begin AS H_begin1,p.H_end AS H_end1,p.H_day_base AS H_day_base1,p.H_day_adjust AS H_day_adjust1,";
                T_SQL = T_SQL + " ISNULL(( SELECT SUM(FR_total_hour) FROM Flow_rest FR WHERE FR.del_tag = '0' AND FR.FR_kind = 'FRK005' AND FR.FR_U_num = UM.U_Num  AND FR.FR_date_begin <= p.H_end AND FR.FR_date_end >= p.H_begin), 0) AS rest1,";
                T_SQL = T_SQL + " p2.H_begin AS H_begin2,p2.H_end AS H_end2,p2.H_day_base AS H_day_base2,p2.H_day_adjust AS H_day_adjust2,";
                T_SQL = T_SQL + " ISNULL(( SELECT SUM(FR_total_hour) FROM Flow_rest FR WHERE FR.del_tag = '0' AND FR.FR_kind = 'FRK005' AND FR.FR_U_num = UM.U_Num  AND FR.FR_date_begin <= p2.H_end AND FR.FR_date_end >= p2.H_begin), 0) AS rest2";
                T_SQL = T_SQL + " FROM User_M UM";
                T_SQL = T_SQL + " LEFT JOIN User_Hday h ON h.del_tag = '0' AND h.U_num = UM.U_num AND GETDATE() BETWEEN h.H_begin AND h.H_end";
                T_SQL = T_SQL + " LEFT JOIN User_Hday p ON h.del_tag = '0' AND p.U_num = UM.U_num AND p.UH_id = h.UH_id - 1";
                T_SQL = T_SQL + " LEFT JOIN User_Hday p2 ON h.del_tag = '0' AND p2.U_num = UM.U_num AND p2.UH_id = h.UH_id - 2";
                T_SQL = T_SQL + " WHERE UM.del_tag = '0' ";
                if(!string.IsNullOrEmpty(model.U_Num_Name)) 
                {
                    T_SQL = T_SQL + " AND (UM.U_num LIKE '%' + @U_Num_Name + '%' OR UM.U_name LIKE '%' + @U_Num_Name + '%') ";
                    parameters.Add(new SqlParameter("@U_Num_Name", model.U_Num_Name));
                }
                if (!string.IsNullOrEmpty(model.U_BC))
                {
                    T_SQL = T_SQL + " AND UM.U_BC = @U_BC";
                    parameters.Add(new SqlParameter("@U_BC", model.U_BC));
                }
                if (!string.IsNullOrEmpty(model.Job_Status))
                {
                    if (model.Job_Status == "Y")
                    {
                        T_SQL = T_SQL + " AND ISNULL(UM.U_leave_date, '') = ''";
                    }
                    else
                    {
                        T_SQL = T_SQL + " AND ISNULL(UM.U_leave_date, '') <> ''";
                    }
                }
                T_SQL = T_SQL + " ORDER BY UM.U_type DESC,UM.U_leave_date,UM.U_id";
                #endregion
                DataTable dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
                if(dtResult.Rows.Count > 0)
                {
                    resultClass.ResultCode = "000";
                    var pageData = FuncHandler.GetPage(dtResult, model.page, 25);
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
            catch (Exception)
            {
                resultClass.ResultCode = "500";
                return StatusCode(500, resultClass);
            }
        }

        #endregion
    }

}
