using KF_WebAPI.BaseClass;
using KF_WebAPI.BaseClass.AE;
using KF_WebAPI.FunctionHandler;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Newtonsoft.Json;
using System.Data;
using System.Linq;
using System.Xml.Linq;

namespace KF_WebAPI.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class AE_ROController : ControllerBase
    {
        /// <summary>
        /// 角色清單查詢 Role_M_LQuery/Role_list.asp
        /// </summary>
        [HttpGet("Role_M_LQuery")]
        public ActionResult<ResultClass<string>> Role_M_LQuery()
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"select R_id,R_num,Case When R_id='100003' Then N'國峯審查' When R_id='100011' Then N'國峯助理' Else R_name END as R_name,del_tag 
                              from Role_M order by R_num";
                #endregion
                var dtResult = _adoData.ExecuteSQuery(T_SQL);
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
        /// 角色明細查詢
        /// </summary>
        [HttpGet("Role_M_SQuery")]
        public ActionResult<ResultClass<string>> Role_M_SQuery(string R_num)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"SELECT RM.R_num,RM.R_name,STUFF((SELECT ' || ' + UM.U_name FROM User_M UM WHERE UM.Role_num = RM.R_num AND UM.U_leave_date IS NULL 
                              FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 4, '') AS U_names,RM.del_tag
                              FROM Role_M RM WHERE RM.R_num = @R_num";
                var parameters = new List<SqlParameter>()
                {
                    new SqlParameter("@R_num",R_num)
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
        /// 新增角色資料/Role_addDB
        /// </summary>
        [HttpPost("Role_M_Ins")]
        public ActionResult<ResultClass<string>> Role_M_Ins(Role_M model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"INSERT INTO Role_M(R_num, R_name, LE_tag, del_tag, add_date, add_num, add_ip)
                              SELECT TOP 1 R_num + 1, @R_name, 'N', '0', GETDATE(), @add_num, @add_ip
                              FROM Role_M
                              ORDER BY R_num DESC";
                var parameters = new List<SqlParameter>()
                {
                    new SqlParameter("@R_name", model.R_name),
                    new SqlParameter("@add_num", model.User),
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
        /// 修改角色資料/Role_editDB
        /// </summary>
        [HttpPost("Role_M_Upd")]
        public ActionResult<ResultClass<string>> Role_M_Upd(Role_M model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"Update Role_M set R_name=@R_name,del_tag=@del_tag,edit_date=getdate(),edit_num=@edit_num,edit_ip=@edit_ip where R_num=@R_num";
                var parameters = new List<SqlParameter>()
                {
                    new SqlParameter("@R_name",model.R_name),
                    new SqlParameter("@del_tag",model.del_tag),
                    new SqlParameter("@edit_num",model.User),
                    new SqlParameter("@edit_ip",clientIp),
                    new SqlParameter("@R_num",model.R_num)
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
        /// 取得角色選單權限 MenuSet_SQuery/Set_Menu_list_one.asp
        /// </summary>
        [HttpGet("MenuSet_SQuery")]
        public ActionResult<ResultClass<string>> MenuSet_SQuery(string roleNum)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"WITH RankedMenu AS ( SELECT MU.Map_id,MT.menu_id,MT.top_id,MT.sub_id,CASE WHEN MT.menu_id = '11089' THEN N'國峯專案' ELSE MT.menu_name END AS menu_name,
                              MU.per_read,MU.per_add,MU.per_edit,MU.per_del,MU.per_query,ROW_NUMBER() OVER (PARTITION BY MT.menu_id ORDER BY MU.Map_id DESC) AS rn 
                              FROM Menu_list MT LEFT JOIN Menu_set MU ON MU.menu_id = MT.menu_id AND MU.del_tag = '0' AND MU.U_num = @roleNum WHERE MT.del_tag = '0' )
                              SELECT Map_id, menu_id, top_id, sub_id, menu_name, per_read, per_add, per_edit, per_del, per_query 
                              FROM RankedMenu WHERE rn = 1 ORDER BY top_id, sub_id";
                var parameters = new List<SqlParameter>() 
                {
                    new SqlParameter("@roleNum",roleNum)
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
        /// 儲存權限設定
        /// </summary>
        [HttpPost("MenuSet_Upd")]
        public ActionResult<ResultClass<string>> MenuSet_Upd(List<MS_PerMission> ModelList)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL_Old = @"WITH RankedMenu AS ( SELECT MU.Map_id,MT.menu_id,MT.top_id,MT.sub_id,CASE WHEN MT.menu_id = '11089' THEN N'國峯專案' ELSE MT.menu_name END AS menu_name,
                                  MU.per_read,MU.per_add,MU.per_edit,MU.per_del,MU.per_query,ROW_NUMBER() OVER (PARTITION BY MT.menu_id ORDER BY MU.Map_id DESC) AS rn 
                                  FROM Menu_list MT LEFT JOIN Menu_set MU ON MU.menu_id = MT.menu_id AND MU.del_tag = '0' AND MU.U_num = @roleNum WHERE MT.del_tag = '0' )
                                  SELECT Map_id, menu_id, top_id, sub_id, menu_name, per_read, per_add, per_edit, per_del, per_query 
                                  FROM RankedMenu WHERE rn = 1 ORDER BY top_id, sub_id";
                var parameters_old = new List<SqlParameter>()
                {
                    new SqlParameter("@roleNum",ModelList[0].R_num)
                };
                List<MS_PerMission> dbData = _adoData.ExecuteQuery(T_SQL_Old, parameters_old).AsEnumerable().Select(row => new MS_PerMission
                {
                    Map_id = row.Field<decimal?>("Map_id"),
                    menu_id = row.Field<decimal>("menu_id"),
                    per_read = row.Field<string>("per_read"),
                    per_add = row.Field<string>("per_add"),
                    per_edit = row.Field<string>("per_edit"),
                    per_del = row.Field<string>("per_del"),
                    per_query = row.Field<string>("per_query"),
                    User = ModelList[0].User
                }).ToList();

                var dbDict = dbData.Where(x => x.Map_id != null).ToDictionary(x => x.Map_id.Value, x => x);

                foreach (var item in ModelList)
                {
                    if(item.Map_id != 0)
                    {
                        if (dbDict.TryGetValue((decimal)item.Map_id, out var existing))
                        {
                            //比對是否有權限變動
                            bool isChanged =
                                item.per_read != existing.per_read ||
                                item.per_add != existing.per_add ||
                                item.per_edit != existing.per_edit ||
                                item.per_del != existing.per_del ||
                                item.per_query != existing.per_query;

                            if (isChanged)
                            {
                                // 權限不同，執行 UPDATE
                                string updateSql = @"UPDATE Menu_set SET per_read = @per_read,per_add = @per_add,per_edit = @per_edit,per_del = @per_del,
                                                 per_query = @per_query,edit_date = getdate(),edit_num = @User,edit_ip=@IP
                                                 WHERE Map_id = @Map_id";
                                var parameters_Upd = new List<SqlParameter>()
                                {
                                    new SqlParameter("@per_read",item.per_read),
                                    new SqlParameter("@per_add",item.per_add),
                                    new SqlParameter("@per_edit",item.per_edit),
                                    new SqlParameter("@per_del",item.per_del),
                                    new SqlParameter("@per_query",item.per_query),
                                    new SqlParameter("@User",ModelList[0].User),
                                    new SqlParameter("@IP",clientIp),
                                    new SqlParameter("@Map_id",item.Map_id)
                                };

                                _adoData.ExecuteNonQuery(updateSql, parameters_Upd);
                            }
                        }
                    }
                    else
                    {
                        // Step 4：資料庫中不存在，新增一筆
                        string insertSql = @"Insert into Menu_set(U_num,U_name,menu_id,menu_name,per_read,per_add,per_edit,per_del,del_tag,add_date,add_num,add_ip) 
                                             select R_num,R_name,@menu_id,@menu_name,@per_read,@per_add,@per_edit,@per_del
                                             ,'0',getdate(),@User,@IP from Role_M where R_num = @roleNum";

                        var parameters_Ins = new List<SqlParameter>()
                        {
                             new SqlParameter("@menu_id",item.menu_id),
                             new SqlParameter("@menu_name",item.menu_name),
                             new SqlParameter("@per_read",item.per_read),
                             new SqlParameter("@per_add",item.per_add),
                             new SqlParameter("@per_edit",item.per_edit),
                             new SqlParameter("@per_del",item.per_del),
                             new SqlParameter("@User",ModelList[0].User),
                             new SqlParameter("@IP",clientIp),
                             new SqlParameter("@roleNum",ModelList[0].R_num)
                        };

                        _adoData.ExecuteNonQuery(insertSql, parameters_Ins);
                    }
                }
                #endregion
                return Ok(resultClass);
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
