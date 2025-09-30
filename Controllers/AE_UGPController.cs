using KF_WebAPI.BaseClass;
using KF_WebAPI.BaseClass.AE;
using KF_WebAPI.FunctionHandler;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Newtonsoft.Json;
using System.Data;

namespace KF_WebAPI.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class AE_UGPController : ControllerBase
    {
        /// <summary>
        /// 取得分組清單 User_G_LQuery/group_M_query.asp
        /// </summary>
        [HttpGet("User_G_LQuery")]
        public ActionResult<ResultClass<string>> User_G_LQuery()
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"select M.group_M_code,M.group_id,M.group_M_title,M.group_M_name,CAST(YEAR(M.group_start_day) - 1911 AS VARCHAR) 
                              + '/' + RIGHT('0' + CAST(MONTH(M.group_start_day) AS VARCHAR), 2) 
                              + '/' + RIGHT('0' + CAST(DAY(M.group_start_day) AS VARCHAR), 2) AS group_start_day_minguo,
                              CAST(YEAR(isnull(M.group_end_day,getdate()+1)) - 1911 AS VARCHAR) 
                              + '/' + RIGHT('0' + CAST(MONTH(isnull(M.group_end_day,getdate()+1)) AS VARCHAR), 2) 
                              + '/' + RIGHT('0' + CAST(DAY(isnull(M.group_end_day,getdate()+1)) AS VARCHAR), 2) AS group_end_day_minguo,
                              STUFF((SELECT ' || ' + ISNULL(Li.item_D_name,d.group_D_name) From User_group d LEFT JOIN Item_list Li ON Li.item_D_txt_A = d.group_D_code 
                              where d.group_D_type = 'Y' and d.del_tag <> '1'
                              and d.group_M_id = M.group_id FOR XML PATH(''), TYPE ).value('.', 'NVARCHAR(MAX)'), 1, 4, '') AS show_U_name
                              from User_group M
                              where M.del_tag <> '1' AND M.group_M_type = 'Y'
                              order by M.del_tag,M.group_sort,M.group_end_day,M.group_start_day,M.group_id";
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
        /// 新增組長 User_G_Ins/group_M_edit.asp
        /// </summary>
        [HttpPost("User_G_Ins")]
        public ActionResult<ResultClass<string>> User_G_Ins(User_group_Ins model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"Insert into User_group (group_M_title,group_cknum,group_M_type,group_M_code,group_M_name,group_D_type,group_D_code,group_D_name
                              ,group_D_color,group_D_txt_A,group_D_txt_B,group_sort,del_tag,show_tag,add_date,add_num,add_ip,group_start_day,group_end_day)
                              Values (@group_M_title,@group_cknum,'Y',@group_M_code,@group_M_name,'N','',''
                              ,'','','',@group_sort,'0','0',GETDATE(),@add_num,@add_ip,@group_start_day,@group_end_day)";
                var parameters = new List<SqlParameter>()
                {
                    new SqlParameter("@group_M_title",model.group_M_title),
                    new SqlParameter("@group_cknum",FuncHandler.GetCheckNum()),
                    new SqlParameter("@group_M_code",model.group_M_code),
                    new SqlParameter("@group_M_name",model.group_M_name),
                    new SqlParameter("@group_sort",model.group_sort),
                    new SqlParameter("@add_num",model.tbInfo.add_num),
                    new SqlParameter("@add_ip",clientIp),
                    new SqlParameter("@group_start_day",FuncHandler.ConvertROCToGregorian(model.str_group_start_day)),
                    new SqlParameter("@group_end_day",FuncHandler.ConvertROCToGregorian(model.str_group_end_day))
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
        /// 取得組長 User_UG_M_SQuery/group_M_edit.asp
        /// </summary>
        [HttpGet("User_UG_M_SQuery")]
        public ActionResult<ResultClass<string>> User_UG_M_SQuery(string group_id)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"select group_M_title,group_sort,CAST(YEAR(group_start_day) - 1911 AS VARCHAR) + '-' + RIGHT('0' + CAST(MONTH(group_start_day) AS VARCHAR), 2) 
                              + '-' + RIGHT('0' + CAST(DAY(group_start_day) AS VARCHAR), 2) AS group_start_day_minguo,
                              CAST(YEAR(isnull(group_end_day,getdate()+1)) - 1911 AS VARCHAR) + '-' + RIGHT('0' + CAST(MONTH(isnull(group_end_day,getdate()+1)) AS VARCHAR), 2) 
                              + '-' + RIGHT('0' + CAST(DAY(isnull(group_end_day,getdate()+1)) AS VARCHAR), 2) AS group_end_day_minguo,
                              group_M_code,group_M_name from User_group where group_id = @group_id";
                var parameters = new List<SqlParameter>()
                {
                    new SqlParameter("@group_id",group_id)
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
        /// 修改組長 User_G_Ins/group_M_edit.asp
        /// </summary>
        [HttpPost("User_G_Upd")]
        public ActionResult<ResultClass<string>> User_G_Upd(User_group_Upd model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"Update User_group set group_M_title=@group_M_title,group_sort=@group_sort,group_start_day=@group_start_day,
                              group_end_day=@group_end_day,group_M_code=@group_M_code,group_M_name=@group_M_name,edit_date=getdate(),edit_num=@edit_num,edit_ip=@edit_ip 
                              where group_id=@group_id";
                var parameters = new List<SqlParameter>()
                {
                    new SqlParameter("@group_M_title",model.group_M_title),
                    new SqlParameter("@group_M_code",model.group_M_code),
                    new SqlParameter("@group_M_name",model.group_M_name),
                    new SqlParameter("@group_sort",model.group_sort),
                    new SqlParameter("@edit_num",model.tbInfo.edit_num),
                    new SqlParameter("@edit_ip",clientIp),
                    new SqlParameter("@group_start_day",FuncHandler.ConvertROCToGregorian(model.str_group_start_day)),
                    new SqlParameter("@group_end_day",FuncHandler.ConvertROCToGregorian(model.str_group_end_day)),
                     new SqlParameter("@group_id",model.group_id)
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
        /// 刪除分組資料
        /// </summary>
        [HttpGet("User_G_Del")]
        public ActionResult<ResultClass<string>> User_G_Del(string group_M_code,string user)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"Update User_group Set del_tag = '1', del_date = Getdate(), del_ip = @IP , del_num=@user 
                              Where group_M_code IN (select group_M_code from User_group where group_id=@group_M_code ) and del_tag='0' ";
                var parameters = new List<SqlParameter>()
                {
                    new SqlParameter("@IP",clientIp),
                    new SqlParameter("@user",user),
                    new SqlParameter("@group_M_code",group_M_code)
                };
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
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass);
            }
        }
        /// <summary>
        /// 取得組員清單
        /// </summary>
        [HttpGet("User_GP_D_LQuery")]
        public ActionResult<ResultClass<string>> User_GP_D_LQuery(string group_id)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var _FuncHandler = new FuncHandler();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"select group_id,group_M_code,group_M_name,group_D_code,ISNULL(Li.item_D_name,d.group_D_name) as group_D_name,group_start_day,group_end_day 
                              from User_group d left join Item_list Li ON Li.item_D_txt_A = d.group_D_code
                              where group_M_code IN (select group_M_code from User_group where group_id=@group_id) and d.del_tag <> '1' and group_D_type='Y' ";
                var parameters = new List<SqlParameter>()
                {
                    new SqlParameter("@group_id",group_id)
                };
                #endregion
                var dtResult = _adoData.ExecuteQuery(T_SQL,parameters).AsEnumerable().Select(row => new
                {
                    group_id = row.Field<decimal>("group_id").ToString(),
                    group_M_code = row.Field<string>("group_M_code"),
                    group_M_name = row.Field<string>("group_M_name"),
                    group_D_code = row.Field<string>("group_D_code"),
                    group_D_name = row.Field<string>("group_D_name"),
                    group_start_day = row.Field<DateTime>("group_start_day").ToString("yyyy-MM-dd"),
                    group_end_day = row.Field<DateTime>("group_end_day").ToString("yyyy-MM-dd")
                });
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
        /// 修改組員
        /// </summary>
        [HttpPost("User_D_Upd")]
        public ActionResult<ResultClass<string>> User_D_Upd(List<User_group_DUpd> ModelList)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL_OLD
                var T_SQL_OLD = @"select group_id,group_M_code,group_M_name,group_D_code,ISNULL(Li.item_D_name,d.group_D_name) as group_D_name,group_start_day,group_end_day 
                              from User_group d left join Item_list Li ON Li.item_D_txt_A = d.group_D_code
                              where group_M_code IN (select group_M_code from User_group where group_id=@group_id) and d.del_tag <> '1' and group_D_type='Y' ";
                var parameters_old = new List<SqlParameter>()
                {
                    new SqlParameter("@group_id",ModelList[0].group_M_id)
                };
                #endregion
                List<User_group_DUpd> dbData = _adoData.ExecuteQuery(T_SQL_OLD, parameters_old).AsEnumerable().Select(row => new User_group_DUpd
                {
                    group_id = row.Field<decimal>("group_id"),
                    group_M_code = row.Field<string>("group_M_code"),
                    group_M_name = row.Field<string>("group_M_name"),
                    group_D_code = row.Field<string>("group_D_code"),
                    group_D_name = row.Field<string>("group_D_name"),
                    group_start_day = FuncHandler.ConvertGregorianToROC(row.Field<DateTime>("group_start_day").ToString("yyyy/MM/dd")),
                    group_end_day = FuncHandler.ConvertGregorianToROC(row.Field<DateTime>("group_end_day").ToString("yyyy/MM/dd"))
                }).ToList();
                var dbDict = dbData.ToDictionary(x => x.group_id, x => x);
                var delTag = 0;
                foreach (var item in ModelList)
                {
                    if (!string.IsNullOrEmpty(item.group_D_code))
                    {
                        if (dbDict.TryGetValue(item.group_id, out var existing))
                        {
                            bool isChanged =
                                     item.group_D_code != existing.group_D_code ||
                                     item.group_D_name != existing.group_D_name ||
                                     item.group_start_day != existing.group_start_day ||
                                     item.group_end_day != existing.group_end_day;

                            if (isChanged)
                            {
                                //在異動的同時要同時變更狀態 del_tag(0=現在,2=未來,3=過去) 依據期間判定 
                                if (item.group_start_day != existing.group_start_day || item.group_end_day != existing.group_end_day)
                                {
                                    DateTime startDate = Convert.ToDateTime(FuncHandler.ConvertROCToGregorian(item.group_start_day)).Date;
                                    DateTime endDate = Convert.ToDateTime(FuncHandler.ConvertROCToGregorian(item.group_end_day)).Date;
                                    DateTime today = DateTime.Today;
                                    if (startDate > today)
                                    {
                                        delTag = 1;
                                    }
                                    else if (endDate <= today)
                                    {
                                        delTag = 3;
                                    }
                                    else
                                    {
                                        delTag = 0;
                                    }
                                }
                                #region SQL_Upd
                                var T_SQL = @"Update User_group Set group_M_code = @group_M_code,group_M_name = @group_M_name,group_D_code = @group_D_code, 
                                          group_D_name = @group_D_name,group_start_day =@group_start_day,group_end_day = @group_end_day,
                                          del_tag = @del_tag,edit_num = @User,edit_date = Getdate(),edit_ip = @IP
                                          Where group_id = @group_id";
                                var parameters = new List<SqlParameter>()
                                {
                                    new SqlParameter("@group_M_code",item.group_M_code),
                                    new SqlParameter("@group_M_name",item.group_M_name),
                                    new SqlParameter("@group_D_code",item.group_D_code),
                                    new SqlParameter("@group_D_name",item.group_D_name),
                                    new SqlParameter("@group_start_day",Convert.ToDateTime(FuncHandler.ConvertROCToGregorian(item.group_start_day))),
                                    new SqlParameter("@group_end_day",Convert.ToDateTime(FuncHandler.ConvertROCToGregorian(item.group_end_day))),
                                    new SqlParameter("@del_tag",delTag),
                                    new SqlParameter("@User",item.User),
                                    new SqlParameter("@IP",clientIp),
                                    new SqlParameter("@group_id",item.group_id)
                                };
                                #endregion
                                int result = _adoData.ExecuteNonQuery(T_SQL, parameters);
                                if (result == 0)
                                {
                                    resultClass.ResultCode = "400";
                                    resultClass.ResultMsg = "修改失敗";
                                    return BadRequest(resultClass);
                                }
                            }
                        }
                        else
                        {
                            #region SQL_Ins
                            var T_SQL = @"Insert into User_group (group_cknum,group_M_id,group_M_type,group_M_code,group_M_name,group_D_type,group_D_code,group_D_name,group_start_day
                                              ,group_end_day,del_tag,show_tag,add_date,add_num,add_ip)
                                              Values (@group_cknum,@group_M_id,'N',@group_M_code,@group_M_name,'Y',@group_D_code,@group_D_name,@group_start_day,@group_end_day
                                              ,@del_tag,0,Getdate(),@User,@IP)";
                            var parameters = new List<SqlParameter>()
                                {
                                    new SqlParameter("@group_cknum",FuncHandler.GetCheckNum()),
                                    new SqlParameter("@group_M_id",item.group_M_id),
                                    new SqlParameter("@group_M_code",item.group_M_code),
                                    new SqlParameter("@group_M_name",item.group_M_name),
                                    new SqlParameter("@group_D_code",item.group_D_code),
                                    new SqlParameter("@group_D_name",item.group_D_name),
                                    new SqlParameter("@group_start_day",Convert.ToDateTime(FuncHandler.ConvertROCToGregorian(item.group_start_day))),
                                    new SqlParameter("@group_end_day",Convert.ToDateTime(FuncHandler.ConvertROCToGregorian(item.group_end_day))),
                                    new SqlParameter("@del_tag",delTag),
                                    new SqlParameter("@User",item.User),
                                    new SqlParameter("@IP",clientIp)
                                };
                            #endregion
                            int result = _adoData.ExecuteNonQuery(T_SQL, parameters);
                            if (result == 0)
                            {
                                resultClass.ResultCode = "400";
                                resultClass.ResultMsg = "新增失敗";
                                return BadRequest(resultClass);
                            }
                        }
                    }
                }
                resultClass.ResultCode = "000";
                resultClass.ResultMsg = "儲存成功";
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
        /// 刪除組員
        /// </summary>
        [HttpGet("User_D_Del")]
        public ActionResult<ResultClass<string>> User_D_Del(string group_id, string user)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"Update User_group Set del_tag = '1', del_date = Getdate(), del_ip = @IP , del_num = @user Where group_id = @group_id";
                var parameters = new List<SqlParameter>()
                {
                    new SqlParameter("@group_id",group_id),
                    new SqlParameter("@user",user),
                    new SqlParameter("@IP",clientIp)
                };
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
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass);
            }
        }
    }
}
