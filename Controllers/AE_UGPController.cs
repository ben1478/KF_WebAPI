using KF_WebAPI.BaseClass;
using KF_WebAPI.BaseClass.AE;
using KF_WebAPI.FunctionHandler;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Newtonsoft.Json;

namespace KF_WebAPI.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class AE_UGPController : ControllerBase
    {
        /// <summary>
        /// 取得員工分組清單 User_G_LQuery/group_M_query.asp
        /// </summary>
        [HttpGet("User_G_LQuery")]
        public ActionResult<ResultClass<string>> User_G_LQuery()
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"select M.group_id,M.group_M_title,M.group_M_name,CAST(YEAR(M.group_start_day) - 1911 AS VARCHAR) 
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
    }
}
