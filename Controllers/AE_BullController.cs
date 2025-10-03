using KF_WebAPI.BaseClass;
using KF_WebAPI.BaseClass.AE;
using KF_WebAPI.FunctionHandler;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Newtonsoft.Json;
using System;
using System.Data;
using static Microsoft.Extensions.Logging.EventSource.LoggingEventSource;

namespace KF_WebAPI.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class AE_BullController : ControllerBase
    {
        /// <summary>
        /// 佈告欄清單查詢 Bulletin_M_LQuery/Bulletin_list.asp
        /// </summary>
        [HttpGet("Bulletin_M_LQuery")]
        public ActionResult<ResultClass<string>> Bulletin_M_LQuery(string? keyWord)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                var _Fun = new FuncHandler();
                #region SQL
                var T_SQL = @"select id,FORMAT(notice_date,'yyyy/MM/dd') AS formatted_notice_date,um.U_name,title,notice_mode,notice_type,cknum
                              ,(case when '置頂'=notice_mode then 0 else 1 end) as notice_mode_sort
                              ,(SELECT COUNT(*) FROM ASP_UpLoad WHERE cknum = bn.cknum and del_tag='0') AS cknum_count
                              from Bulletin bn left join User_M um on um.U_num = bn.add_num where bn.del_tag = '0' ";
                if(!string.IsNullOrEmpty(keyWord) ) 
                {
                    T_SQL += " and title like @keyword";
                }
                T_SQL += " order by notice_mode_sort, notice_date desc";
                var parameters = new List<SqlParameter>()
                {
                     new SqlParameter("@keyword",'%'+ keyWord +'%')
                };
                #endregion
                var result = _adoData.ExecuteQuery(T_SQL, parameters).AsEnumerable().Select(row => new
                {
                    id = row.Field<int>("id"),
                    notice_date = row.Field<string>("formatted_notice_date"),
                    U_name = _Fun.DeCodeBNWords(row.Field<string>("U_name")),
                    title = _Fun.DeCodeBNWords(row.Field<string>("title")),
                    notice_mode = row.Field<string>("notice_mode"),
                    notice_type = row.Field<string>("notice_type"),
                    cknum = row.Field<string>("cknum"),
                    cknum_count = row.Field<int>("cknum_count")
                });
                if (result.Count() > 0)
                {
                    resultClass.ResultCode = "000";
                    resultClass.objResult = JsonConvert.SerializeObject(result);
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
        /// 新增佈告欄資料
        /// </summary>
        [HttpPost("Bulletin_M_Ins")]
        public ActionResult<ResultClass<string>> Bulletin_M_Ins(Bulletin_Ins model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"Insert into Bulletin(cknum,title,content,notice_type,notice_mode,notice_date,sort,del_tag,add_date,add_num,add_ip)
                              Values (@cknum,@title,@content,@notice_type,@notice_mode,@notice_date,0,0,Getdate(),@add_num,@add_ip)";
                var parameters = new List<SqlParameter>()
                {
                    new SqlParameter("@cknum",FuncHandler.GetCheckNum()),
                    new SqlParameter("@title",model.title),
                    new SqlParameter("@content",model.bulletin_content),
                    new SqlParameter("@notice_type",model.notice_type),
                    new SqlParameter("@notice_mode",model.notice_mode),
                    new SqlParameter("@notice_date",FuncHandler.ConvertROCToGregorian(model.notice_date)),
                    new SqlParameter("@add_num",model.tbInfo.add_num),
                    new SqlParameter("@add_ip",clientIp)
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
        /// 取得佈告欄資料
        /// </summary>
        [HttpGet("Bulletin_M_SQuery")]
        public ActionResult<ResultClass<string>> Bulletin_M_SQuery(string id)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                var _Fun = new FuncHandler();
                #region SQL
                var T_SQL = @"select FORMAT(notice_date,'yyyy/MM/dd') AS formatted_notice_date,um.U_name,title,notice_mode,notice_type,cknum,content
                              from Bulletin bn left join User_M um on um.U_num = bn.add_num where bn.del_tag = '0' and id=@id ";
                var parameters = new List<SqlParameter>()
                {
                     new SqlParameter("@id",id)
                };
                #endregion
                var result = _adoData.ExecuteQuery(T_SQL, parameters).AsEnumerable().Select(row => new
                {
                    notice_date = FuncHandler.ConvertGregorianToROC(row.Field<string>("formatted_notice_date")),
                    U_name = _Fun.DeCodeBNWords(row.Field<string>("U_name")),
                    title = _Fun.DeCodeBNWords(row.Field<string>("title")),
                    notice_mode = row.Field<string>("notice_mode"),
                    notice_type = row.Field<string>("notice_type"),
                    cknum = row.Field<string>("cknum"),
                    content = row.Field<string>("content")
                });
                if (result.Count() > 0)
                {
                    resultClass.ResultCode = "000";
                    resultClass.objResult = JsonConvert.SerializeObject(result);
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
        /// 修改佈告欄資料
        /// </summary>
        [HttpPost("Bulletin_M_Upd")]
        public ActionResult<ResultClass<string>> Bulletin_M_Upd(Bulletin_Ins model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"Update Bulletin Set title=@title,content=@content,notice_type=@notice_type,notice_mode=@notice_mode,edit_date=Getdate(),
                              notice_date=@notice_date,edit_num=@User,edit_ip=@IP where id=@ID";
                var parameters = new List<SqlParameter>()
                {
                    new SqlParameter("@title",model.title),
                    new SqlParameter("@content",model.bulletin_content),
                    new SqlParameter("@notice_type",model.notice_type),
                    new SqlParameter("@notice_mode",model.notice_mode),
                    new SqlParameter("@notice_date",FuncHandler.ConvertROCToGregorian(model.notice_date)),
                    new SqlParameter("@User",model.tbInfo.add_num),
                    new SqlParameter("@IP",clientIp),
                    new SqlParameter("@ID",model.id)
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
        /// 刪除佈告欄
        /// </summary>
        [HttpGet("Bull_M_Del")]
        public ActionResult<ResultClass<string>> Bull_M_Del(string id,string user)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"Update Bulletin Set del_tag = '1', del_date = Getdate(), del_ip = @IP , del_num = @user Where id = @id";
                var parameters = new List<SqlParameter>()
                {
                    new SqlParameter("@id",id),
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
