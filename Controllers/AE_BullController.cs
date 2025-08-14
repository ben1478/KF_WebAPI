using KF_WebAPI.BaseClass;
using KF_WebAPI.FunctionHandler;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Newtonsoft.Json;
using System;

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
                #region SQL
                var T_SQL = @"select id,FORMAT(notice_date,'yyyy/MM/dd') AS formatted_notice_date,um.U_name,title,notice_mode,notice_type,cknum
                              ,(case when '置頂'=notice_mode then 0 else 1 end) as notice_mode_sort
                              from Bulletin bn left join User_M um on um.U_num = bn.add_num";
                if(!string.IsNullOrEmpty(keyWord) ) 
                {
                    T_SQL += " where title like @keyword";
                }
                T_SQL += " order by notice_mode_sort, notice_date desc";
                var parameters = new List<SqlParameter>()
                {
                     new SqlParameter("@keyword",'%'+ keyWord +'%')
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
    }
}
