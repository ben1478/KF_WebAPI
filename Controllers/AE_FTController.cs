using KF_WebAPI.BaseClass;
using KF_WebAPI.BaseClass.AE;
using KF_WebAPI.FunctionHandler;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Newtonsoft.Json;
using System.Data;
using System.Reflection.Metadata.Ecma335;
using static Microsoft.Extensions.Logging.EventSource.LoggingEventSource;

namespace KF_WebAPI.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class AE_FTController : ControllerBase
    {
        /// <summary>
        /// 取得業績折扣標準設定資料 Feat_M_LQuery/FR_M_query.asp
        /// </summary>
        [HttpGet("Feat_M_LQuery")]
        public ActionResult<ResultClass<string>> Feat_M_LQuery(string bcType) 
        {
            //bcType: general 6區 BC0900 數位行銷部
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                var _Fun = new FuncHandler();
                #region SQL_M
                var T_SQL_M = @"select * from Feat_rule where del_tag = '0' AND FR_M_type = 'Y' and U_BC=@U_BC order by FR_sort,FR_id";
                var parameters_m = new List<SqlParameter>()
                {
                     new SqlParameter("@U_BC",bcType)
                };
                #endregion
                var resultMList = _adoData.ExecuteQuery(T_SQL_M, parameters_m).AsEnumerable().Select(row => new Feat_M
                {
                    FR_id = row.Field<decimal>("FR_id"),
                    FR_M_code = row.Field<string>("FR_M_code"),
                    FR_M_name = _Fun.DeCodeBNWords(row.Field<string>("FR_M_name"))
                }).ToList();

                foreach (var item in resultMList) 
                {
                    #region SQL_D
                    var T_SQL_D = @"select * from Feat_rule where del_tag = '0' AND FR_D_type = 'Y' AND FR_M_code = @FR_M_code and U_BC=@U_BC
                                order by FR_D_ratio_A desc,FR_D_rate desc,FR_D_discount";
                    var parameters_d = new List<SqlParameter>()
                    {
                        new SqlParameter("@FR_M_code",item.FR_M_code),
                        new SqlParameter("@U_BC",bcType)
                    };
                    #endregion
                    var resultDList = _adoData.ExecuteQuery(T_SQL_D,parameters_d).AsEnumerable().Select(row => new Feat_D 
                    {
                        FR_D_ratio_A = row.Field<decimal>("FR_D_ratio_A"),
                        FR_D_ratio_B = row.Field<decimal>("FR_D_ratio_B"),
                        FR_D_rate = row.Field<string>("FR_D_rate"),
                        FR_D_discount = row.Field<string>("FR_D_discount"),
                        FR_D_replace = row.Field<string>("FR_D_replace")
                    }).ToList();

                    //合併
                    item.feat_Ds.AddRange(resultDList);
                }

                resultClass.ResultCode = "000";
                resultClass.objResult = JsonConvert.SerializeObject(resultMList);
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
        /// 新增申貸方案
        /// </summary>
        [HttpPost("Feat_M_Ins")]
        public ActionResult<ResultClass<string>> Feat_M_Ins(Feat_M_req model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"Insert into Feat_rule (FR_cknum,FR_M_type,FR_M_code,FR_M_name,FR_D_type,FR_D_code,FR_D_name,FR_D_rate,FR_D_discount,FR_D_replace,
                              FR_sort,del_tag,show_tag,add_date,add_num,add_ip,U_BC)
                              Values (@FR_cknum,'Y',@FR_M_code,@FR_M_name,'N','','','','','',99999,0,0,Getdate(),@add_num,@add_ip,@U_BC)";
                var parameters = new List<SqlParameter>()
                {
                    new SqlParameter("@FR_cknum",FuncHandler.GetCheckNum()),
                    new SqlParameter("@FR_M_code",model.FR_M_code),
                    new SqlParameter("@FR_M_name",model.FR_M_name),
                    new SqlParameter("@add_num",model.tbInfo.add_num),
                    new SqlParameter("@add_ip",clientIp),
                    new SqlParameter("@U_BC",model.U_BC)
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
        /// 刪除申貸方案
        /// </summary>
        [HttpGet("Feat_M_Del")]
        public ActionResult<ResultClass<string>> Feat_M_Del(string FR_M_code,string user,string bcType)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"Update Feat_rule Set del_tag = '1',del_date=Getdate(),del_num=@User,del_ip=@IP  where FR_M_code=@FR_M_code and del_tag='0' and U_BC=@u_bc";
                var parameters = new List<SqlParameter>() 
                {
                    new SqlParameter("@User",user),
                    new SqlParameter("@IP",clientIp),
                    new SqlParameter("@FR_M_code",FR_M_code),
                    new SqlParameter("@u_bc",bcType)
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
        /// 取得方案折扣設定
        /// </summary>
        [HttpGet("Feat_D_LQuery")]
        public ActionResult<ResultClass<string>> Feat_D_LQuery(string FR_M_code,string bcType)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL_D
                var T_SQL = @"select * from Feat_rule where del_tag = '0' AND FR_D_type = 'Y' AND FR_M_code = @FR_M_code and U_BC=@U_BC
                                order by FR_D_ratio_A desc,FR_D_rate desc,FR_D_discount";
                var parameters = new List<SqlParameter>()
                {
                    new SqlParameter("@FR_M_code",FR_M_code),
                    new SqlParameter("@U_BC",bcType)
                };
                #endregion
                var resultList = _adoData.ExecuteQuery(T_SQL, parameters).AsEnumerable().Select(row => new
                {
                    FR_id = row.Field<decimal>("FR_id"),
                    FR_D_ratio_A = row.Field<decimal>("FR_D_ratio_A"),
                    FR_D_ratio_B = row.Field<decimal>("FR_D_ratio_B"),
                    FR_D_rate = row.Field<string>("FR_D_rate"),
                    FR_D_discount = row.Field<string>("FR_D_discount"),
                    FR_D_replace = row.Field<string>("FR_D_replace")
                }).ToList();

                resultClass.ResultCode = "000";
                resultClass.objResult = JsonConvert.SerializeObject(resultList);
                return Ok(resultClass);
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass);
            }
        }

        //TODO儲存方案折扣設定
    }
}
