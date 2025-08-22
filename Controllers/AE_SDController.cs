using KF_WebAPI.BaseClass;
using KF_WebAPI.BaseClass.AE;
using KF_WebAPI.FunctionHandler;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Newtonsoft.Json;
using System.Data;
using System.Reflection;

namespace KF_WebAPI.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class AE_SDController : ControllerBase
    {
        /// <summary>
        /// 取得尚未列入呆帳清單的呆帳資料
        /// </summary>
        [HttpGet("Fina_BDebt_LQuery")]
        public ActionResult<ResultClass<string>> Fina_BDebt_LQuery()
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"select distinct H.CS_name,H.CS_PID,M.RCM_id,M.sett_AMT,M.amount_total from Receivable_M M
                              Inner join Receivable_D D on M.RCM_id = D.RCM_id
                              Inner join House_apply H on H.HA_id = M.HA_id
                              where D.bad_debt_type='Y' and D.check_pay_type='N' and D.cancel_type='N'
                              and not exists(select 1 from StagnationDebt_M where rcm_id = M.RCM_id and del_tag='0')";
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
        /// 新增呆帳催繳名單
        /// </summary>
        [HttpPost("SD_M_Ins")]
        public ActionResult<ResultClass<string>> SD_M_Ins(StagnationDebt_M model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"Insert into StagnationDebt_M(rcm_id,total_due_amount,total_paid_amount,total_bad_debt,remarks,del_tag,add_date,add_num,add_ip) 
                              Values (@rcm_id,@total_due_amount,@total_paid_amount,@total_bad_debt,@remarks,'0',getdate(),@add_num,@add_ip)";
                var parameters = new List<SqlParameter>() 
                {
                    new SqlParameter("@rcm_id",model.rcm_id),
                    new SqlParameter("@total_due_amount",model.total_due_amount),
                    new SqlParameter("@total_paid_amount",model.total_paid_amount),
                    new SqlParameter("@total_bad_debt",model.total_bad_debt),
                    new SqlParameter("@remarks",model.remarks),
                    new SqlParameter("@add_num",model.tbInfo.add_num),
                    new SqlParameter("@add_ip",clientIp)
                };
                #endregion
                int result = _adoData.ExecuteNonQuery(T_SQL, parameters);
                if (result == 0)
                {
                    resultClass.ResultCode = "400";
                    resultClass.ResultMsg = "儲存失敗";
                    return BadRequest(resultClass);
                }
                else
                {
                    resultClass.ResultCode = "000";
                    resultClass.ResultMsg = "儲存成功";
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
        /// 取得呆帳催繳清單
        /// </summary>
        [HttpGet("SD_M_LQuery")]
        public ActionResult<ResultClass<string>> SD_M_LQuery()
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"select SM.sdm_id,HA.CS_name,HA.CS_PID,FORMAT(SM.total_bad_debt,'N0') as str_total_bad_debt,SM.remarks 
                              ,ISNULL(CAST(YEAR(latest_collection_date) - 1911 AS varchar) + '/' +RIGHT('0' + CAST(MONTH(latest_collection_date) AS varchar), 2) 
                              + '/' +RIGHT('0' + CAST(DAY(latest_collection_date) AS varchar), 2), '') AS collection_date_roc
                              from StagnationDebt_M SM
                              inner join Receivable_M RM on RM.RCM_id = SM.rcm_id
                              inner join House_apply HA on HA.HA_id = RM.HA_id
                              left join (select sdm_id,MAX(collection_date) AS latest_collection_date 
                              FROM StagnationDebt_D where del_tag = '0' GROUP BY sdm_id) AS MAX_SD ON MAX_SD.sdm_id = SM.sdm_id
                              where SM.del_tag = '0' ";
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
        /// 取得呆帳單筆紀錄
        /// </summary>
        [HttpGet("SD_M_SQuery")]
        public ActionResult<ResultClass<string>> SD_M_SQuery(string sdm_id)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"select HA.CS_name,SM.* from StagnationDebt_M SM
                              left join Receivable_M RM on RM.RCM_id = SM.rcm_id
                              left join House_apply HA on HA.HA_id=RM.HA_id where sdm_id=@sdm_id and SM.del_tag ='0' ";
                var parameters = new List<SqlParameter>()
                {
                    new SqlParameter("@sdm_id",sdm_id)
                };
                #endregion
                var dtResult = _adoData.ExecuteQuery(T_SQL,parameters);
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
        /// 修改呆帳催繳名單
        /// </summary>
        [HttpPost("SD_M_Upd")]
        public ActionResult<ResultClass<string>> SD_M_Upd(StagnationDebt_M model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();
            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"Update StagnationDebt_M set total_due_amount=@total_due_amount,total_paid_amount=@total_paid_amount,total_bad_debt=@total_bad_debt,
                             remarks=@remarks,edit_date=getdate(),edit_num=@edit_num,edit_ip=@edit_ip where sdm_id=@sdm_id";
                var parameters = new List<SqlParameter>()
                {
                    new SqlParameter("@sdm_id",model.sdm_id),
                    new SqlParameter("@total_due_amount",model.total_due_amount),
                    new SqlParameter("@total_paid_amount",model.total_paid_amount),
                    new SqlParameter("@total_bad_debt",model.total_bad_debt),
                    new SqlParameter("@remarks",model.remarks),
                    new SqlParameter("@edit_num",model.tbInfo.edit_num),
                    new SqlParameter("@edit_ip",clientIp)
                };
                #endregion
                int result = _adoData.ExecuteNonQuery(T_SQL, parameters);
                if (result == 0)
                {
                    resultClass.ResultCode = "400";
                    resultClass.ResultMsg = "儲存失敗";
                    return BadRequest(resultClass);
                }
                else
                {
                    resultClass.ResultCode = "000";
                    resultClass.ResultMsg = "儲存成功";
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
        /// 取得該筆呆帳催繳所有明細資料
        /// </summary>
        [HttpGet("SD_Info_LQuery")]
        public ActionResult<ResultClass<string>> SD_Info_LQuery(string sdm_id)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"select SM.sdm_id,HA.CS_name,HA.CS_PID,FORMAT(SM.total_due_amount,'N0') as str_total_due_amount
                              ,FORMAT(SM.total_paid_amount,'N0') as str_total_paid_amount,FORMAT(SM.total_bad_debt,'N0') as str_total_bad_debt,SM.remarks 
                              from StagnationDebt_M SM
                              left join Receivable_M RM on RM.RCM_id = SM.rcm_id
                              left join House_apply HA on HA.HA_id=RM.HA_id
                              where sdm_id=@sdm_id and SM.del_tag ='0'";
                var parameters = new List<SqlParameter>()
                {
                    new SqlParameter("@sdm_id",sdm_id)
                };
                var T_SQL_D = @"select ISNULL(CAST(YEAR(collection_date) - 1911 AS varchar) + '/' +RIGHT('0' + CAST(MONTH(collection_date) AS varchar), 2) 
                                + '/' +RIGHT('0' + CAST(DAY(collection_date) AS varchar), 2), '') AS collection_date_roc,
                                FORMAT(payment_amount,'N0') as str_payment_amount,* 
                                from StagnationDebt_D where sdm_id=@sdm_id and del_tag='0'";
                var parameters_d = new List<SqlParameter>()
                {
                    new SqlParameter("@sdm_id",sdm_id)
                };
                #endregion
                var model = _adoData.ExecuteQuery(T_SQL, parameters).AsEnumerable().Select(row => new StagnationDebt_Res 
                {
                    sdm_id = row.Field<int>("sdm_id"),
                    str_total_due_amount = row.Field<string>("str_total_due_amount"),
                    str_total_paid_amount = row.Field<string>("str_total_paid_amount"),
                    str_total_bad_debt = row.Field<string>("str_total_bad_debt"),
                    SD_Name = row.Field<string>("CS_name"),
                    remarks = row.Field<string>("remarks"),
                    SD_CID = row.Field<string>("CS_PID")
                }).FirstOrDefault();

                model.SDList = _adoData.ExecuteQuery(T_SQL_D, parameters_d).AsEnumerable().Select(row => new StagnationDebt_D 
                {
                    sdd_id = row.Field<int>("sdd_id"),
                    str_payment_amount = row.Field<string>("str_payment_amount"),
                    collection_date_roc = row.Field<string>("collection_date_roc"),
                    collection_not = row.Field<string>("collection_not")
                }).ToList();

                resultClass.ResultCode = "000";
                resultClass.objResult = JsonConvert.SerializeObject(model);
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
        /// 新增催繳明細資料
        /// </summary>
        [HttpPost("SD_D_Ins")]
        public ActionResult<ResultClass<string>> SD_D_Ins(StagnationDebt_D model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"Insert into StagnationDebt_D (sdm_id,payment_amount,collection_date,collection_not,del_tag,add_date,add_num,add_ip)
                              Values (@sdm_id,@payment_amount,@collection_date,@collection_not,'0',getdate(),@add_num,@add_ip)";
                var parameters = new List<SqlParameter>() 
                {
                    new SqlParameter("@sdm_id",model.sdm_id),
                    new SqlParameter("@payment_amount", model.payment_amount ?? (object)DBNull.Value),
                    new SqlParameter("@collection_date",FuncHandler.ConvertROCToGregorian(model.collection_date_roc)),
                    new SqlParameter("@collection_not",model.collection_not),
                    new SqlParameter("@add_num",model.tbInfo.add_num),
                    new SqlParameter("@add_ip",clientIp)
                };
                #endregion
                int result = _adoData.ExecuteNonQuery(T_SQL, parameters);
                if (result == 0)
                {
                    resultClass.ResultCode = "400";
                    resultClass.ResultMsg = "儲存失敗";
                    return BadRequest(resultClass);
                }
                else
                {
                    resultClass.ResultCode = "000";
                    resultClass.ResultMsg = "儲存成功";
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
        /// 修改催繳明細資料
        /// </summary>
        [HttpPost("SD_D_Upd")]
        public ActionResult<ResultClass<string>> SD_D_Upd(StagnationDebt_D model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"Update StagnationDebt_D set payment_amount=@payment_amount,collection_date=@collection_date,collection_not=@collection_not, 
                              edit_date=getdate(),edit_num=@edit_num,edit_ip=@edit_ip
                              where sdd_id=@sdd_id";
                var parameters = new List<SqlParameter>()
                {
                    new SqlParameter("@sdd_id",model.sdd_id),
                    new SqlParameter("@payment_amount", model.payment_amount ?? (object)DBNull.Value),
                    new SqlParameter("@collection_date",FuncHandler.ConvertROCToGregorian(model.collection_date_roc)),
                    new SqlParameter("@collection_not",model.collection_not),
                    new SqlParameter("@edit_num",model.tbInfo.edit_num),
                    new SqlParameter("@edit_ip",clientIp)
                };
                #endregion
                int result = _adoData.ExecuteNonQuery(T_SQL, parameters);
                if (result == 0)
                {
                    resultClass.ResultCode = "400";
                    resultClass.ResultMsg = "儲存失敗";
                    return BadRequest(resultClass);
                }
                else
                {
                    resultClass.ResultCode = "000";
                    resultClass.ResultMsg = "儲存成功";
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
        /// 取得呆帳催繳單筆紀錄
        /// </summary>
        [HttpGet("SD_D_SQuery")]
        public ActionResult<ResultClass<string>> SD_D_SQuery(string sdd_id)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"select ISNULL(CAST(YEAR(collection_date) - 1911 AS varchar) + '-' +RIGHT('0' + CAST(MONTH(collection_date) AS varchar), 2) 
                                + '-' +RIGHT('0' + CAST(DAY(collection_date) AS varchar), 2), '') AS collection_date_roc,
                                FORMAT(payment_amount,'N0') as str_payment_amount,* 
                                from StagnationDebt_D where sdd_id=@sdd_id";
                var parameters = new List<SqlParameter>()
                {
                    new SqlParameter("@sdd_id",sdd_id)
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
        /// 刪除呆帳催繳單筆紀錄
        /// </summary>
        [HttpGet("SD_D_Del")]
        public ActionResult<ResultClass<string>> SD_D_Del(string User,string sdd_id)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"Update StagnationDebt_D set del_tag='1', del_date=getdate(),del_num=@del_num,del_ip=@del_ip where sdd_id=@sdd_id";
                var parameters = new List<SqlParameter>()
                {
                    new SqlParameter("@del_num",User),
                    new SqlParameter("@del_ip",clientIp),
                    new SqlParameter("@sdd_id",sdd_id)
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
        /// 刪除呆帳紀錄
        /// </summary>
        [HttpGet("SD_M_Del")]
        public ActionResult<ResultClass<string>> SD_M_Del(string User, string sdm_id)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"Update StagnationDebt_M set del_tag='1', del_date=getdate(),del_num=@del_num,del_ip=@del_ip where sdm_id=@sdm_id";
                var parameters = new List<SqlParameter>()
                {
                    new SqlParameter("@del_num",User),
                    new SqlParameter("@del_ip",clientIp),
                    new SqlParameter("@sdm_id",sdm_id)
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
