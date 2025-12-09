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
                var _Fun = new FuncHandler();
                #region SQL
                var T_SQL = @"select distinct H.CS_name,H.CS_PID,M.RCM_id,M.sett_AMT,M.amount_total,RD.RemainingPrincipal 
                              from Receivable_M M
                              Inner join Receivable_D D on M.RCM_id = D.RCM_id
                              Inner join House_apply H on H.HA_id = M.HA_id
                              OUTER APPLY ( SELECT TOP 1 RemainingPrincipal FROM Receivable_D WHERE RCM_id = M.RCM_id AND check_pay_type = 'Y' 
                              ORDER BY RC_count DESC ) RD
                              where D.bad_debt_type='Y' and D.check_pay_type='N' and D.cancel_type='N'
                              and not exists(select 1 from StagnationDebt_M where rcm_id = M.RCM_id and del_tag='0')";
                #endregion
                var result = _adoData.ExecuteSQuery(T_SQL).AsEnumerable().Select(row => new
                {
                    CS_name = _Fun.DeCodeBNWords(row.Field<string>("CS_name")),
                    CS_PID = row.Field<string>("CS_PID"),
                    RCM_id = row.Field<decimal>("RCM_id"),
                    sett_AMT = row.Field<int>("sett_AMT"),
                    amount_total = row.Field<decimal>("amount_total"),
                    RemainingPrincipal = row.Field<decimal>("RemainingPrincipal")
                }).ToList();
                resultClass.ResultCode = "000";
                resultClass.objResult = JsonConvert.SerializeObject(result);
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
                var T_SQL = @"Insert into StagnationDebt_M(RCM_id,amount_total,RemainingPrincipal,total_due_amount,total_paid_amount,total_bad_debt,remarks,is_close,del_tag,add_date,add_num,add_ip) 
                              Values (@RCM_id,@amount_total,@RemainingPrincipal,@total_due_amount,@total_paid_amount,@total_bad_debt,@remarks,'N','0',getdate(),@add_num,@add_ip)";
                var parameters = new List<SqlParameter>() 
                {
                    new SqlParameter("@RCM_id",model.RCM_id),
                    new SqlParameter("@amount_total",model.amount_total),
                    new SqlParameter("@RemainingPrincipal",model.RemainingPrincipal),
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
                var _Fun = new FuncHandler();
                #region SQL
                var T_SQL = @"SELECT SM.sdm_id,HA.CS_name,HA.CS_PID,
                              FORMAT(SM.total_bad_debt - ISNULL(PD.total_payment, 0), 'N0') AS str_total_bad_debt,SM.remarks,
                              ISNULL(CAST(YEAR(MAX_SD.latest_collection_date) - 1911 AS varchar) + '/' 
                              + RIGHT('0' + CAST(MONTH(MAX_SD.latest_collection_date) AS varchar), 2) + '/' 
                              + RIGHT('0' + CAST(DAY(MAX_SD.latest_collection_date) AS varchar), 2), 
                              '') AS collection_date_roc,
                              FORMAT(SM.amount_total, 'N0') AS str_amount_total,
                              FORMAT(SM.RemainingPrincipal, 'N0') AS str_RemainingPrincipal,
                              CASE WHEN SM.is_close = 'Y' THEN '已結清' ELSE '未結清' END AS str_close
                              FROM StagnationDebt_M SM
                              INNER JOIN Receivable_M RM ON RM.RCM_id = SM.rcm_id
                              INNER JOIN House_apply HA ON HA.HA_id = RM.HA_id
                              LEFT JOIN (SELECT sdm_id,ISNULL(SUM(payment_amount),0) AS total_payment FROM StagnationDebt_D WHERE del_tag = '0' GROUP BY sdm_id
                              ) PD ON PD.sdm_id = SM.sdm_id
                              LEFT JOIN (SELECT sdm_id, MAX(collection_date) AS latest_collection_date FROM StagnationDebt_D WHERE del_tag = '0' GROUP BY sdm_id
                              ) MAX_SD ON MAX_SD.sdm_id = SM.sdm_id
                              WHERE SM.del_tag = '0'";
                #endregion
                var result = _adoData.ExecuteSQuery(T_SQL).AsEnumerable().Select(row => new
                {
                    sdm_id = row.Field<int>("sdm_id"),
                    CS_name = _Fun.DeCodeBNWords(row.Field<string>("CS_name")),
                    CS_PID = row.Field<string>("CS_PID"),
                    str_total_bad_debt = row.Field<string>("str_total_bad_debt"),
                    remarks = row.Field<string>("remarks"),
                    collection_date_roc = row.Field<string>("collection_date_roc"),
                    str_amount_total = row.Field<string>("str_amount_total"),
                    str_RemainingPrincipal = row.Field<string>("str_RemainingPrincipal"),
                    str_close = row.Field<string>("str_close")
                }).ToList();
                resultClass.ResultCode = "000";
                resultClass.objResult = JsonConvert.SerializeObject(result);
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
                var _Fun = new FuncHandler();
                #region SQL
                var T_SQL = @"select HA.CS_name,SM.* from StagnationDebt_M SM
                              left join Receivable_M RM on RM.RCM_id = SM.rcm_id
                              left join House_apply HA on HA.HA_id=RM.HA_id where sdm_id=@sdm_id and SM.del_tag ='0' ";
                var parameters = new List<SqlParameter>()
                {
                    new SqlParameter("@sdm_id",sdm_id)
                };
                #endregion
                var result = _adoData.ExecuteQuery(T_SQL,parameters).AsEnumerable().Select(row => new
                {
                    CS_name = _Fun.DeCodeBNWords(row.Field<string>("CS_name")),
                    total_due_amount = row.Field<decimal>("total_due_amount"),
                    total_paid_amount = row.Field<decimal>("total_paid_amount"),
                    total_bad_debt = row.Field<decimal>("total_bad_debt"),
                    remarks = row.Field<string>("remarks")
                }).ToList();
                resultClass.ResultCode = "000";
                resultClass.objResult = JsonConvert.SerializeObject(result);
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
                var _Fun = new FuncHandler();
                #region SQL
                var T_SQL = @"select SM.sdm_id,HA.CS_name,HA.CS_PID,SM.remarks
                              ,FORMAT(SM.total_due_amount,'N0') as str_total_due_amount
                              ,FORMAT(SM.total_paid_amount + ISNULL(PD.total_payment,0),'N0') as str_total_paid_amount
                              ,FORMAT(SM.total_bad_debt - ISNULL(PD.total_payment,0),'N0') as str_total_bad_debt
                              ,FORMAT(SM.amount_total,'N0') as str_amount_total
                              ,FORMAT(SM.RemainingPrincipal,'N0') as str_RemainingPrincipal
                              from StagnationDebt_M SM
                              left join Receivable_M RM on RM.RCM_id = SM.rcm_id
                              left join House_apply HA on HA.HA_id=RM.HA_id
                              LEFT JOIN (SELECT sdm_id,ISNULL(SUM(payment_amount),0) AS total_payment FROM StagnationDebt_D WHERE del_tag = '0' GROUP BY sdm_id
                              ) PD ON PD.sdm_id = SM.sdm_id
                              where SM.sdm_id=@sdm_id and SM.del_tag ='0'";
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
                    str_amount_total = row.Field<string>("str_amount_total"),
                    str_RemainingPrincipal = row.Field<string>("str_RemainingPrincipal"),
                    str_total_due_amount = row.Field<string>("str_total_due_amount"),
                    str_total_paid_amount = row.Field<string>("str_total_paid_amount"),
                    str_total_bad_debt = row.Field<string>("str_total_bad_debt"),
                    SD_Name = _Fun.DeCodeBNWords(row.Field<string>("CS_name")),
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
                    #region 判定是否結清且修改狀態
                    var T_SQL_I = @"UPDATE StagnationDebt_M SET is_close = 'Y', edit_date = GETDATE(),edit_num = 'sys',edit_ip = '::1' WHERE sdm_id = @sdm_id AND del_tag = '0'
                                    AND ( SELECT ISNULL(SUM(payment_amount), 0) FROM StagnationDebt_D WHERE del_tag = '0' AND sdm_id = @sdm_id ) = total_bad_debt";
                    var parameters_i = new List<SqlParameter>()
                    {
                        new SqlParameter("@sdm_id",model.sdm_id)
                    };
                    _adoData.ExecuteQuery(T_SQL_I, parameters_i);
                    #endregion
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
                    #region 判定是否結清且修改狀態
                    var T_SQL_I = @"UPDATE StagnationDebt_M SET is_close = 'Y', edit_date = GETDATE(),edit_num = 'sys',edit_ip = '::1' WHERE sdm_id = @sdm_id AND del_tag = '0'
                                    AND ( SELECT ISNULL(SUM(payment_amount), 0) FROM StagnationDebt_D WHERE del_tag = '0' AND sdm_id = @sdm_id ) = total_bad_debt";
                    var parameters_i = new List<SqlParameter>()
                    {
                        new SqlParameter("@sdm_id",model.sdm_id)
                    };
                    _adoData.ExecuteQuery(T_SQL_I, parameters_i);
                    #endregion
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
        /// <summary>
        /// 取得業務所有呆帳資料
        /// </summary>
        [HttpGet("SD_Plnum_LQuery")]
        public ActionResult<ResultClass<string>> SD_Plnum_LQuery(string planNum)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"
                    select HA.CS_name,SM.amount_total - ISNULL(SD.total_payment,0) AS amount_bad_total from StagnationDebt_M SM
                    LEFT JOIN (SELECT sdm_id,ISNULL(SUM(payment_amount),0) AS total_payment FROM StagnationDebt_D WHERE del_tag = '0'GROUP BY sdm_id
                    ) SD ON SD.sdm_id = SM.sdm_id
                    Inner join Receivable_M RM on RM.RCM_id = SM.RCM_id
                    Inner join House_apply HA on HA.HA_id = RM.HA_id
                    Inner join House_sendcase HS on HS.HS_id = RM.HS_id
                    where SM.del_tag='0' and HA.plan_num=@plan_num";
                var parameters = new List<SqlParameter>()
                {
                    new SqlParameter("@plan_num",planNum)
                };
                #endregion
                var dtResult=_adoData.ExecuteQuery(T_SQL, parameters);
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
        /// 裕富機車貸款查詢
        /// </summary>
        /// <param name="get_amount_date_S"></param>
        /// <param name="get_amount_date_E"></param>
        /// <returns></returns>
        [HttpGet("SD_MotoList_LQuery")]
        public ActionResult<ResultClass<string>> SD_MotoList_LQuery(string get_amount_date_S, string get_amount_date_E)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"select  HA.ha_id,R.rcm_id,LEFT(HA.cs_name, 1) + 'XX' AS cs_name,pro_name, HS.interest_rate_pass+'%' rate, isnull(HS.get_amount,'')+'萬'get_amount,isnull(format(HS.get_amount_date,'yyyy/MM/dd'),'')get_amount_date
                    ,R.month_total,R.amount_per_month
                    from House_sendcase HS
                    LEFT JOIN House_apply HA ON HS.HA_id = HA.HA_id　LEFT JOIN House_pre_project HP ON HP.HP_project_id = HS.HP_project_id
                    Left Join ( SELECT item_D_name pro_name ,item_D_code FROM Item_list WHERE item_M_code = 'project_title'　AND item_D_type='Y' AND del_tag='0'
                    ) P on P.item_D_code = HP.project_title
                    left join Receivable_M R on HA.HA_id=R.HA_id
                    where HS.del_tag='0'and HA.del_tag='0'and R.del_tag='0'and get_amount_type='GTAT002'
                    AND HP.project_title IN('PJ00046','PJ00047') and format(HS.get_amount_date,'yyyy/MM/dd') between @get_amount_date_S and @get_amount_date_E  order by get_amount_date";
                var parameters = new List<SqlParameter>()
                {
                    new SqlParameter("@get_amount_date_S",get_amount_date_S),
                    new SqlParameter("@get_amount_date_E",get_amount_date_E)
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
        /// 機車貸款繳款明細
        /// </summary>
        /// <param name="RCM_id"></param>
        /// <returns></returns>
        [HttpGet("SD_MotoRC_LQuery")]
        public ActionResult<ResultClass<string>> SD_MotoRC_LQuery(string RCM_id)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"select *,
                        case when delayday >= 3 And delayday <= 6  then 
		                        100+CEILING( RC_amount *0.16/365*delayday)
	                        when delayday >= 7 And delayday <= 14  then 
		                        200+CEILING( RC_amount *0.16/365*delayday)
	                        when delayday > 14  then 
		                        300+CEILING( RC_amount *0.16/365*delayday)
	                        else 0
	                        end Delaymoney
                        from (
                        SELECT M.amount_total,RC_amount ,D.rc_count,format(D.RC_date,'yyyy/MM/dd')rc_date
                        ,isnull(format(D.check_pay_date,'yyyy/MM/dd'),'')check_pay_date 
                        ,CASE WHEN RecPayAmt IS NULL THEN ''
                              ELSE convert(varchar, RecPayAmt)
                         END recpayamt,CASE WHEN RC_date < SYSDATETIME() AND check_pay_date IS NULL THEN 
                        isnull(DATEDIFF(DAY, RC_date, SYSDATETIME()), 0)
	                        ELSE isnull(DATEDIFF(DAY, RC_date, check_pay_date), 0)
                        END delayday,D.ex_remainingPrincipal,D.rc_note                                                                                                 
                        FROM Receivable_M M left join Receivable_D D on M.RCM_id=D.RCM_id
                        where M.del_tag='0' and D.del_tag='0' and  M.RCM_id=@RCM_id ) M";
                var parameters = new List<SqlParameter>()
                {
                    new SqlParameter("@RCM_id",RCM_id)
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


    }
}
