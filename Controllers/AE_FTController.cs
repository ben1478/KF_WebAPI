using Azure;
using KF_WebAPI.BaseClass;
using KF_WebAPI.BaseClass.AE;
using KF_WebAPI.DataLogic;
using KF_WebAPI.FunctionHandler;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Newtonsoft.Json;
using System.Data;
using System.Reflection;
using System.Reflection.Metadata.Ecma335;
using static Microsoft.Extensions.Logging.EventSource.LoggingEventSource;

namespace KF_WebAPI.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class AE_FTController : ControllerBase
    {
        AE_FT _Ft = new AE_FT();

        #region 業績折扣標準設定
        /// <summary>
        /// 取得業績折扣標準設定資料 Feat_M_LQuery/FR_M_query.asp
        /// bcType: general 6區 BC0900 數位行銷部
        /// </summary>
        [HttpGet("Feat_M_LQuery")]
        public ActionResult<ResultClass<string>> Feat_M_LQuery(string bcType)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                resultClass = _Ft.Feat_M_LQuery(bcType);
                resultClass.ResultCode = "000";
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
            model.tbInfo.add_ip = clientIp;

            try
            {
                resultClass = _Ft.Feat_M_Ins(model);
                if (resultClass.ResultCode == "000")
                {
                    return Ok(resultClass);
                }
                else
                {
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
        /// 刪除申貸方案
        /// </summary>
        [HttpGet("Feat_M_Del")]
        public ActionResult<ResultClass<string>> Feat_M_Del(string FR_M_code, string user, string bcType)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

            try
            {
                resultClass = _Ft.Feat_M_Del(FR_M_code, user, bcType, clientIp);
                if (resultClass.ResultCode == "000")
                {
                    return Ok(resultClass);
                }
                else
                {
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
        /// 取得方案折扣設定
        /// </summary>
        [HttpGet("Feat_D_LQuery")]
        public ActionResult<ResultClass<string>> Feat_D_LQuery(string FR_M_code, string bcType)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                resultClass = _Ft.Feat_D_LQuery(FR_M_code, bcType);
                resultClass.ResultCode = "000";
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
        /// 刪除單筆折扣
        /// </summary>
        [HttpGet("Feat_D_Del")]
        public ActionResult<ResultClass<string>> Feat_D_Del(string FR_id, string user)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

            try
            {
                resultClass = _Ft.Feat_D_Del(FR_id, user, clientIp);
                if (resultClass.ResultCode == "000")
                {
                    return Ok(resultClass);
                }
                else
                {
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
        /// 變更方案折扣設定
        /// </summary>
        [HttpPost("Feat_D_Upd")]
        public ActionResult<ResultClass<string>> Feat_D_Upd(List<Feat_D> modelList)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

            foreach (var item in modelList) 
            {
                item.tbInfo.add_ip = clientIp;
            }
            try
            {
                resultClass = _Ft.Feat_D_Upd(modelList);
                if (resultClass.ResultCode == "000")
                {
                    return Ok(resultClass);
                }
                else
                {
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
        #endregion

        #region 佣金標準設定
        /// <summary>
        /// 取得新鑫佣金標準設定資料 Feat_NM_LQuery/FR_M_query.asp?IsComm=Y&menuID=11084
        /// </summary>
        [HttpGet("Feat_NM_LQuery")]
        public ActionResult<ResultClass<string>> Feat_NM_LQuery()
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            try
            {
                resultClass = _Ft.Feat_NM_LQuery();
                resultClass.ResultCode = "000";
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
        /// 新增新鑫申貸方案 Feat_M_Ins/FR_M_add.asp?IsComm=Y
        /// </summary>
        [HttpPost("Feat_NM_Ins")]
        public ActionResult<ResultClass<string>> Feat_NM_Ins(Feat_M_req model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();
            model.tbInfo.add_ip = clientIp;

            try
            {
                resultClass = _Ft.Feat_NM_Ins(model);
                if (resultClass.ResultCode == "000")
                {
                    return Ok(resultClass);
                }
                else
                {
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
        /// 刪除新鑫申貸方案
        /// </summary>
        [HttpGet("Feat_NM_Del")]
        public ActionResult<ResultClass<string>> Feat_NM_Del(string FR_M_code, string user, string bcType)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

            try
            {
                resultClass = _Ft.Feat_NM_Del(FR_M_code, user, bcType, clientIp);
                if (resultClass.ResultCode == "000")
                {
                    return Ok(resultClass);
                }
                else
                {
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
        /// 取得新鑫方案折扣設定
        /// </summary>
        [HttpGet("Feat_ND_LQuery")]
        public ActionResult<ResultClass<string>> Feat_ND_LQuery(string FR_M_code, string bcType)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                resultClass = _Ft.Feat_ND_LQuery(FR_M_code, bcType);
                resultClass.ResultCode = "000";
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
        /// 刪除新鑫單筆折扣
        /// </summary>
        [HttpGet("Feat_ND_Del")]
        public ActionResult<ResultClass<string>> Feat_ND_Del(string FR_id, string user)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

            try
            {
                resultClass = _Ft.Feat_ND_Del(FR_id, user, clientIp);
                if (resultClass.ResultCode == "000")
                {
                    return Ok(resultClass);
                }
                else
                {
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
        /// 變更新鑫方案折扣設定
        /// </summary>
        [HttpPost("Feat_ND_Upd")]
        public ActionResult<ResultClass<string>> Feat_ND_Upd(List<Feat_D> modelList)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

            foreach (var item in modelList)
            {
                item.tbInfo.add_ip = clientIp;
            }
            try
            {
                resultClass = _Ft.Feat_ND_Upd(modelList);
                if (resultClass.ResultCode == "000")
                {
                    return Ok(resultClass);
                }
                else
                {
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
        /// 取得國峯佣金設定
        /// </summary>
        [HttpGet("Feat_KF_LQuery")]
        public ActionResult<ResultClass<string>> Feat_KF_LQuery()
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                resultClass = _Ft.Feat_KF_LQuery();
                resultClass.ResultCode = "000";
                return Ok(resultClass);
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass);
            }
        }
        #endregion



    }
}
