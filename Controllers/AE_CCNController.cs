using KF_WebAPI.BaseClass;
using KF_WebAPI.DataLogic;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using static KF_WebAPI.BaseClass.AE.Telemarketing;

namespace KF_WebAPI.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class AE_CCNController : Controller
    {
        private AE_CNN _CCN = new AE_CNN();

        //電銷可回收名單
        [HttpPost("Recycle_M_Lquery")]
        public ActionResult<ResultClass<string>> Recycle_M_Lquery(Recycle_req model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                resultClass = _CCN.Recycle_M_Lquery(model);
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

        [HttpGet("Call_Detail_LQuery")]
        public ActionResult<ResultClass<string>> Call_Detail_LQuery(decimal tmID)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                resultClass = _CCN.Call_Detail_LQuery(tmID);
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

        [HttpPost("Recycle_M_Upd")]
        public ActionResult<ResultClass<string>> Recycle_M_Upd(string user,[FromBody] List<string> ids)
        {

            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                resultClass = _CCN.Recycle_M_Upd(user,ids);
                resultClass.ResultCode = "000";
                resultClass.ResultMsg = "回收成功";
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
