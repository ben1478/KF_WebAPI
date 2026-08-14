using KF_WebAPI.BaseClass;
using KF_WebAPI.DataLogic;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace KF_WebAPI.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class AE_CCNController : Controller
    {
        private AE_CNN _CCN = new AE_CNN();

        //電銷可回收名單
        public ActionResult<ResultClass<string>> Recycle_M_Lquery()
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
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
