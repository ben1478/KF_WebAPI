
using Azure.Core;
using KF_WebAPI.BaseClass;
using KF_WebAPI.BaseClass.WebRobot;
using KF_WebAPI.DataLogic;
using KF_WebAPI.FunctionHandler;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;

namespace KF_WebAPI.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class WebRobotController : ControllerBase
    {
        private Common _Comm = new();
        private WebRobot _WebRobot = new();

        [Route("InsertWebRobot_M")]
        [HttpPost]
        public ActionResult<ResultClass<BaseResult>> InsertWebRobot_M([FromBody] WebRobot_M objects)
        {
           
            ResultClass<string> resultClass = new ResultClass<string>();
           
            try
            {
                _WebRobot.InsertWebRobot_M(objects);
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "999";
                resultClass.ResultMsg = "更新失敗;" + ex.Message;
            }
            return Ok(resultClass);
        }


        [Route("InsertWebRobot_D")]
        [HttpPost]
        public ActionResult<ResultClass<BaseResult>> InsertWebRobot_D([FromBody] WebRobot_D objects)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            try
            {
                _WebRobot.InsertWebRobot_D(objects);
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "999";
                resultClass.ResultMsg = "更新失敗;" + ex.Message;
            }
            return Ok(resultClass);
        }


    }
}
