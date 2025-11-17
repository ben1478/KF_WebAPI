
using Azure.Core;
using KF_WebAPI.BaseClass;
using KF_WebAPI.BaseClass.WebRobot;
using KF_WebAPI.DataLogic;
using KF_WebAPI.FunctionHandler;
using KF_WebAPI.Service;
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

            ResultClass<int> resultClass = new ResultClass<int>();

            try
            {
                resultClass= _WebRobot.InsertWebRobot_M(objects);
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
            ResultClass<int> resultClass = new ResultClass<int>();
            try
            {
                resultClass= _WebRobot.InsertWebRobot_D(objects);
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "999";
                resultClass.ResultMsg = "更新失敗;" + ex.Message;
            }
            return Ok(resultClass);
        }

        [HttpPost("GetProxyFlowAsync")]
        public async Task<ActionResult<ResultClass<List<string>>>> GetProxyFlowAsync()
        {
            var result = await _WebRobot.GetProxyFlowAsync();

            if (result.ResultCode == "000")
                return Ok(result);
            else
                return BadRequest(result);
        }




    }
}
