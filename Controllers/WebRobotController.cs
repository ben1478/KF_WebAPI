
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
        private readonly IWebRobotService _webRobotService;
       
        // 透過建構子注入介面
        public WebRobotController(IWebRobotService webRobotService)
        {
            _webRobotService = webRobotService;
        }

        [Route("InsertWebRobot_M")]
        [HttpPost]
        public ActionResult<ResultClass<BaseResult>> InsertWebRobot_M([FromBody] WebRobot_M objects)
        {

            ResultClass<int> resultClass = new ResultClass<int>();

            try
            {
                resultClass= _webRobotService.InsertWebRobot_M(objects);
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
                resultClass= _webRobotService.InsertWebRobot_D(objects);
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "999";
                resultClass.ResultMsg = "更新失敗;" + ex.Message;
            }
            return Ok(resultClass);
        }

        [HttpPost("GetProxyFlowAsync")]
        public async Task<ActionResult<List<string>>> GetProxyFlowAsync()
        {
            var result = await _webRobotService.GetProxyFlowAsync();

            return Ok(result);
        }

        [HttpPost("GetRemoteAction")]
        public  ActionResult<ResultClass<BaseResult>> GetRemoteAction([FromBody] string objects)
        {
            var result =  _webRobotService.GetRemoteAction(objects);

            return Ok(result);
        }

        [HttpPost("GetSysKeyWords")]
        public ActionResult<ResultClass<List<string>>> GetSysKeyWords()
        {
            var result =  _webRobotService.GetSysKeyWords();

            return Ok(result);
        }




    }
}
