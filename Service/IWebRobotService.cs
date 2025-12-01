using KF_WebAPI.BaseClass.WebRobot;
using KF_WebAPI.BaseClass;
using KF_WebAPI.DataLogic;
using Newtonsoft.Json;

namespace KF_WebAPI.Service
{
    public interface IWebRobotService
    {
        ResultClass<int> InsertWebRobot_M(WebRobot_M objects);
        ResultClass<int> InsertWebRobot_D(WebRobot_D objects);
        Task<ResultClass<List<string>>> GetProxyFlowAsync();
        ResultClass<List<string>> GetSysKeyWords();
        ResultClass<RemoteAction> GetRemoteAction(string ComputerInfo);
    }

    public class WebRobotService : IWebRobotService
    {
        private readonly WebRobot _webRobot;

        public WebRobotService()
        {
            _webRobot = new WebRobot();
        }

        public ResultClass<int> InsertWebRobot_M(WebRobot_M objects)
        {
            return _webRobot.InsertWebRobot_M(objects);
        }

        public ResultClass<int> InsertWebRobot_D(WebRobot_D objects)
        {
            return _webRobot.InsertWebRobot_D(objects);
        }

        public ResultClass<List<string>> GetSysKeyWords()
        {
            return _webRobot.GetSysKeyWords();
        }
        public ResultClass<RemoteAction> GetRemoteAction(string ComputerInfo)
        {
            return _webRobot.GetRemoteAction(ComputerInfo);
        }

       

        public async Task<ResultClass<List<string>>> GetProxyFlowAsync()
        {
            return await _webRobot.GetProxyFlowAsync();
        }
    }
}
