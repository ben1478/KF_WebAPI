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


        private const string API_LINK_TEMPLATE = "https://tq.lunaproxy.com/getflowip?neek=1807058&num=30&regions=tw&ip_si=1&level=1&sb=";

        public async Task<ResultClass<List<string>>> GetProxyFlowAsync()
        {
            ResultClass<List<string>> resultClass = new();

            try
            {
                string json = "";
                using (HttpClient client = new HttpClient())
                {
                    json = await client.GetStringAsync(API_LINK_TEMPLATE);
                }

                List<string> proxies = new List<string>();

                try
                {
                    var result = JsonConvert.DeserializeObject<ProxyResponse>(json);
                    string ErrMessage = "";
                    if (result != null)
                    {
                        ErrMessage = "ErrorCode：" + result.code.ToString() + ";ErrMsg" + result.msg;
                    }
                    resultClass.ResultCode = "999";
                    resultClass.ResultMsg = ErrMessage;
                    resultClass.objResult = null;
                }
                catch (Newtonsoft.Json.JsonException)
                {
                    proxies = json.Split(new char[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                }

                resultClass.ResultCode = "000";
                resultClass.ResultMsg = "";
                resultClass.objResult = proxies;
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "999";
                resultClass.ResultMsg = ex.Message;
                resultClass.objResult = null;
            }

            return resultClass;
        }

    }

}
