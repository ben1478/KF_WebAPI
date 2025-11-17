using KF_WebAPI.BaseClass;
using KF_WebAPI.BaseClass.WebRobot;
using Newtonsoft.Json;

namespace KF_WebAPI.Service
{
    public interface IProxyService
    {
        Task<ResultClass<List<string>>> GetProxyFlowAsync();
    }

    public class ProxyService : IProxyService
    {
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
