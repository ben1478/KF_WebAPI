using System.Net.Http;
using System.Text;
using KF_WebAPI.BaseClass;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;

namespace KF_WebAPI.FunctionHandle
{
    public class ExternalAPI
    {
      

        public HttpRequestMessage CallAPI(string Name, string p_strJson)
        {
            var uri = new Uri("https://egateway.tac.com.tw/uat/api/yrc/agent/" + Name);
            var request = new HttpRequestMessage(HttpMethod.Post, uri);

            var content = new StringContent(p_strJson, Encoding.UTF8, "application/json");
            request.Content = content;
            return request;
        }


        public BaseResult CheckStatusCode(HttpResponseMessage p_response)
        {
            BaseResult? m_BaseResult = new()
            {
                code = "000",
                msg = ""
            };

            if (!p_response.IsSuccessStatusCode)
            {
                m_BaseResult.code = "999";
                // Handle error response
                if (p_response.StatusCode == System.Net.HttpStatusCode.NotFound)
                {
                    m_BaseResult.msg = "API endpoint not found";
                }
                else if (p_response.StatusCode == System.Net.HttpStatusCode.Unauthorized)
                {
                    m_BaseResult.msg = "API endpoint requires authentication";
                }
                else
                {
                    m_BaseResult.msg = $"API call failed with status code {p_response.StatusCode}";
                }
            }
            return m_BaseResult;
        }
    }
}
