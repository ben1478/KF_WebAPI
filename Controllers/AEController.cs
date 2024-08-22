using KF_WebAPI.BaseClass;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.IdentityModel.Tokens;
using System;
using System.Data.SqlClient;
using System.IdentityModel.Tokens.Jwt;
using System.Net;
using System.Security.Claims;
using System.Text;

namespace KF_WebAPI.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class AEController : ControllerBase
    {

        [HttpPost("SendSMS")]
        public ActionResult<ResultClass<string>> SendSMS(string smbody, string dstaddr)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            try
            {
                StringBuilder reqUrl = new StringBuilder();
                reqUrl.Append("http://smsapi.mitake.com.tw/api/mtk/SmSend?CharsetURL=UTF-8");
                StringBuilder m_params = new StringBuilder();
                m_params.Append("username=52611690SMS");
                m_params.Append("&password=QST5n6AJzb");
                m_params.Append("&dstaddr=" + dstaddr);
                m_params.Append("&smbody=" + smbody);
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(new Uri(reqUrl.ToString()));
                request.Method = "POST";
                request.ContentType = "application/x-www-form-urlencoded";
                byte[] bs = Encoding.UTF8.GetBytes(m_params.ToString());
                request.ContentLength = bs.Length;
                request.GetRequestStream().Write(bs, 0, bs.Length);
             
                using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                {
                    using (StreamReader sr = new StreamReader(response.GetResponseStream()))
                    {
                        resultClass.objResult = sr.ReadToEnd();
                    }
                }
                resultClass.ResultCode = "000";
               
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "999";
                resultClass.ResultMsg = $" response: {ex.Message}";
            }
            return Ok(resultClass);
        }



    }

}
