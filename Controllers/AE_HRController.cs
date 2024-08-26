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
    public class AE_HRController : ControllerBase
    {
        /// <summary>
        /// 請假單-Flow_rest/Flow_rest_list.asp
        /// </summary>
        /// <param name="smbody"></param>
        /// <param name="dstaddr"></param>
        /// <returns></returns>
        [HttpPost("Flow_Rest_Query")]
        public ActionResult<ResultClass<string>> Flow_Rest_Query(string smbody, string dstaddr)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            try
            {
               
              
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
