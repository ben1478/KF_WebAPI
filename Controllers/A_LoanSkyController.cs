using Grpc.Net.Client;
using KF_WebAPI.BaseClass;
using KF_WebAPI.BaseClass.AE;
using KF_WebAPI.BaseClass.LoadSky;
using KF_WebAPI.FunctionHandler;
using LoanSky;
using LoanSky.Model;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using ProtoBuf.Grpc.Client;
using System.Data;
using Microsoft.Data.SqlClient;
using KF_WebAPI.BaseClass.Max104;
using System.Net.Mail;
using System.Net.Http;
using System.Transactions;
using KF_WebAPI.DataLogic;
using System.Collections.Generic;

namespace KF_WebAPI.Controllers
{
    [ApiController]
    public class A_LoadSkyController : ControllerBase
    {
        /// <summary>
        /// 是否呼叫測試API,true:用資料庫模擬API,false:呼叫裕富API;
        /// </summary>
        private readonly HttpClient _httpClient;
        private Common _Comm = new();

        private string account = "testCGU"; // 測試用帳號:testCGU
        private GrpcChannel channel = GrpcChannel.ForAddress("https://land.loansky.net:5055");
        private OrderRealEstateAdapterRequest request = new OrderRealEstateAdapterRequest
        {
            ApiKey = "1AA86E9C37F24767ACA1087DEBA6D322"     //測試用帳號 ApiKey:1AA86E9C37F24767ACA1087DEBA6D322
        };
        #region LoanSky 宏邦API
        /// <summary>
        /// <summary>
        /// Grpc多筆匯入 CreateOrderRealEstatesAsync
        /// </summary>
        [HttpPost("CreateOrderRealEstatesAsync")]
        public async Task<ActionResult<ResultClass<string>>> CreateOrderRealEstatesAsync(LoanSky_Req req)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            AE_LoanSky _AE_LoanSky = new AE_LoanSky();
            // 取得待轉-基本資料
            try
            {
                runOrderRealEstateRequest ReqClass = _AE_LoanSky.GetOrderRealEstateRequest(req);

                // todo:基本資料轉值-參照對照表:建物類型/使用現況/車位型態/縣市代碼/鄉鎮市區代碼
                if (ReqClass.oreRequest != null)
                {
                    // 取得AE要拋轉LoanSky案件資料-pdf附件
                    List<OrderRealEstateAttachmentRequest> attachments = _AE_LoanSky.GetOrderRealEstateNoRequest(req.HA_cknum);
                    if (attachments != null)
                    {
                        foreach (var item in attachments)
                        {
                            ReqClass.oreRequest.Attachments.Add(item);
                        }
                    }
                    

                    ResultClass<string> APIResult = new();
                    APIResult = await _Comm.Call_LoanSkyAPI(req.p_USER, "CreateOrderRealEstateAsync", ReqClass.oreRequest, req);

                    return Ok(resultClass);
                }
                else
                {
                    resultClass.ResultCode = "400";
                    resultClass.ResultMsg = "查無資料";
                    return BadRequest(resultClass);
                }
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass);
            }
            finally
            {
                await channel.ShutdownAsync();
            }

        }

        #endregion
    }
}
