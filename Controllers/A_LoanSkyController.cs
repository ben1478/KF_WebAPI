using Grpc.Net.Client;
using KF_WebAPI.BaseClass;
using KF_WebAPI.BaseClass.LoanSky;
using KF_WebAPI.DataLogic;
using KF_WebAPI.FunctionHandler;
using LoanSky.Model;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;

namespace KF_WebAPI.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class A_LoanSkyController : ControllerBase
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

                if (ReqClass.oreRequest != null)
                {
                    ResultClass<string> APIResult = new();
                    APIResult = await _Comm.Call_LoanSkyAPI(req.p_USER, "CreateOrderRealEstateAsync", ReqClass.oreRequest, req);
                    _AE_LoanSky.House_pre_Update(ReqClass); // 更新House_pre裡LoanSky.相關欄位
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


        [HttpPost("IsLoanSkyFieldsNull")]
        public ActionResult<ResultClass<string>> IsLoanSkyFieldsNull(LoanSky_Req req)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            AE_LoanSky _AE_LoanSky = new AE_LoanSky();
            // 取得待轉-基本資料
            try
            {
                runOrderRealEstateRequest ReqClass = _AE_LoanSky.IsLoanSkyFieldsNull(req);

                resultClass.ResultCode = "000";
                resultClass.ResultMsg = ReqClass.message; // 將LoanSky的錯誤訊息回傳;若message為空值代表問題
                resultClass.objResult = JsonConvert.SerializeObject(ReqClass);
                return Ok(resultClass);

            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass);
            }

        }

        /// <summary>
        /// 取得 縣市代碼
        /// </summary>
        /// <returns></returns>
        [HttpGet("GetAllCity")]
        public ActionResult<ResultClass<string>> GetAllCity()
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            AE_LoanSky ae_LoanSky = new AE_LoanSky();
            try
            {
                List<CityCode> cities = ae_LoanSky.AE2MoiCityCode(string.Empty);

                if (cities != null && cities.Count > 0)
                {
                    cities.Insert(0, new CityCode { city_num = "all", city_name = "全部" });

                    resultClass.ResultCode = "000";
                    resultClass.objResult = JsonConvert.SerializeObject(cities);
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
        }


        /// <summary>
        /// 取得「縣市」下所有「區」
        /// </summary>
        /// <returns></returns>
        [HttpGet("GetAllAreaByCity")]
        public ActionResult<ResultClass<string>> GetAllAreaByCity(string city_num)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            AE_LoanSky ae_LoanSky = new AE_LoanSky();
            try
            {
                List<AreaCode> areas = ae_LoanSky.AE2MoiTownCode(city_num, string.Empty);

                if (areas != null && areas.Count > 0)
                {
                    areas.Insert(0, new AreaCode { area_num = "all", area_name = "全部" });

                    resultClass.ResultCode = "000";
                    resultClass.objResult = JsonConvert.SerializeObject(areas);
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
        }

        /// <summary>
        /// 取得「縣市+區」下所有「段」
        /// </summary>
        /// <returns></returns>
        [HttpGet("GetAllRoadByAreaCity")]
        public ActionResult<ResultClass<string>> GetAllRoadByAreaCity(string city_num, string area_num)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            AE_LoanSky ae_LoanSky = new AE_LoanSky();
            try
            {
                List<SectionCode> road = ae_LoanSky.GetMoiSectionCode(city_num, area_num, string.Empty);

                if (road != null && road.Count > 0)
                {
                    road.Insert(0, new SectionCode { road_num = "all", road_name = "全部" });

                    resultClass.ResultCode = "000";
                    resultClass.objResult = JsonConvert.SerializeObject(road);
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
        }
        #endregion
    }
}
