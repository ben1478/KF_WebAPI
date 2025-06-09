using KF_WebAPI.BaseClass;
using KF_WebAPI.BaseClass.AE;
using KF_WebAPI.BaseClass.LoanSky;
using KF_WebAPI.DataLogic;
using KF_WebAPI.FunctionHandler;
using LoanSky.Model;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;

namespace KF_WebAPI.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class A_LoanSkyController : ControllerBase
    {
        string apiCode = "LoanSky";
        private Common _Comm = new();
        FuncHandler _fun = new();
        AE_LoanSky _AE_LoanSky = new();
        ResultClass<string> resultClass = new();

        #region LoanSky 宏邦API
        [HttpPost("IsLoanSkyFieldsNull")]
        public async Task<ActionResult<ResultClass<string>>> IsLoanSkyFieldsNull(LoanSky_Req req)
        {
            req.LoanSkyApiKey = string.IsNullOrEmpty(req.LoanSkyApiKey)? "7677949CB3B34E6CA75B8D81E2975478" : req.LoanSkyApiKey; //正式區："7677949CB3B34E6CA75B8D81E2975478", 測試區:"1AA86E9C37F24767ACA1087DEBA6D322"
            req.LoanSkyAccount = string.Empty;// "testCGU"; // 測試用帳號:testCGU; 正式用帳號:LoanSkyAccount=string.empty
            req.LoanSkyUrl = string.IsNullOrEmpty(req.LoanSkyUrl) ? "https://land.loansky.net:5075" : req.LoanSkyUrl; //正式區： https://land.loansky.net:5075; 測試區:https://land.loansky.net:5055
            string apiName = "CreateOrderRealEstateAsync";
            string apikey = JsonConvert.SerializeObject(req);
            
            
            runOrderRealEstateRequest ReqClass = await _AE_LoanSky.IsLoanSkyFieldsNull(req);
            // 若有錯誤訊息則返回前端
            if(string.IsNullOrEmpty(ReqClass.message)==false)
            {
                resultClass.ResultCode = "500";
                resultClass.ResultMsg = ReqClass.message; // 將LoanSky的錯誤訊息回傳;若message為空值代表問題
                resultClass.objResult = JsonConvert.SerializeObject(ReqClass);
                return BadRequest(resultClass);
            }
            
            try
            {
                resultClass = await _Comm.Call_LoanSkyAPI(req.p_USER, apiName, ReqClass.oreRequest, req);
                if (resultClass.ResultCode == "200")
                {
                    #region 更新House_pre裡LoanSky.相關欄位
                    ReqClass.housePre_res.MoiCityCode = ReqClass.oreRequest.MoiCityCode;
                    ReqClass.housePre_res.MoiTownCode = ReqClass.oreRequest.MoiTownCode;
                    ReqClass.housePre_res.MoiSectionCode = ReqClass.oreRequest.Nos.FirstOrDefault().MoiSectionCode;
                    ReqClass.housePre_res.BuildingState = ReqClass.oreRequest.BuildingState;
                    ReqClass.housePre_res.ParkCategory = ReqClass.oreRequest.ParkCategory;
                    ReqClass.housePre_res.edit_num = req.p_USER;

                    _AE_LoanSky.House_pre_Update(ReqClass.housePre_res); // 更新House_pre裡LoanSky.相關欄位
                    #endregion
                    resultClass.ResultMsg = "轉宏邦(已傳)";
                    return Ok(resultClass);
                }
                else
                {
                    resultClass.ResultCode = "500";
                    resultClass.ResultMsg = resultClass.ResultMsg;
                    return BadRequest(resultClass);
                }
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                resultClass.ResultMsg = ex.Message;
                _fun.ExtAPILogIns(apiCode, apiName, apikey, req.LoanSkyApiKey,
                    "", "500", $" response: {ex.Message}");
                return BadRequest(resultClass);
            }
        }

        [HttpPost("House_pre_Update")]
        public ActionResult<ResultClass<string>> House_pre_Update(HousePre_res housePre)
        {

            resultClass = _AE_LoanSky.House_pre_Update(housePre);
            return Ok(resultClass);
        }

        /// <summary>
        /// 取得 縣市代碼
        /// </summary>
        /// <returns></returns>
        [HttpGet("GetAllCity")]
        public ActionResult<ResultClass<string>> GetAllCity()
        {
            try
            {
                List<CityCode> cities = _AE_LoanSky.AE2MoiCityCode(string.Empty);

                if (cities != null && cities.Count > 0)
                {
                    cities.Insert(0, new CityCode { city_num = "all", city_name = "全部" });

                    resultClass.ResultCode = "000";
                    resultClass.objResult = JsonConvert.SerializeObject(cities);
                    return Ok(resultClass);
                }
                else
                {
                    resultClass.ResultCode = "500";
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
            try
            {
                List<AreaCode> areas = _AE_LoanSky.AE2MoiTownCode(city_num, string.Empty);

                if (areas != null && areas.Count > 0)
                {
                    areas.Insert(0, new AreaCode { area_num = "all", area_name = "全部" });

                    resultClass.ResultCode = "000";
                    resultClass.objResult = JsonConvert.SerializeObject(areas);
                    return Ok(resultClass);
                }
                else
                {
                    resultClass.ResultCode = "500";
                    resultClass.ResultMsg = "查無資料";
                    return BadRequest(resultClass);
                }
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                resultClass.ResultMsg = $" response: {ex.Message}";
                return BadRequest(resultClass);
            }
        }

        /// <summary>
        /// 取得「縣市+區」下所有「段」
        /// </summary>
        /// <returns></returns>
        [HttpGet("GetAllRoadByAreaCity")]
        public ActionResult<ResultClass<string>> GetAllRoadByAreaCity(string city_num, string area_num)
        {
            try
            {
                List<SectionCode> road = _AE_LoanSky.GetMoiSectionCode(city_num, area_num, string.Empty);

                if (road != null && road.Count > 0)
                {
                    road.Insert(0, new SectionCode { road_num = "all", road_name = "全部" });

                    resultClass.ResultCode = "000";
                    resultClass.objResult = JsonConvert.SerializeObject(road);
                    return Ok(resultClass);
                }
                else
                {
                    resultClass.ResultCode = "500";
                    resultClass.ResultMsg = area_num.Equals("all")? "區/市/鄉/鎮:資料有錯請檢查" :
                        $"縣市代碼:{city_num} 或區/市/鄉/鎮代碼:{area_num}在對照表查無資料";
                    return BadRequest(resultClass);
                }
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                resultClass.ResultMsg = $" response: {ex.Message}";
                return BadRequest(resultClass);
            }
        }

        [HttpGet("GetPre_building_kind")]
        public ActionResult<ResultClass<string>> GetPre_building_kind()
        {
            try
            {
                var res = _AE_LoanSky.GetPre_building_kind();

                if (res != null)
                {
                    resultClass.ResultCode = "000";
                    resultClass.objResult = JsonConvert.SerializeObject(res);
                    return Ok(resultClass);
                }
                else
                {
                    resultClass.ResultCode = "500";
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
        [HttpGet("GetBuildingState")]
        public ActionResult<ResultClass<string>> GetBuildingState()
        {
            try
            {
                var res = _AE_LoanSky.GetBuildingState();

                if (res != null)
                {
                    resultClass.ResultCode = "000";
                    resultClass.objResult = JsonConvert.SerializeObject(res);
                    return Ok(resultClass);
                }
                else
                {
                    resultClass.ResultCode = "500";
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
