using LoanSky;
using LoanSky.Model;
using Grpc.Net.Client;
using ProtoBuf.Grpc.Client;
using KF_WebAPI.BaseClass;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using KF_WebAPI.BaseClass.LoadSky;

namespace KF_WebAPI.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class AE_LoadSkyController : ControllerBase
    {
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
            
            // 取得待轉資料


            // to 案件資料
            var orderRealEstateRequest = new OrderRealEstateRequest
            {
                Account = account, // 承辦人員帳號
                BusinessUserName = "台北",  //經辦人名稱
                BusinessTel = "82096100",  // 經辦人電話
                BusinessFax = "123",   // 經辦人傳真
                Applicant = "測試",    // 申請人
                IdNo = "123",         // 公司統編/身分證號
                Condominium = "123",   // 社區(大樓)名稱
                BuildingState = "公寓(5樓含以下無電梯)",  // 建物類型(請參照對照表)
                Situation = "123",   //使用現況
                ParkCategory = "坡道平面",  // 車位型態(請參照對照表)
                Note = "123",  // 備註
                MoiCityCode = "H", // 縣市代碼(請參照對照表)
                MoiTownCode = "07" // 鄉鎮市區代碼(請參照對照表)
            };

            // 附件，必須是PDF檔
            var attachment = new OrderRealEstateAttachmentRequest
            {
                OrginalFileName = "123.pdf",   // 原始檔案名稱(需包含副檔名且是PDF檔)
                File = System.IO.File.ReadAllBytes(@"C:\Users\user\Downloads\123.pdf")
            };

            orderRealEstateRequest.Attachments.Add(attachment);

            // 申請標的:BuildNos或LandNos，必須填其中一個
            var no = new OrderRealEstateNoRequest
            {
                MoiSectionCode = "1400",  // 段代碼
                BuildNos = "258",   // 建號多筆用逗號分隔
                LandNos = "158" // 地號多筆用逗號分隔
            };
            orderRealEstateRequest.Nos.Add(no);

            request.OrderRealEstates.Add(orderRealEstateRequest);

            ResultClass<string> resultClass = new ResultClass<string>();
            try
            {

                var company = channel.CreateGrpcService<ICompany>();
                var reply = await company.CreateOrderRealEstateAsync(request);

                resultClass.ResultCode = "000";
                resultClass.objResult = JsonConvert.SerializeObject(reply);
                return Ok(resultClass);
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                resultClass.ResultMsg = $"Response error: {ex.Message}";
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
