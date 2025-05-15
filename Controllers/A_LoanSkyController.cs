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

namespace KF_WebAPI.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class A_LoadSkyController : ControllerBase
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
            ResultClass<string> resultClass = new ResultClass<string>();

            // 取得待轉-基本資料
            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"
                     select *
                         ,(select CS_PID from House_apply where HA_cknum = @HA_cknum) as CS_PID
                         ,(select  case when U_num in (SELECT item_D_code FROM dbo.Item_list where item_M_code ='newloan' and item_M_type='N') then 'BC0800' else U_BC  end U_BC FROM User_M where U_num = House_pre.add_num AND del_tag='0') as U_BC
                         ,(select U_name FROM User_M where U_num = House_pre.add_num AND del_tag='0') as add_name
                         ,(select U_num  FROM User_M where U_num = House_pre.add_num AND del_tag='0') as user_num
                         ,(select item_D_name from Item_list where item_M_code = 'building_kind' AND item_D_type='Y' AND item_D_code = House_pre.pre_building_kind  AND show_tag='0' AND del_tag='0') as show_pre_building_kind
                         ,(select item_D_name from Item_list where item_M_code = 'parking_kind' AND item_D_type='Y' AND item_D_code = House_pre.pre_parking_kind  AND show_tag='0' AND del_tag='0') as show_pre_parking_kind
                         ,(select item_D_name from Item_list where item_M_code = 'building_material' AND item_D_type='Y' AND item_D_code = House_pre.pre_building_material  AND show_tag='0' AND del_tag='0') as show_pre_building_material
                         ,isnull((select U_name FROM User_M where U_num = House_pre.pre_process_num AND del_tag='0'),'') as pre_process_name
                         from House_pre 
                         where del_tag = '0' AND HP_id = @HP_id  AND HP_cknum = @HP_cknum
                    ";
                parameters.Add(new SqlParameter("@HA_cknum", req.HA_cknum));
                parameters.Add(new SqlParameter("@HP_id", req.HP_id));
                parameters.Add(new SqlParameter("@HP_cknum", req.HP_cknum));

                #endregion

                // todo:基本資料轉值-參照對照表:建物類型/使用現況/車位型態/縣市代碼/鄉鎮市區代碼

                DataTable dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
                if (dtResult.Rows.Count > 0)
                {
                    // to 案件資料
                    var orderRealEstateRequest = dtResult.AsEnumerable().Select(row => new OrderRealEstateRequest
                    {
                        Account = account, // 承辦人員帳號:測試用帳號:testCGU
                        BusinessUserName = row.IsNull("add_name") ? "" : row.Field<string>("add_name"),  //經辦人名稱
                        BusinessTel = "82096100",  // 經辦人電話
                        BusinessFax = "123",   // 經辦人傳真
                        Applicant = row.IsNull("pre_apply_name") ? "" : row.Field<string>("pre_apply_name"),    // 申請人
                        IdNo = row.IsNull("CS_PID") ? "" : row.Field<string>("CS_PID"),         // 公司統編/身分證號
                        Condominium = row.IsNull("pre_community_name") ? "" : row.Field<string>("pre_community_name"),   // 社區(大樓)名稱
                        BuildingState = row.IsNull("show_pre_building_kind") ? "" : row.Field<string>("show_pre_building_kind"),  // 建物類型(請參照對照表)
                        Situation = row.IsNull("pre_use_kind") ? "" : row.Field<string>("pre_use_kind"),   //使用現況
                        ParkCategory = row.IsNull("show_pre_parking_kind") ? "" : row.Field<string>("show_pre_parking_kind"),  // 車位型態(請參照對照表)
                        Note = "123",  // 備註:建物類型備註pre_note+主要建築材料備註pre_building_material_note+委託人備註pre_principal_note+其他備註pre_other_note
                        MoiCityCode = row.IsNull("pre_city") ? "" : row.Field<string>("pre_city"), // 縣市代碼(請參照對照表)
                        MoiTownCode = row.IsNull("pre_area") ? "" : row.Field<string>("pre_area") // 鄉鎮市區代碼(請參照對照表)
                    }).SingleOrDefault();

                    // 申請標的:BuildNos或LandNos，必須填其中一個
                    var no = dtResult.AsEnumerable().Select(row => new OrderRealEstateNoRequest
                    {
                        MoiSectionCode = "1400",  // 段代碼:查詢條件：縣市代碼+區代碼+段名稱
                        BuildNos = row.IsNull("pre_build_num") ? "" : row.Field<string>("pre_build_num"),   // 建號 pre_build_num 多筆用逗號分隔
                        LandNos = row.IsNull("pre_land_num") ? "" : row.Field<string>("pre_land_num") // 地號 pre_land_num 多筆用逗號分隔
                    }).SingleOrDefault();
                    orderRealEstateRequest.Nos.Add(no);


                    // todo:取得待轉-PDF附件
                    var t_sqlAttachment = @"
                         select top 4 * 
                            from ASP_UpLoad
                            where del_tag = '0' AND cknum = @HA_cknum
                                and upload_name_code like '%.pdf'
                                order by add_date desc
                    ";
                    parameters.Add(new SqlParameter("@HA_cknum", req.HA_cknum));
                    DataTable dtResultAttachment = _adoData.ExecuteQuery(t_sqlAttachment, parameters);
                    if (dtResultAttachment.Rows.Count > 0)
                    {
                        foreach (var item in dtResultAttachment.AsEnumerable())
                        {
                            var attachment = new OrderRealEstateAttachmentRequest
                            {
                                OrginalFileName = item.IsNull("upload_name_show") ? "" : item.Field<string>("upload_name_show"),   // 原始檔案名稱(需包含副檔名且是PDF檔)
                                File = item.IsNull("upload_name_code") ? null : System.IO.File.ReadAllBytes($@"C:\Users\user\Downloads\{item.Field<string>("upload_name_code")}")
                            };
                            orderRealEstateRequest.Attachments.Add(attachment);
                        }
                    }
                    request.OrderRealEstates.Add(orderRealEstateRequest);

                    var company = channel.CreateGrpcService<ICompany>();
                    var reply = await company.CreateOrderRealEstateAsync(request);

                    resultClass.ResultCode = "000";
                    resultClass.objResult = JsonConvert.SerializeObject(reply);

                    // todo:寫入 api_error_log

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
