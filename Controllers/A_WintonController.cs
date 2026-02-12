using KF_WebAPI.BaseClass.Max104;
using KF_WebAPI.BaseClass;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using System.Net.Http;
using KF_WebAPI.DataLogic;
using System.Text;
using KF_WebAPI.BaseClass.Winton;
using KF_WebAPI.FunctionHandler;
using System;
using Microsoft.Data.SqlClient;
using KF_WebAPI.BaseClass.AE;
using System.Data;
using System.Reflection;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using Newtonsoft.Json.Linq;
using System.Text.RegularExpressions;

namespace KF_WebAPI.Controllers
{
    [ApiController]
    public class A_WintonController : Controller
    {
        private readonly string urlBase = "http://192.168.1.244:49780/datasnap/";
        private readonly string AWintonCustID = "1112545";
        private string ACompNo = "52611690";
        private readonly string AUserId = "API";
        private readonly string APassword = "52611690";
        private readonly string apiCode = "Winton";

        private static readonly HttpClient _httpClient = new HttpClient();

        FuncHandler _fun = new FuncHandler();

        [HttpPost("eToken")]
        public async Task<IActionResult> GetLoginToken(string vpCom)
        {
            var apiName = "rest/TdmServerMethodsSYS/eLoginToken";
            var url = urlBase + apiName;

            var requestData = new
            {
                AWintonCustID = AWintonCustID,
                ACompNo = vpCom,
                AUserId = AUserId,
                APassword = APassword,
                AAB = "A",
                AYear = (DateTime.Now.Year - 1911).ToString()
            };

            var jsonData = JsonConvert.SerializeObject(requestData);
            var content = new StringContent(jsonData, Encoding.UTF8, "application/json");

            try
            {
                var response = await _httpClient.PostAsync(url, content);
                response.EnsureSuccessStatusCode();

                var responseContent = await response.Content.ReadAsStringAsync();

                var responseObject = JsonConvert.DeserializeObject<ApietokenResponse>(responseContent);

                var token = responseObject.Data;

                return Ok(token);
            }
            catch (HttpRequestException ex)
            {
                _fun.ExtAPILogIns(apiCode, apiName, "", "", jsonData, "500", $" error: {ex.Message}");
                return BadRequest();
            }
        }

        [HttpPost("SendSummons")]
        public async Task<IActionResult> SendSummons(string vpCom, string Form_ID)
        {
            var apiName = "rest/TdmServerMethodsTR/ImpWD4MFGL";
            var url = urlBase + apiName;

            var result = await GetLoginToken(vpCom) as OkObjectResult;
            var token = result.Value.ToString();

            var jsonData = "";
            try
            {
                var m_Summons_req = new Summons_req();
                m_Summons_req.AToken = token;
                m_Summons_req.AUpdateType = "0";
                m_Summons_req.ACreateCB01 = "1";

                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"select MA.MF_ID,VP.*,VD.* from InvPrepay_M VP inner join InvPrepay_D VD on VP.VP_ID=VD.VP_ID
                    left join Manufacturer MA on MA.Company_name = VP.payee_name
                    where VP.VP_ID =@VP_ID";
                var parameters = new List<SqlParameter>
                {
                    new SqlParameter("@VP_ID", Form_ID)
                };
                #endregion
                var dtResult = _adoData.ExecuteQuery(T_SQL, parameters);

                var vpMFGDate = dtResult.Rows[0]["VP_MFG_Date"];
                var MF_ID = dtResult.Rows[0]["MF_ID"].ToString();

                var T_SQL_MF = @"select * from InvPrepay_M where LEFT(VP_MFG_Date,7)=@yyyMMdd";
                var parameters_mf = new List<SqlParameter> 
                {
                    new SqlParameter("@yyyMMdd", vpMFGDate)
                };
                var no = _adoData.ExecuteQuery(T_SQL_MF, parameters_mf).AsEnumerable().Count();
                var vpMFGID = vpMFGDate + no.ToString("D4");

               
                var T_SQL_UPD = @"Update InvPrepay_M set VP_MFG_ID=@VP_MFG_ID where VP_ID=@VP_ID";
                var parameters_upd = new List<SqlParameter> 
                {
                    new SqlParameter("@VP_MFG_ID", vpMFGID),
                    new SqlParameter("@VP_ID", Form_ID)
                };
                _adoData.ExecuteQuery(T_SQL_UPD, parameters_upd);

                m_Summons_req.ADataSetMaster = new Summons_M_req 
                {
                    MFGL003 = "A" + vpMFGID,
                    MFGL004 = "3",
                    MFGL005 = DateTime.Now.ToString("yyyy-MM-dd"),
                    MFGL006 = dtResult.Rows[0]["VP_Summary"].ToString(),
                    MFGL009 = "N"
                };

                int indexnumber = 1;
                m_Summons_req.ADataSetDetail = dtResult.AsEnumerable().Select(row => new Summons_D_req {
                    DTGL004 = m_Summons_req.ADataSetMaster.MFGL003,
                    DTGL005 = (indexnumber++).ToString("D4"),
                    DTGL008 = row.Field<string>("VD_Account_code"),
                    // 2025.12.3 先取消比對廠商 不丟值 DTGL009 = MF_ID,
                    DTGL011 = row.Field<string>("VD_Fee_Summary"),
                    DTGL012 = "1",
                    DTGL013 = row.Field<string>("VD_VAT") == "Y" ? (int)Math.Round(row.Field<int>("VD_Fee") / 1.05) : row.Field<int>("VD_Fee"),
                    DTGL014 = row.Field<string>("VD_BC"),
                    DTGL021 = row.Field<string>("VD_VAT") == "Y" ? (int)Math.Round(row.Field<int>("VD_Fee") / 1.05) : row.Field<int>("VD_Fee"),
                    DTGL028 = 1
                }).ToList();

                m_Summons_req.ADataSetDetail.Add(new Summons_D_req {
                    DTGL004 = m_Summons_req.ADataSetMaster.MFGL003,
                    DTGL005 = (indexnumber++).ToString("D4"),
                    DTGL008 = "2145",
                    // 2025.12.3 先取消比對廠商 不丟值 DTGL009 = MF_ID,
                    DTGL011 = dtResult.Rows[0]["VP_Nsummary"].ToString(),
                    DTGL012 = "2",
                    DTGL013 = m_Summons_req.ADataSetDetail.Sum(x => x.DTGL013),
                    DTGL021 = m_Summons_req.ADataSetDetail.Sum(x => x.DTGL013),
                    DTGL028 = 1
                });

                jsonData = JsonConvert.SerializeObject(m_Summons_req);
                var content = new StringContent(jsonData, Encoding.UTF8, "application/json");

                var response = await _httpClient.PostAsync(url, content);
                response.EnsureSuccessStatusCode();

                var responseContent = await response.Content.ReadAsStringAsync();
                var responseObject = JsonConvert.DeserializeObject<ApiResponse>(responseContent);


                if (responseObject.Status == "200" && responseObject.Error == null && responseObject.Data.success==1)
                {
                    _fun.ExtAPILogIns(apiCode, apiName, Form_ID, token, jsonData, "200", JsonConvert.SerializeObject(responseObject));
                    return Ok();
                }
                else
                {
                    _fun.ExtAPILogIns(apiCode, apiName, Form_ID, token, jsonData, "500", JsonConvert.SerializeObject(responseObject));
                    return BadRequest();
                }
            }
            catch (Exception ex)
            {
                _fun.ExtAPILogIns(apiCode, apiName, Form_ID, token, jsonData, "500", $" error: {ex.Message}");
                return BadRequest();
            }
        }

        /// <summary>
        /// strType:來源 FormID:PKey
        /// </summary>
        [HttpPost("SendManufacturer")]
        public async Task<IActionResult> SendManufacturer(string strType, string FormID,string User)
        {
            var apiName = "rest/TdmServerMethodsOT/ImpWD2SU01";
            var url = urlBase + apiName;

            var jsonData = "";
            Manufacturer_req model = await Manufacturer_Map(strType, FormID);
            try
            {
                jsonData = JsonConvert.SerializeObject(model);
                var content = new StringContent(jsonData, Encoding.UTF8, "application/json");

                var response = await _httpClient.PostAsync(url, content);
                response.EnsureSuccessStatusCode();

                var responseContent = await response.Content.ReadAsStringAsync();
                var responseObject = JsonConvert.DeserializeObject<ApiResponse>(responseContent);


                if (responseObject.Status == "200" && responseObject.Error == null)
                {
                    _fun.ExtAPILogIns(apiCode, apiName, model.ADataSet.SU01001, model.AToken, jsonData, "200", JsonConvert.SerializeObject(responseObject), User);
                    return Ok();
                }
                else
                {
                    _fun.ExtAPILogIns(apiCode, apiName, model.ADataSet.SU01001, model.AToken, jsonData, "500", JsonConvert.SerializeObject(responseObject), User);
                    return BadRequest();
                }
            }
            catch (Exception ex)
            {
                _fun.ExtAPILogIns(apiCode, apiName, model.ADataSet.SU01001, model.AToken, jsonData, "500", $" error: {ex.Message}");
                return BadRequest();
            }
        }

        private async Task<Manufacturer_req> Manufacturer_Map(string strType, string FormID)
        {
            var result = await GetLoginToken(ACompNo);

            if (result is OkObjectResult okResult)
            {
                var token = okResult.Value.ToString();

                Manufacturer_req model = new Manufacturer_req();
                model.AToken = token;
                model.AUpdateType = "0";
                model.ADataSet = new Manufacturer_M_req();
                ADOData _adoData = new ADOData();
                //來源為廠商資料
                if (strType == "1")
                {
                    #region SQL
                    var T_SQL = @"select * from Manufacturer where MF_ID=@MF_ID";
                    var parameters = new List<SqlParameter>
                    {
                        new SqlParameter("@MF_ID", FormID)
                    };
                    #endregion
                    var dtResult = _adoData.ExecuteQuery(T_SQL, parameters);

                    model.ADataSet.SU01001 = dtResult.Rows[0]["MF_ID"].ToString();
                    model.ADataSet.SU01002 = strType;
                    model.ADataSet.SU01082 = dtResult.Rows[0]["MF_ID"].ToString();
                    model.ADataSet.SU01004 = dtResult.Rows[0]["Company_name"].ToString();
                    model.ADataSet.SU01003 = dtResult.Rows[0]["Company_name"].ToString();
                    model.ADataSet.SU01007 = dtResult.Rows[0]["Company_number"].ToString();
                    model.ADataSet.SU01011 = dtResult.Rows[0]["Company_addr"] != DBNull.Value ? dtResult.Rows[0]["Company_addr"].ToString() : null;
                    model.ADataSet.SU01010 = dtResult.Rows[0]["Company_busin"] != DBNull.Value ? dtResult.Rows[0]["Company_busin"].ToString() : null;
                    model.ADataSet.SU01014 = dtResult.Rows[0]["Company_tel"] != DBNull.Value ? dtResult.Rows[0]["Company_tel"].ToString() : null;
                    model.ADataSet.SU01016 = dtResult.Rows[0]["Company_fax"] != DBNull.Value ? dtResult.Rows[0]["Company_fax"].ToString() : null;
                }

                //來源為客戶資料
                if(strType == "2")
                {
                    #region SQL
                    var T_SQL = @"select Item_list.item_D_code as U_BC_Show,* from House_apply 
                                  left join User_M on House_apply.plan_num = User_M.U_num
                                  left join User_Spec_Group on House_apply.plan_num = User_Spec_Group.U_num
                                  left join Item_list on item_M_code='winton_company' and item_D_txt_A = ISNULL(User_Spec_Group.Spec_Group,User_M.U_BC)
                                  where HA_id=@HA_id";
                    var parameters = new List<SqlParameter>
                    {
                        new SqlParameter("@HA_id", FormID)
                    };
                    #endregion
                    var dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
                    model.ADataSet.SU01001 = dtResult.Rows[0]["CS_PID"].ToString();
                    model.ADataSet.SU01002 = strType;
                    model.ADataSet.SU01003 = dtResult.Rows[0]["CS_name"].ToString();
                    model.ADataSet.SU01004 = dtResult.Rows[0]["CS_name"].ToString();

                    model.ADataSet.SU01011 = dtResult.Rows[0]["CS_register_address"].ToString();
                    model.ADataSet.SU01019 = dtResult.Rows[0]["CS_MTEL1"].ToString();
                    model.ADataSet.SU01029 = dtResult.Rows[0]["U_BC_Show"].ToString();
                    model.ADataSet.SU01038 = "2";

                    model.ADataSet.SU01076 = 1;
                    model.ADataSet.SU01082 = dtResult.Rows[0]["CS_PID"].ToString();
                    model.ADataSet.SU01096 = dtResult.Rows[0]["CS_EMAIL"].ToString();
                    model.ADataSet.SU01107 = "B";
                    model.ADataSet.SU01110 = "1";
                    model.ADataSet.SU01112 = "Y";

                    //存入載具
                    if (dtResult.Rows[0]["IsVehicle"].ToString() == "Y")
                    {
                        model.ADataSet.SU01123 = 1;
                        model.ADataSet.SU01124 = "EJ0110";
                        model.ADataSet.SU01125 = dtResult.Rows[0]["Vehicle"].ToString();
                    }
                }
                return model;
            }
            else
            {
                _fun.ExtAPILogIns(apiCode, "eToken", FormID, "", "", "500", $" error: Failed to get token");
                throw new Exception("Failed to get token.");
            }
        }

    }
}
