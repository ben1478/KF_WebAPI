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

namespace KF_WebAPI.Controllers
{
    [ApiController]
    public class A_WintonController : Controller
    {
        private readonly string urlBase = "http://192.168.1.244:49780/datasnap/";
        private readonly string AWintonCustID = "1112545";
        private readonly string ACompNo = "52611690";
        private readonly string AUserId = "API";
        private readonly string APassword = "52611690";
        private readonly string apiCode = "Winton";

        HttpClient _httpClient = new HttpClient();
        FuncHandler _fun;

        [HttpPost("eToken")]
        public async Task<IActionResult> GetLoginToken()
        {
            var apiName = "rest/TdmServerMethodsSYS/eLoginToken";
            var url = urlBase + apiName;

            var requestData = new
            {
                AWintonCustID = AWintonCustID,
                ACompNo = ACompNo,
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

                var responseObject = JsonConvert.DeserializeObject<ApiResponse>(responseContent);

                var token = responseObject.Data;

                return Ok(token);
            }
            catch (HttpRequestException)
            {
                _fun.ExtAPILogIns("Winton", apiName, "", "", jsonData, "500", "token失誤");
                return BadRequest();
            }
        }

        [HttpPost("SendSummons")]
        public async Task<IActionResult> Summons_Ins(string Form_ID)
        {
            var apiName = "rest/TdmServerMethodsTR/ImpWD4MFGL";
            var url = urlBase + apiName;

            var result = await GetLoginToken() as OkObjectResult;
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
                var parameters = new List<SqlParameter>();
                var T_SQL = @"select VP_type,VP_BC,VP_Pay_Type,VP_Total_Money,bank_code,bank_name,branch_name,bank_account,payee_name,VP_MFG_Date 
                    from InvPrepay_M where VP_ID=@VP_ID ";
                parameters.Add(new SqlParameter("@VP_ID", Form_ID));

                var parameters_d = new List<SqlParameter>();
                var T_SQL_D = @"select Form_ID,VD_Fee_Summary,VD_Fee,VD_Account_code,VD_Account from InvPrepay_D where VP_ID=@VP_ID ";
                parameters_d.Add(new SqlParameter("@VP_ID", Form_ID));
                #endregion
                var dtResult = _adoData.ExecuteQuery(T_SQL, parameters);

                var vpMFGDate = dtResult.Rows[0]["VP_MFG_Date"];
                var parameters_mf = new List<SqlParameter>();
                var T_SQL_MF = @"select * from InvPrepay_M where LEFT(VP_MFG_Date,7)=@yyyMMdd";
                parameters_mf.Add(new SqlParameter("@yyyMMdd", vpMFGDate));
                var no = _adoData.ExecuteQuery(T_SQL_MF, parameters_mf).AsEnumerable().Count();
                vpMFGDate = vpMFGDate + no.ToString("D4");

                var m_modelist_m = new Summons_M_req();
                m_modelist_m.MFGL003 = "A" + vpMFGDate.ToString();
                m_modelist_m.MFGL004 = "3";
                m_modelist_m.MFGL005 = DateTime.Now.ToString("yyyy-MM-dd");
                m_modelist_m.MFGL006 = Form_ID;
                m_modelist_m.MFGL009 = "N";
                m_Summons_req.ADataSetMaster = m_modelist_m;

                int indexnumber = 0;
                //TODO 供應商代號
                var modelist_d = _adoData.ExecuteQuery(T_SQL_D, parameters_d).AsEnumerable().Select(row => new Summons_D_req 
                {
                    DTGL004 = m_Summons_req.ADataSetMaster.MFGL003,
                    DTGL005 = (indexnumber++).ToString("D4"),
                    DTGL008 = row.Field<string>("VD_Account_code"),
                    DTGL009 = "C0001",
                    DTGL011 = row.Field<string>("VD_Fee_Summary"),
                    DTGL012 = "1",
                    DTGL013 = row.Field<int>("VD_Fee"),
                    DTGL021 = row.Field<int>("VD_Fee"),
                    DTGL028 = 1
                }).ToList();


                //TODO 反寫一筆貸的明細 DTGL011 要開欄位填
                modelist_d.Add(new Summons_D_req {
                    DTGL004 = m_Summons_req.ADataSetMaster.MFGL003,
                    DTGL005 = (indexnumber++).ToString("D4"),
                    DTGL008 = "2145",
                    DTGL009 = "",
                    DTGL011 = modelist_d.First().DTGL011,
                    DTGL012 = "2",
                    DTGL013 = modelist_d.Sum(x=>x.DTGL013),
                    DTGL021 = modelist_d.Sum(x => x.DTGL013),
                    DTGL028 = 1
                });

                m_Summons_req.ADataSetDetail = modelist_d;

                jsonData = JsonConvert.SerializeObject(m_Summons_req);
                var content = new StringContent(jsonData, Encoding.UTF8, "application/json");

                var response = await _httpClient.PostAsync(url, content);
                response.EnsureSuccessStatusCode();

                var responseContent = await response.Content.ReadAsStringAsync();
                var responseObject = JsonConvert.DeserializeObject<ApiResponse<Dictionary<string, object>>>(responseContent);

                if (responseObject.Status == "200" && responseObject.Error == null)
                {
                    _fun.ExtAPILogIns("Winton", apiName, Form_ID, token, jsonData, "200", responseObject.ToString());
                    return Ok();
                }
                else
                {
                    _fun.ExtAPILogIns("Winton", apiName, Form_ID, token, jsonData, "500", responseObject.ToString());
                    return BadRequest();
                }
            }
            catch (Exception ex)
            {
                _fun.ExtAPILogIns("Winton", apiName, Form_ID, token, jsonData, "500", $" error: {ex.Message}");
                return BadRequest();
            }
        }
    }
}
