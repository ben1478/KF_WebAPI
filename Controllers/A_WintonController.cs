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
using static OfficeOpenXml.ExcelErrorValue;
using System.Collections.Generic;
using Grpc.Core;
using System.Linq;

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
                    _fun.ExtAPILogIns(apiCode, "ImpWD2SU01", model.ADataSet.SU01001, model.AToken, jsonData, "200", JsonConvert.SerializeObject(responseObject), User);
                    return Ok();
                }
                else
                {
                    _fun.ExtAPILogIns(apiCode, "ImpWD2SU01", model.ADataSet.SU01001, model.AToken, jsonData, "500", JsonConvert.SerializeObject(responseObject), User);
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

                    model.ADataSet.SU01010 = dtResult.Rows[0]["CS_register_address"].ToString();
                    model.ADataSet.SU01011 = dtResult.Rows[0]["CS_register_address"].ToString();
                    model.ADataSet.SU01012 = dtResult.Rows[0]["CS_register_address"].ToString();
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


        /// <summary>
        /// 清償轉文中開發票to銷帳+
        /// </summary>
        /// <param name="List"></param>
        /// <returns></returns>
        [HttpPost("SendPayOffForInv")]
        public async Task<ResultClass<string>> SendPayOffForInv([FromBody] List<PayOff_Win_Inv> List)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

            try
            {
                var result = await GetLoginToken(ACompNo);
                AE_Rpt _Rpt = new AE_Rpt();
                if (result is OkObjectResult okResult)
                {
                    var apiName = "rest/TdmServerMethodsIN/ImpWD4MF10";
                    var url = urlBase + apiName;

                    for (int i = 0; i < List.Count; i++)
                    {
                        #region 抓取可用的發票組別
                        string yyyMM = (DateTime.Now.Year - 1911).ToString() + DateTime.Now.Month.ToString("D2");
                        var GroupNo = _Rpt.CheckInvGp(yyyMM, "01");
                        #endregion

                        #region maping
                        ReceivableForInv_req model = new ReceivableForInv_req();
                        model.AToken = okResult.Value.ToString();
                        model.ADocType = "20";
                        model.AUpdateType = 0;
                        model.InvoiceGroup = GroupNo;

                        model.ADataSetMaster = new List<ReceivableForInv_M_req>();
                        model.ADataSetDetail = new List<ReceivableForInv_D_req>();

                        ReceivableForInv_M_req model_M = new ReceivableForInv_M_req();
                        model_M.MF10003 = "P" + List[i].HS_id + "-" + List[i].RC_count.ToString("D3");
                        model_M.MF10004 = DateTime.Now.ToString("yyyy-MM-dd");
                        model_M.MF10008 = List[i].CS_PID;
                        model_M.MF10010 = "00";
                        model_M.MF10011 = "00";
                        model_M.MF10012 = "";
                        model_M.MF10018 = "2";
                        model_M.MF10022 = "20";
                        model_M.MF10059 = "1";
                        model_M.MF10066 = DateTime.Now.ToString("yyyy-MM-dd");
                        model_M.MF10091 = "N";
                        model_M.MF10093 = "Y";
                        model_M.MF10094 = "Y";
                        model.ADataSetMaster.Add(model_M);

                       
                        if (List[i].Delay_AMT > 0)//延滯總利息-005
                        {
                            ReceivableForInv_D_req model_D1 = new ReceivableForInv_D_req();
                            model_D1.DT10004 = model_M.MF10003;
                            model_D1.DT10006 = "005";
                            model_D1.DT10030 = 1;
                            model_D1.DT10040 = List[i].Delay_AMT;
                            model.ADataSetDetail.Add(model_D1);
                        }

                        #region 如果是汽機車會有手續費(未沖銷筆數*20)
                        string[] strPjt = new string[] { "PJ00046", "PJ00047", "PJ00048", "PJ00099" };
                        if (strPjt.Contains(List[i].project_title))
                        {
                            ReceivableForInv_D_req model_D1 = new ReceivableForInv_D_req();
                            model_D1.DT10004 = model_M.MF10003;
                            model_D1.DT10006 = "003";
                            model_D1.DT10030 = 1;
                            model_D1.DT10040 = (List[i].month_total - List[i].RC_count + 1 )*20;
                            model.ADataSetDetail.Add(model_D1);
                            List[i].Interest_AMT = List[i].Interest_AMT - model_D1.DT10040;
                        }
                        #endregion

                        if (List[i].Interest_AMT > 0)//結清利息-014
                        {
                            ReceivableForInv_D_req model_D1 = new ReceivableForInv_D_req();
                            model_D1.DT10004 = model_M.MF10003;
                            model_D1.DT10006 = "014";
                            model_D1.DT10030 = 1;
                            model_D1.DT10040 = List[i].Interest_AMT;
                            model.ADataSetDetail.Add(model_D1);
                        }

                        if (List[i].Break_AMT > 0)//違約金/作業費
                        {
                            ReceivableForInv_D_req model_D1 = new ReceivableForInv_D_req();
                            model_D1.DT10004 = model_M.MF10003;
                            if (List[i].Break_Type =="A")//A:違約金
                            {
                                model_D1.DT10006 = "004";
                            }
                            else///B:作業費
                            {
                                model_D1.DT10006 = "006";
                            }
                            model_D1.DT10030 = 1;
                            model_D1.DT10040 = List[i].Break_AMT;
                            model.ADataSetDetail.Add(model_D1);
                        }

                        #endregion

                        var jsonData = "";
                        jsonData = JsonConvert.SerializeObject(model);
                        var content = new StringContent(jsonData, Encoding.UTF8, "application/json");

                        var response = await _httpClient.PostAsync(url, content);
                        response.EnsureSuccessStatusCode();

                        var responseContent = await response.Content.ReadAsStringAsync();
                        var responJson = JObject.Parse(responseContent);

                        string errmsg = (string)responJson["data"]["result"][0]["errmsg"];
                        string Status = (string)responJson["status"];

                        _fun.ExtAPILogIns(apiCode, "ImpWD4MF10", model_M.MF10003, model.AToken, jsonData, Status, JsonConvert.SerializeObject(responJson));

                        if (string.IsNullOrEmpty(errmsg))
                        {
                            errmsg = "成功";

                            #region 抓發票資料更新發票號碼
                            string INV_NO = await GetSalesOrder(okResult.Value.ToString(), model_M.MF10003);
                            #endregion

                            #region 異動Receivable_D
                            _Rpt.UpdReceivableD(List[i], clientIp, INV_NO, 0, "3");
                            #endregion
                        }
                        else
                        {
                            errmsg = "失敗:" + errmsg;

                            if (errmsg.Contains("無可用的發票號碼"))
                            {
                                #region 異動可用的發票組別
                                _Rpt.UpdInvGp(yyyMM, GroupNo);
                                #endregion
                                i--;
                                errmsg = "跳下一組發票號碼";
                            }
                        }
                        List[i].Win_Msg = errmsg;
                    }
                }

                resultClass.objResult = JsonConvert.SerializeObject(List);
                return resultClass;
            }
            catch (Exception)
            {
                throw;
            }
        }
        /// <summary>
        /// ACH 轉文中開發票to銷帳
        /// </summary>
        /// <param name="List"></param>
        /// <returns></returns>
        [HttpPost("SendReceivableForInv")]
        public async Task<ResultClass<string>> SendReceivableForInv([FromBody] List<Receivable_Win_Inv> List)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();
            ReceivableForInv_req model = new ReceivableForInv_req();
            ReceivableForInv_M_req model_M = new ReceivableForInv_M_req();
            AE_Rpt _Rpt = new AE_Rpt();

            try
            {
                var result = await GetLoginToken(ACompNo);

                if (result is OkObjectResult okResult)
                {
                    var apiName = "rest/TdmServerMethodsIN/ImpWD4MF10";
                    var url = urlBase + apiName;

                    for (int i = 0; i < List.Count; i++){

                        #region 抓取可用的發票組別
                        string yyyMM = (DateTime.Now.Year - 1911).ToString() + DateTime.Now.Month.ToString("D2");
                        var GroupNo = _Rpt.CheckInvGp(yyyMM, "01");
                        #endregion

                        #region maping
                        model.AToken = okResult.Value.ToString();
                        model.ADocType = "20";
                        model.AUpdateType = 0;
                        model.InvoiceGroup = GroupNo;

                        model.ADataSetMaster = new List<ReceivableForInv_M_req>();
                        model.ADataSetDetail = new List<ReceivableForInv_D_req>();

                        model_M.MF10003 = "S" + List[i].HS_id + "-" + List[i].RC_count.ToString("D3");
                        model_M.MF10004 = DateTime.Now.ToString("yyyy-MM-dd");
                        model_M.MF10008 = List[i].CS_PID;
                        model_M.MF10010 = "00";
                        model_M.MF10011 = "00";
                        model_M.MF10012 = "";
                        model_M.MF10018 = "2";
                        model_M.MF10022 = "20";
                        model_M.MF10059 = "1";
                        model_M.MF10066 = DateTime.Now.ToString("yyyy-MM-dd");
                        model_M.MF10091 = "N";
                        model_M.MF10093 = "Y";
                        model_M.MF10094 = "Y";
                        model.ADataSetMaster.Add(model_M);

                        ReceivableForInv_D_req model_D1 = new ReceivableForInv_D_req();
                        model_D1.DT10004 = model_M.MF10003;
                        model_D1.DT10006 = "002";
                        model_D1.DT10030 = 1;
                        model_D1.DT10040 = List[i].interest;
                        model_D1.DT10021 = (List[i].amount_total / 10000) + "萬" + "(" + List[i].RC_count + "/" + List[i].month_total + ")";
                        model.ADataSetDetail.Add(model_D1);

                        ReceivableForInv_D_req model_D2 = new ReceivableForInv_D_req();
                        model_D2.DT10004 = model_M.MF10003;
                        model_D2.DT10006 = "003";
                        model_D2.DT10030 = 1;
                        model_D2.DT10040 = 20;
                        model_D2.DT10021 = (List[i].amount_total / 10000) + "萬" + "(" + List[i].RC_count + "/" + List[i].month_total + ")";
                        model.ADataSetDetail.Add(model_D2);
                        #endregion

                        var jsonData = "";
                        jsonData = JsonConvert.SerializeObject(model);
                        var content = new StringContent(jsonData, Encoding.UTF8, "application/json");

                        var response = await _httpClient.PostAsync(url, content);
                        response.EnsureSuccessStatusCode();

                        var responseContent = await response.Content.ReadAsStringAsync();
                        var responJson = JObject.Parse(responseContent);

                        string errmsg = "";
                        if (responJson["data"]["result"] != null && responJson["data"]["result"].HasValues)
                        {
                            errmsg = (string)responJson["data"]["result"][0]?["errmsg"];
                        }
                        else
                        {
                            errmsg = "";
                        }

                        string Status = (string)responJson["status"];

                        _fun.ExtAPILogIns(apiCode, "ImpWD4MF10", model_M.MF10003, model.AToken, jsonData, Status, JsonConvert.SerializeObject(responJson));

                        if (string.IsNullOrEmpty(errmsg) && Status == "200")
                        {
                            errmsg = "成功";

                            #region 抓發票資料更新發票號碼
                            string INV_NO = await GetSalesOrder(okResult.Value.ToString(), model_M.MF10003);
                            #endregion

                            #region 異動Receivable_D
                            _Rpt.UpdReceivableD(List[i], clientIp, INV_NO, List[i].amount_per_month, "2");
                            #endregion
                        }
                        else
                        {
                            errmsg = "失敗:" + errmsg;

                           

                            if (errmsg.Contains("無可用的發票號碼"))
                            {
                                #region 異動可用的發票組別
                                _Rpt.UpdInvGp(yyyMM, GroupNo);
                                #endregion
                                i--;
                                errmsg = "跳下一組發票號碼";
                            }

                            if (Status == "500")
                            {
                                errmsg += (string)responJson["error"];
                            }
                        }
                        //20260422 Bug修正
                        if (List.ElementAtOrDefault(i) != null)
                        {
                            List[i].Win_Msg = errmsg;
                        }
                    }
                }

                resultClass.objResult = JsonConvert.SerializeObject(List);
                return resultClass;
            }
            catch (Exception ex)
            {
                _fun.ExtAPILogIns(apiCode, "ImpWD4MF10", model_M.MF10003, model.AToken, ex.Message, "500", ex.Message);
                throw;
            }
        }
        /// <summary>
        /// 自行繳款 轉文中開發票to銷帳
        /// </summary>
        /// <param name="List"></param>
        /// <returns></returns>
        [HttpPost("SendPaySelfForInv")]
        public async Task<ResultClass<string>> SendPaySelfForInv([FromBody] List<PaySelf_Win_Inv> List)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();
            ReceivableForInv_req model = new ReceivableForInv_req();
            ReceivableForInv_M_req model_M = new ReceivableForInv_M_req();
            AE_Rpt _Rpt = new AE_Rpt();

            try
            {
                var apiName = "rest/TdmServerMethodsIN/ImpWD4MF10";
                var url = urlBase + apiName;

                var result = await GetLoginToken(ACompNo);

                if (result is OkObjectResult okResult)
                {
                    for (int i = 0; i < List.Count; i++)
                    {
                        #region 抓取可用的發票組別
                        string yyyMM = (DateTime.Now.Year - 1911).ToString() + DateTime.Now.Month.ToString("D2");
                        var GroupNo = _Rpt.CheckInvGp(yyyMM, "01");
                        #endregion

                        #region maping
                        model.AToken = okResult.Value.ToString();
                        model.ADocType = "20";
                        model.AUpdateType = 0;
                        model.InvoiceGroup = GroupNo;

                        model.ADataSetMaster = new List<ReceivableForInv_M_req>();
                        model.ADataSetDetail = new List<ReceivableForInv_D_req>();

                        model_M.MF10003 = "S" + List[i].HS_id + "-" + List[i].RC_count.ToString("D3");
                        model_M.MF10004 = DateTime.Now.ToString("yyyy-MM-dd");
                        model_M.MF10008 = List[i].CS_PID;
                        model_M.MF10010 = "00";
                        model_M.MF10011 = "00";
                        model_M.MF10012 = "";
                        model_M.MF10018 = "2";
                        model_M.MF10022 = "20";
                        model_M.MF10059 = "1";
                        model_M.MF10066 = DateTime.Now.ToString("yyyy-MM-dd");
                        model_M.MF10091 = "N";
                        model_M.MF10093 = "Y";
                        model_M.MF10094 = "Y";
                        model.ADataSetMaster.Add(model_M);

                        ReceivableForInv_D_req model_D1 = new ReceivableForInv_D_req();
                        model_D1.DT10004 = model_M.MF10003;
                        model_D1.DT10006 = "002";
                        model_D1.DT10030 = 1;
                        model_D1.DT10040 = List[i].interest;
                        model_D1.DT10021 = (List[i].amount_total / 10000) + "萬" + "(" + List[i].RC_count + "/" + List[i].month_total + ")";
                        model.ADataSetDetail.Add(model_D1);

                        ReceivableForInv_D_req model_D2 = new ReceivableForInv_D_req();
                        model_D2.DT10004 = model_M.MF10003;
                        model_D2.DT10006 = "003";
                        model_D2.DT10030 = 1;
                        model_D2.DT10040 = 20;
                        model_D2.DT10021 = (List[i].amount_total / 10000) + "萬" + "(" + List[i].RC_count + "/" + List[i].month_total + ")";
                        model.ADataSetDetail.Add(model_D2);
                        #endregion

                        var jsonData = "";
                        jsonData = JsonConvert.SerializeObject(model);
                        var content = new StringContent(jsonData, Encoding.UTF8, "application/json");

                        var response = await _httpClient.PostAsync(url, content);
                        response.EnsureSuccessStatusCode();

                        var responseContent = await response.Content.ReadAsStringAsync();
                        var responJson = JObject.Parse(responseContent);

                        string errmsg = (string)responJson["data"]["result"][0]["errmsg"];
                        string Status = (string)responJson["status"];

                        _fun.ExtAPILogIns(apiCode, "ImpWD4MF10", model_M.MF10003, model.AToken, jsonData, Status, JsonConvert.SerializeObject(responJson));

                        if (string.IsNullOrEmpty(errmsg))
                        {
                            errmsg = "成功";

                            #region 抓發票資料更新發票號碼
                            string INV_NO = await GetSalesOrder(okResult.Value.ToString(), model_M.MF10003);
                            #endregion

                            #region 異動Receivable_D
                            var cpPayAmt = List[i].CP_Pay_Amt ?? 0;
                            List[i].RC_note = List[i].CP_bus_remark;
                            _Rpt.UpdReceivableD(List[i], clientIp, INV_NO, cpPayAmt, "1");
                            #endregion

                            #region 異動自主繳款資料
                            _Rpt.UpdClientPay(List[i].RCD_id.ToString(), List[i].User);
                            #endregion
                        }
                        else
                        {
                            errmsg = "失敗:" + errmsg;

                            if (errmsg.Contains("無可用的發票號碼"))
                            {
                                #region 異動可用的發票組別
                                _Rpt.UpdInvGp(yyyMM, GroupNo);
                                #endregion
                                i--;
                                errmsg = "跳下一組發票號碼";
                            }
                        }
                        List[i].Win_Msg = errmsg;
                    }
                }

                resultClass.objResult = JsonConvert.SerializeObject(List);
                return resultClass;

            }
            catch (Exception ex)
            {
                _fun.ExtAPILogIns(apiCode, "ImpWD4MF10", model_M.MF10003, model.AToken, ex.Message, "500", ex.Message);
                throw;
            }
        }

        private async Task<string> GetSalesOrder(string Token,string ID)
        {
            var apiName = "rest/TdmServerMethodsIN/ExpWD4MF10";
            var url = urlBase + apiName;

            try
            {
                var INV_NO = "";

                SalesOrder_req model = new SalesOrder_req();
                model.AToken = Token;
                model.ADocType = "20";
                model.AExpRange = "2";
                model.ANoB = ID;
                model.ANoE = ID;

                var jsonData = "";
                jsonData = JsonConvert.SerializeObject(model);
                var content = new StringContent(jsonData, Encoding.UTF8, "application/json");

                var response = await _httpClient.PostAsync(url, content);
                response.EnsureSuccessStatusCode();

                var responseContent = await response.Content.ReadAsStringAsync();
                var responJson = JObject.Parse(responseContent);

                string Status = (string)responJson["status"];

                if (Status == "200")
                {
                    INV_NO = (string)responJson["data"]["adatasetmaster"][0]["mf10019"];
                }

                _fun.ExtAPILogIns(apiCode, "ExpWD4MF10", ID, model.AToken, jsonData, Status, JsonConvert.SerializeObject(responJson));

                return INV_NO;
            }
            catch (Exception)
            {

                throw;
            }
            
        }
    }
}
