using KF_WebAPI.BaseClass;
using KF_WebAPI.BaseClass.Max104;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using KF_WebAPI.FunctionHandler;
using KF_WebAPI.DataLogic;
using System.Data;

using static KF_WebAPI.FunctionHandler.Common;
using Microsoft.Data.SqlClient;
using System;
using Azure.Core;
using KF_WebAPI.BaseClass.AE;

namespace KF_WebAPI.Controllers
{



    [ApiController]
    public class A_Max104Controller : Controller
    {

        /// <summary>
        /// 是否呼叫測試API,true:用資料庫模擬API,false:呼叫裕富API;
        /// </summary>
        private readonly HttpClient _httpClient;
        private Common _Comm = new();
        private KFData _KFData = new();
        private readonly string g_CO_CODE = "52611690";
        private readonly string g_CO_ID = "1";
        public A_Max104Controller(IHttpClientFactory httpClientFactory)
        {
            _httpClient = httpClientFactory.CreateClient();
          
        }


        /// <summary>
        /// 104登入,取得ACCESS_TOKEN
        /// </summary>
        /// <param name="USER_ACCOUNT"></param>
        /// <param name="USER_PWD"></param>
        /// <returns></returns>
        [Route("Login")]
        [HttpPost]
        public async Task<ActionResult<string>> Login(string UserID)
        {
            string TransactionId = DateTime.Now.ToString("yyyyMMddhhmmssffff");
            ResultClass_104<signIn> resultClass = new ResultClass_104<signIn>();
            try
            {
                string USER_ACCOUNT = "KF_ERP";
                string USER_PWD = "Kf52611690";
                string p_JSON = "{ \"USER_ACCOUNT\": \"" + USER_ACCOUNT + "\", \"USER_PWD\": \"" + USER_PWD + "\"}";

                ResultClass<string> APIResult = new();
                APIResult = await _Comm.Call_104API(UserID,"auth/signIn", p_JSON, TransactionId, _httpClient);

                if (APIResult.ResultCode == "200")
                {
                    resultClass = JsonConvert.DeserializeObject<ResultClass_104<signIn>>(APIResult.objResult);
                }
                else
                {
                    resultClass.code = "999";
                    resultClass.msg = $"{APIResult.ResultMsg}";
                }
            }
            catch (Exception ex)
            {
                resultClass.code = "999";
                resultClass.msg = $" response: {ex.Message}";
            }
            return Ok(resultClass);
        }


        /// <summary>
        /// 取得104的公司ID
        /// </summary>
        /// <param name="ACCESS_TOKEN"></param>
        /// <returns></returns>
        [Route("GetCompany")]
        [HttpPost]
        public async Task<ActionResult<string>> GetCompany(string UserID,string ACCESS_TOKEN)
        {
            string TransactionId = DateTime.Now.ToString("yyyyMMddhhmmssffff");
            arrResultClass_104<company> resultClass = new arrResultClass_104<company>();
            try
            {

                string p_JSON = "{ \"ACCESS_TOKEN\": \"" + ACCESS_TOKEN + "\"}";

                ResultClass<string> APIResult = new();
                APIResult = await _Comm.Call_104API(UserID, "os/company", p_JSON, TransactionId, _httpClient, ACCESS_TOKEN);

                if (APIResult.ResultCode == "200")
                {
                    resultClass = JsonConvert.DeserializeObject<arrResultClass_104<company>>(APIResult.objResult);
                }
                else
                {
                    resultClass.code = "999";
                    resultClass.msg = $"{APIResult.ResultMsg}";
                }
            }
            catch (Exception ex)
            {
                resultClass.code = "999";
                resultClass.msg = $" response: {ex.Message}";
            }
            return Ok(resultClass);
        }



        [Route("GetEmp_ID")]
        [HttpPost]
        public async Task<ActionResult<string>> GetEmp_ID(string UserID, string EMP_NO, string ACCESS_TOKEN)
        {
            string TransactionId = DateTime.Now.ToString("yyyyMMddhhmmssffff");
            arrResultClass_104<emp_id> resultClass = new arrResultClass_104<emp_id>();
            try
            {

                string p_JSON = "{ \"ACCESS_TOKEN\": \"" + ACCESS_TOKEN + "\", \"CO_CODE\": \"" + g_CO_CODE + "\", \"EMP_NO\": \"" + EMP_NO + "\"}";

                ResultClass<string> APIResult = new();
                APIResult = await _Comm.Call_104API(UserID,"ed/emp_id", p_JSON, TransactionId, _httpClient, ACCESS_TOKEN);

                if (APIResult.ResultCode == "200")
                {
                    resultClass = JsonConvert.DeserializeObject<arrResultClass_104<emp_id>>(APIResult.objResult);
                }
                else
                {
                    resultClass.code = "999";
                    resultClass.msg = $"{APIResult.ResultMsg}";
                }
            }
            catch (Exception ex)
            {
                resultClass.code = "999";
                resultClass.msg = $" response: {ex.Message}";
            }
            return Ok(resultClass);
        }

        /// <summary>
        /// 取得104的員工ID更新回USER_M (ID_104)
        /// </summary>
        /// <param name="ACCESS_TOKEN"></param>
        /// <returns></returns>
        [Route("UPD_EmpIdToKF_ERP")]
        [HttpPost]
        public async Task<ActionResult<string>> UPD_EmpIdToKF_ERP(string UserID, string ACCESS_TOKEN)
        {
            string TransactionId = DateTime.Now.ToString("yyyyMMddhhmmssffff");
            arrResultClass_104<emp_id> resultClass = new arrResultClass_104<emp_id>();
            try
            {
                string p_JSON = "{ \"ACCESS_TOKEN\": \"" + ACCESS_TOKEN + "\", \"CO_ID\": 1, \"LIMIT\": 300}";
                ResultClass<string> APIResult = new();
                APIResult = await _Comm.Call_104API(UserID,"ed/emp", p_JSON, TransactionId, _httpClient, ACCESS_TOKEN);
                ADOData _adoData = new ADOData();
                if (APIResult.ResultCode == "200")
                {
                    resultClass = JsonConvert.DeserializeObject<arrResultClass_104<emp_id>>(APIResult.objResult);
                    if (resultClass.code == "200")
                    {
                        string m_SQL = "Update User_M set ID_104=@ID_104 where U_num=@U_num ";
                        foreach (emp_id _emp_id in ((emp_id[])resultClass.data))
                        {
                            var parameters = new List<SqlParameter>();
                            parameters.Add(new SqlParameter("@ID_104", _emp_id.EMP_ID));
                            parameters.Add(new SqlParameter("@U_num", _emp_id.EMP_NO));
                            int result = _adoData.ExecuteNonQuery(m_SQL, parameters);
                        }
                    }
                }
                else
                {
                    resultClass.code = "999";
                    resultClass.msg = $"{APIResult.ResultMsg}";
                }
            }
            catch (Exception ex)
            {
                resultClass.code = "999";
                resultClass.msg = $" response: {ex.Message}";
            }
            return Ok(resultClass);
        }


        [Route("batchOtNew")]
        [HttpPost]
        public async Task<ActionResult<string>> batchOtNew(string UserID, string ACCESS_TOKEN, string FR_kind, string Date_S, string Date_E)
        {
            AE_HR m_AE_HR = new AE_HR();
            string SESSION_KEY = DateTime.Now.ToString("yyyyMMddhhmmssffff");
            batchOtNew ReqClass = m_AE_HR.GetAE_OT_DATA(SESSION_KEY, FR_kind, Date_S, Date_E);

            ResultClass_104<object> resultClass = new ResultClass_104<object>();
            try
            {
                string p_JSON = JsonConvert.SerializeObject(ReqClass);

                ResultClass<string> APIResult = new();
                APIResult = await _Comm.Call_104API(UserID,"wf/wf021/batchOtNew", p_JSON, ReqClass.SESSION_KEY, _httpClient, ACCESS_TOKEN);

                if (APIResult.ResultCode == "200")
                {
                    resultClass = JsonConvert.DeserializeObject<ResultClass_104<object>>(APIResult.objResult);

                    m_AE_HR.ModifyAE_SESSION_KEY("Flow_rest", SESSION_KEY, FR_kind, Date_S, Date_E);
                }
                else
                {
                    resultClass.code = "999";
                    resultClass.msg = $"{APIResult.ResultMsg}";
                    if (APIResult.objResult != null)
                    {
                        resultClass.data = APIResult.objResult;
                    }
                }
               
            }
            catch (Exception ex)
            {
                resultClass.code = "999";
                resultClass.msg = $" response: {ex.Message}";
            }
            return Ok(resultClass);
        }


        [Route("batchOtSign")]
        [HttpPost]
        public async Task<ActionResult<string>> batchOtSign(string UserID, string ACCESS_TOKEN, [FromBody] batchSign ReqClass)
        {
            
            ResultClass_104<Object> resultClass = new ResultClass_104<Object>();
            try
            {
               
                string p_JSON = JsonConvert.SerializeObject(ReqClass);
                ResultClass<string> APIResult = new();
                APIResult = await _Comm.Call_104API(UserID,"wf/wf021/batchOtSign", p_JSON, ReqClass.SESSION_KEY, _httpClient, ACCESS_TOKEN);

                if (APIResult.ResultCode == "200")
                {
                    resultClass = JsonConvert.DeserializeObject<ResultClass_104<Object>>(APIResult.objResult);
                    AE_HR m_AE_HR = new AE_HR();
                    m_AE_HR.ModifyAE_Sign("Flow_rest", ReqClass.SESSION_KEY);
                }
                else
                {
                    resultClass.code = "999";
                    resultClass.msg = $"{APIResult.ResultMsg}";
                }
                
            }
            catch (Exception ex)
            {
                resultClass.code = "999";
                resultClass.msg = $" response: {ex.Message}";
            }
            return Ok(resultClass);
        }


        [Route("batchOtDelete")]
        [HttpPost]
        public async Task<ActionResult<string>> batchOtDelete([FromBody] batchSign ReqClass, string UserID, string TOKEN_KEY)
        {

            ResultClass_104<object> resultClass = new ResultClass_104<object>();
            try
            {

                string p_JSON = JsonConvert.SerializeObject(ReqClass);
                ResultClass<string> APIResult = new();
                APIResult = await _Comm.Call_104API(UserID, "wf/wf021/batchOtDelete", p_JSON, ReqClass.SESSION_KEY, _httpClient, TOKEN_KEY);

                if (APIResult.ResultCode == "200")
                {
                    resultClass = JsonConvert.DeserializeObject<ResultClass_104<object>>(APIResult.objResult);
                }
                else
                {
                    resultClass.code = "999";
                    resultClass.msg = $"{APIResult.ResultMsg}";
                }
            }
            catch (Exception ex)
            {
                resultClass.code = "999";
                resultClass.msg = $" response: {ex.Message}";
            }
            return Ok(resultClass);
        }


        [Route("deleteOt")]
        [HttpPost]
        public async Task<ActionResult<string>> deleteOt([FromBody] deleteOt ReqClass, string UserID, string TOKEN_KEY)
        {
            string SESSION_KEY = DateTime.Now.ToString("yyyyMMddhhmmssffff");
            ResultClass_104<object> resultClass = new ResultClass_104<object>();
            try
            {
                string p_JSON = JsonConvert.SerializeObject(ReqClass);

                ResultClass<string> APIResult = new();
                APIResult = await _Comm.Call_104API(  UserID,"wf/wf020/deleteOt", p_JSON, SESSION_KEY, _httpClient, TOKEN_KEY);

                if (APIResult.ResultCode == "200")
                {
                    resultClass = JsonConvert.DeserializeObject<ResultClass_104<object>>(APIResult.objResult);
                }
                else
                {
                    resultClass.code = "999";
                    resultClass.msg = $"{APIResult.ResultMsg}";
                }
            }
            catch (Exception ex)
            {
                resultClass.code = "999";
                resultClass.msg = $" response: {ex.Message}";
            }
            return Ok(resultClass);
        }


        [Route("batchLeaveSign")]
        [HttpPost]
        public async Task<ActionResult<string>> batchLeaveSign(string UserID, string ACCESS_TOKEN, [FromBody] batchSign ReqClass)
        {

            ResultClass_104<Object> resultClass = new ResultClass_104<Object>();
            try
            {
                string p_JSON = JsonConvert.SerializeObject(ReqClass);
                ResultClass<string> APIResult = new();
                APIResult = await _Comm.Call_104API(UserID, "wf/wf011/batchLeaveSign", p_JSON, ReqClass.SESSION_KEY, _httpClient, ACCESS_TOKEN);

                if (APIResult.ResultCode == "200")
                {
                    resultClass = JsonConvert.DeserializeObject<ResultClass_104<Object>>(APIResult.objResult);
                    AE_HR m_AE_HR = new AE_HR();
                    m_AE_HR.ModifyAE_Sign("Flow_rest", ReqClass.SESSION_KEY);
                }
                else
                {
                    resultClass.code = "999";
                    resultClass.msg = $"{APIResult.ResultMsg}";
                }
               
            }
            catch (Exception ex)
            {
                resultClass.code = "999";
                resultClass.msg = $" response: {ex.Message}";
            }
            return Ok(resultClass);
        }


        [Route("batchLeaveNew")]
        [HttpPost]
        public async Task<ActionResult<string>> batchLeaveNew(string UserID, string ACCESS_TOKEN, string FR_kind, string Date_S, string Date_E)
        {
            AE_HR m_AE_HR = new AE_HR();
            string SESSION_KEY = DateTime.Now.ToString("yyyyMMddhhmmssffff");
            batchLeaveNew ReqClass = m_AE_HR.GetAE_LEAVE_DATA(SESSION_KEY, FR_kind,  Date_S,  Date_E);
            ResultClass_104<object> resultClass = new ResultClass_104<object>();
            try
            {

                string p_JSON = JsonConvert.SerializeObject(ReqClass);
                ResultClass<string> APIResult = new();
                APIResult = await _Comm.Call_104API(UserID, "wf/wf011/batchLeaveNew", p_JSON, ReqClass.SESSION_KEY, _httpClient, ACCESS_TOKEN);

                if (APIResult.ResultCode == "200")
                {
                    resultClass = JsonConvert.DeserializeObject<ResultClass_104<object>>(APIResult.objResult);
                    
                    m_AE_HR.ModifyAE_SESSION_KEY("Flow_rest", SESSION_KEY, FR_kind, Date_S, Date_E);
                }
                else
                {
                    resultClass.code = "999";
                    resultClass.msg = $"{APIResult.ResultMsg}";
                    if (APIResult.objResult != null)
                    {
                        resultClass.data = APIResult.objResult;
                    }
                }
                
            }
            catch (Exception ex)
            {
                resultClass.code = "999";
                resultClass.msg = $" response: {ex.Message}";
            }
            return Ok(resultClass);
        }


        [Route("batchLeaveLate15")]
        [HttpPost]
        public async Task<ActionResult<string>> batchLeaveLate15(string UserID, string ACCESS_TOKEN, string FR_kind, string Date_S, string Date_E)
        {
            AE_HR m_AE_HR = new AE_HR();
            string SESSION_KEY = DateTime.Now.ToString("yyyyMMddhhmmssffff");
            
            ResultClass_104<object> resultClass = new ResultClass_104<object>();
            try
            {
               
                    batchLeaveNew ReqClass = m_AE_HR.GetLate15To104(SESSION_KEY,  Date_S, Date_E);
                    string p_JSON = JsonConvert.SerializeObject(ReqClass);
                    ResultClass<string> APIResult = new();
                    APIResult = await _Comm.Call_104API(UserID, "wf/wf011/batchLeaveNew", p_JSON, ReqClass.SESSION_KEY, _httpClient, ACCESS_TOKEN);

                    if (APIResult.ResultCode == "200")
                    {
                        resultClass = JsonConvert.DeserializeObject<ResultClass_104<object>>(APIResult.objResult);

                        m_AE_HR.ModifyAE_SESSION_KEY("Late15To104", SESSION_KEY, FR_kind, Date_S, Date_E);
                    }
                    else
                    {
                        resultClass.code = "999";
                        resultClass.msg = $"{APIResult.ResultMsg}";
                        if (APIResult.objResult != null)
                        {
                            resultClass.data = APIResult.objResult;
                        }
                    }
               
                

            }
            catch (Exception ex)
            {
                resultClass.code = "999";
                resultClass.msg = $" response: {ex.Message}";
            }
            return Ok(resultClass);
        }

        [Route("batchLeaveLate15Sign")]
        [HttpPost]
        public async Task<ActionResult<string>> batchLeaveLate15Sign(string UserID, string ACCESS_TOKEN, [FromBody] batchSign ReqClass)
        {

            ResultClass_104<Object> resultClass = new ResultClass_104<Object>();
            try
            {
                string p_JSON = JsonConvert.SerializeObject(ReqClass);
                ResultClass<string> APIResult = new();
                APIResult = await _Comm.Call_104API(UserID, "wf/wf011/batchLeaveSign", p_JSON, ReqClass.SESSION_KEY, _httpClient, ACCESS_TOKEN);

                if (APIResult.ResultCode == "200")
                {
                    resultClass = JsonConvert.DeserializeObject<ResultClass_104<Object>>(APIResult.objResult);
                    AE_HR m_AE_HR = new AE_HR();
                    m_AE_HR.ModifyAE_Sign("Late15To104", ReqClass.SESSION_KEY);
                }
                else
                {
                    resultClass.code = "999";
                    resultClass.msg = $"{APIResult.ResultMsg}";
                }

            }
            catch (Exception ex)
            {
                resultClass.code = "999";
                resultClass.msg = $" response: {ex.Message}";
            }
            return Ok(resultClass);
        }



        [Route("batchLeaveDelete")]
        [HttpPost]
        public async Task<ActionResult<string>> batchLeaveDelete([FromBody] batchSign ReqClass, string UserID, string TOKEN_KEY)
        {

            ResultClass_104<object> resultClass = new ResultClass_104<object>();
            try
            {
              
                string p_JSON = JsonConvert.SerializeObject(ReqClass);

                ResultClass<string> APIResult = new();
                APIResult = await _Comm.Call_104API(UserID, "wf/wf011/batchLeaveDelete", p_JSON, ReqClass.SESSION_KEY, _httpClient, TOKEN_KEY);

                if (APIResult.ResultCode == "200")
                {
                    resultClass = JsonConvert.DeserializeObject<ResultClass_104<object>>(APIResult.objResult);
                }
                else
                {
                    resultClass.code = "999";
                    resultClass.msg = $"{APIResult.ResultMsg}";
                }
            }
            catch (Exception ex)
            {
                resultClass.code = "999";
                resultClass.msg = $" response: {ex.Message}";
            }
            return Ok(resultClass);
        }


        /// <summary>
        /// 取得104行事曆類別
        /// </summary>
        /// <param name="ACCESS_TOKEN"></param>
        /// <returns></returns>
        [Route("calendar_leave")]
        [HttpPost]
        public async Task<ActionResult<string>> calendar_leave(string UserID, string ACCESS_TOKEN)
        {
            string TransactionId = DateTime.Now.ToString("yyyyMMddhhmmssffff");
            arrResultClass_104<calendar_leave> resultClass = new arrResultClass_104<calendar_leave>();
            try
            {
                string p_JSON = "{ \"ACCESS_TOKEN\": \"" + ACCESS_TOKEN + "\"}";
                ResultClass<string> APIResult = new();
                APIResult = await _Comm.Call_104API(UserID, "am/calendar_leave", p_JSON, TransactionId, _httpClient, ACCESS_TOKEN);
                ADOData _adoData = new ADOData();
                if (APIResult.ResultCode == "200")
                {
                    resultClass = JsonConvert.DeserializeObject<arrResultClass_104<calendar_leave>>(APIResult.objResult);
                }
                else
                {
                    resultClass.code = "999";
                    resultClass.msg = $"{APIResult.ResultMsg}";
                }
            }
            catch (Exception ex)
            {
                resultClass.code = "999";
                resultClass.msg = $" response: {ex.Message}";
            }
            return Ok(resultClass);
        }


        /// <summary>
        /// 取得行事曆基本資料
        /// </summary>
        /// <param name="ACCESS_TOKEN"></param>
        /// <returns></returns>
        [Route("calendar_basic")]
        [HttpPost]
        public async Task<ActionResult<string>> calendar_basic(string UserID, string ACCESS_TOKEN)
        {
            string TransactionId = DateTime.Now.ToString("yyyyMMddhhmmssffff");
            arrResultClass_104<object> resultClass = new arrResultClass_104<object>();
            try
            {
                string p_JSON = "{ \"ACCESS_TOKEN\": \"" + ACCESS_TOKEN + "\",\"CO_ID\": 1}";
                ResultClass<string> APIResult = new();
                APIResult = await _Comm.Call_104API(UserID, "am/calendar_basic", p_JSON, TransactionId, _httpClient, ACCESS_TOKEN);
                ADOData _adoData = new ADOData();
                if (APIResult.ResultCode == "200")
                {
                    resultClass = JsonConvert.DeserializeObject<arrResultClass_104<Object>>(APIResult.objResult);
                }
                else
                {
                    resultClass.code = "999";
                    resultClass.msg = $"{APIResult.ResultMsg}";
                }
            }
            catch (Exception ex)
            {
                resultClass.code = "999";
                resultClass.msg = $" response: {ex.Message}";
            }
            return Ok(resultClass);
        }



        /// <summary>
        /// 
        /// </summary>
        /// <param name="ACCESS_TOKEN"></param>
        /// <param name="CALENDAR_YEAR"></param>
        /// <returns></returns>
        [Route("calendar_day")]
        [HttpPost]
        public async Task<ActionResult<string>> calendar_day(string UserID, string ACCESS_TOKEN,string CALENDAR_YEAR)
        {
            string TransactionId = DateTime.Now.ToString("yyyyMMddhhmmssffff");
            ResultClass_104<string> resultClass = new ResultClass_104<string>();
            try
            {
                string p_JSON = "{ \"ACCESS_TOKEN\": \"" + ACCESS_TOKEN + "\",\"CO_ID\": 1,\"CALENDAR_BASIC_ID\": 1,\"CALENDAR_YEAR\": " + CALENDAR_YEAR + "}";
                ResultClass<string> APIResult = new();
                APIResult = await _Comm.Call_104API(UserID, "am/calendar_day", p_JSON, TransactionId, _httpClient, ACCESS_TOKEN);
              
                if (APIResult.ResultCode == "200")
                {
                    arrResultClass_104<calendar_day> objResult = JsonConvert.DeserializeObject<arrResultClass_104<calendar_day>>(APIResult.objResult);
                    AE_HR m_hr = new AE_HR();
                    m_hr.InsertAE_Calendar_day(objResult, CALENDAR_YEAR);
                    resultClass.code = "200";
                }
                else
                {
                    resultClass.code = "999";
                    resultClass.msg = $"{APIResult.ResultMsg}";
                }
            }
            catch (Exception ex)
            {
                resultClass.code = "999";
                resultClass.msg = $" response: {ex.Message}";
            }
            return Ok(resultClass);
        }




        /// <summary>
        /// 一般假勤項目檔
        /// </summary>
        /// <param name="ACCESS_TOKEN"></param>
        /// <returns></returns>
        [Route("leaveitem")]
        [HttpPost]
        public async Task<ActionResult<string>>leaveitem(string UserID, string ACCESS_TOKEN)
        {
            string TransactionId = DateTime.Now.ToString("yyyyMMddhhmmssffff");
            arrResultClass_104<leaveitem> resultClass = new arrResultClass_104<leaveitem>();
            try
            {
                string p_JSON = "{ \"ACCESS_TOKEN\": \"" + ACCESS_TOKEN + "\",\"CO_ID\": 1}";
                ResultClass<string> APIResult = new();
                APIResult = await _Comm.Call_104API(UserID, "am/leaveitem", p_JSON, TransactionId, _httpClient, ACCESS_TOKEN);

                if (APIResult.ResultCode == "200")
                {
                    arrResultClass_104<leaveitem> objResult = JsonConvert.DeserializeObject<arrResultClass_104<leaveitem>>(APIResult.objResult);
                    AE_HR m_hr = new AE_HR();
                    m_hr.InsertAE_Leaveitem(objResult);
                    resultClass.code = "200";
                }
                else
                {
                    resultClass.code = "999";
                    resultClass.msg = $"{APIResult.ResultMsg}";
                }
            }
            catch (Exception ex)
            {
                resultClass.code = "999";
                resultClass.msg = $" response: {ex.Message}";
            }
            return Ok(resultClass);
        }






    }

}
