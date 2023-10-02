using KF_WebAPI.BaseClass;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using KF_WebAPI.FunctionHandler;
using KF_WebAPI.DataLogic;
using System.Data;

using static KF_WebAPI.FunctionHandler.Common;
using System.Data.SqlClient;

namespace KF_WebAPI.Controllers
{
    public class FileObject
    {
        /// <summary>
        /// 回覆代碼
        /// </summary>
        public string? form_no { get; set; }
        public string? salesNo { get; set; }
        /// <summary>
        /// 回覆訊息 
        /// </summary>
        public attachmentFile[]? attachmentFiles { get; set; }
    }


    [ApiController]
    public class YuRichAPIController : Controller
    {

        /// <summary>
        /// 是否呼叫測試API,true:用資料庫模擬API,false:呼叫裕富API;
        /// </summary>
        public Boolean _isCallTESTAPI;
        private readonly HttpClient _httpClient;
        private readonly string _branchNo = "0001";
        private readonly string _dealerNo = "MM09";
        private readonly string _source = "22";
        private AesEncryption _AE = new();
        private Common _Comm = new();
        private KFData _KFData = new();

        public YuRichAPIController(IHttpClientFactory httpClientFactory)
        {
            _httpClient = httpClientFactory.CreateClient();
            //定義測試與否
            _isCallTESTAPI = _Comm.isCallTESTAPI;
        }

        #region  內部API
        [Route("GetFileBySeq")]
        [HttpPost]
        public async Task<ActionResult<string>> GetFileBySeq(string KeyID, string Type, string Index)
        {
            ResultClass<attachmentFile> resultClass = new();
            try
            {
                resultClass = _KFData.GetFile(KeyID, Type, Index);
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "999";
                resultClass.ResultMsg = ex.Message;
            }
            return Ok(resultClass);
        }

        [Route("UpdReceive")]
        [HttpPost]
        public ActionResult<ResultClass<BaseResult>> UpdReceive([FromBody] objInsertReceive objects, string form_no, string salesNo, string Case_Company)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            try
            {

                if (objects._Receive1 != null)
                {
                    DateTime now = DateTime.Now;
                    string formattedDateTime = now.ToString("yyyy-MM-dd HH:mm:ss.fff");

                    objects._Receive1.Action = "upd";
                    objects._Receive1.Case_Company = Case_Company;
                    objects._Receive1.form_no = form_no;
                    objects._Receive1.upd_user = salesNo;
                    objects._Receive1.upd_date = formattedDateTime;
                    objects._Receive1.status = "1";
                    objects._Receive1.casestatus = "0";
                }

                    ResultClass<int> m_Insert = _Comm.UpdateDataByClass("tbReceive","form_no", objects._Receive1);
                if (m_Insert.ResultCode == "000")
                {
                    if (objects._Receive1.attachmentFile != null)
                    {
                        ResultClass<int> m_InsertFile = _KFData.InsertFile(form_no, "Receive", salesNo, "", objects._Receive1.attachmentFile);
                        if (m_InsertFile.ResultCode != "000")
                        {
                            resultClass.ResultCode = "999";
                            resultClass.ResultMsg ="上傳檔案失敗!!"+ m_InsertFile.ResultMsg;
                            return Ok(resultClass);
                        }
                    }
                    resultClass.ResultCode = "000";
                    resultClass.ResultMsg = form_no;
                }
                else
                {
                    resultClass.ResultCode = "999";
                    resultClass.ResultMsg = "更新失敗;" + m_Insert.ResultMsg;
                }

            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "999";
                resultClass.ResultMsg = "更新失敗;" + ex.Message;
            }
            return Ok(resultClass);
        }

        [Route("InsertReceive")]
        [HttpPost]
        public ActionResult<ResultClass<BaseResult>> InsertReceive([FromBody] objInsertReceive objects, string salesNo,string Case_Company)
        {
           
            ResultClass<string> resultClass = new ResultClass<string>();
            try
            {
                ResultClass<string> resultSeq = _KFData.GetSeqNo(Case_Company);
                if (resultSeq is not null && resultSeq.ResultCode == "000") 
                {
                    if (objects._Receive1 != null)
                    {
                        objects._Receive1.Action = "Add";
                        objects._Receive1.Case_Company = Case_Company;
                        objects._Receive1.form_no = resultSeq.objResult;
                        objects._Receive1.add_user = salesNo;
                        objects._Receive1.upd_user = salesNo;
                        objects._Receive1.status = "1";
                        objects._Receive1.casestatus = "0";
                        

                        ResultClass<int> m_Insert = _Comm.InsertDataByClass("tbReceive", objects._Receive1);
                        if (m_Insert.ResultCode == "000")
                        {
                            if (objects._Receive1.attachmentFile != null)
                            {
                                ResultClass<int> m_InsertFile = _KFData.InsertFile(resultSeq.objResult, "Receive", salesNo,"", objects._Receive1.attachmentFile);
                                if (m_InsertFile.ResultCode == "000")
                                {
                                    resultClass.ResultCode = "000";
                                    resultClass.ResultMsg =  objects._Receive1.form_no;
                                }
                                else
                                {
                                    resultClass.ResultCode = "999";
                                    resultClass.ResultMsg = "上傳檔案失敗;" + m_InsertFile.ResultMsg;
                                }
                            }
                            else 
                            {
                                resultClass.ResultCode = "999";
                                resultClass.ResultMsg = "上傳檔案失敗;";
                            }
                        }
                        else
                        {
                            resultClass.ResultCode = "999";
                            resultClass.ResultMsg = "存檔失敗;"+ m_Insert.ResultMsg;
                        }
                    }
                    else
                    {
                        resultClass.ResultCode = "999";
                        resultClass.ResultMsg = "存檔失敗!!!";
                    }
                }
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "999";
                resultClass.ResultMsg = "存檔失敗;" + ex.Message;
            }
            return Ok(resultClass);
        }

        [Route("GetYRData")]
        [HttpGet]
        public ActionResult<ResultClass<List<objReceive>>> GetYRData(string form_no)
        {
            ResultClass<List<objReceive>> resultClass = _KFData.GetReceive(form_no);
            return resultClass;
        }

        [Route("Receive")]
        [HttpPost]
        public async Task<ActionResult<string>> Receive([FromBody] Receive ReqClass, string Form_No, string salesNo)
        {
            string TransactionId = DateTime.Now.ToString("yyyyMMddhhmmssffff");
            if (ReqClass is null)
            {
                throw new ArgumentNullException(nameof(ReqClass));
            }
            ResultClass<Result_R> resultClass = new ResultClass<Result_R>();
            try
            {
                string p_JSON = JsonConvert.SerializeObject(ReqClass);
                ResultClass<string> APIResult = new();
                if (_isCallTESTAPI)
                {
                    APIResult = _KFData.GetTestJsonByAPI("TEST0001", "Receive", salesNo, Form_No, TransactionId);
                }
                else
                {
                    APIResult = await _Comm.CallYuRichAPINew("Receive", salesNo, Form_No, p_JSON, TransactionId, _httpClient);
                }
                if (APIResult.ResultCode == "000")
                {
                    Result_R m_Result = JsonConvert.DeserializeObject<Result_R>(APIResult.objResult);
                    if (m_Result.code == "1000")
                    {
                        resultClass.objResult = m_Result;
                        _KFData.UpdReceive(Form_No, APIResult.transactionId, m_Result.examineNo);
                    }
                    else
                    {
                        resultClass.ResultCode = "999";
                        resultClass.ResultMsg = m_Result.msg;
                        _Comm.InsertErrorLog("Receive", TransactionId, m_Result.msg);
                    }
                }
                else
                {
                    resultClass.ResultCode = "999";
                    resultClass.ResultMsg = $"{APIResult.ResultMsg}";
                }
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "999";
                resultClass.ResultMsg = $" response: {ex.Message}";
            }
            return Ok(resultClass);
        }

        [Route("QueryAppropriation")]
        [HttpPost]
        public async Task<ActionResult<string>> QueryAppropriation([FromBody] QueryAppropriation ReqClass, string Form_No, bool isUpdDB = true)
        {
            string TransactionId = DateTime.Now.ToString("yyyyMMddhhmmssffff");
            if (ReqClass is null)
            {
                throw new ArgumentNullException(nameof(ReqClass));
            }
            ResultClass<Result_QA> resultClass = new ResultClass<Result_QA>();
            try
            {
                string m_Uesr = ReqClass.salesNo;
                string p_JSON = JsonConvert.SerializeObject(ReqClass);
                ResultClass<string> APIResult = new ();
                if (_isCallTESTAPI)
                {
                    APIResult  = _KFData.GetTestJsonByAPI("TEST0001", "QueryAppropriation", m_Uesr, Form_No, TransactionId, "TESTQA002");
                }
                else
                {
                    APIResult = await _Comm.CallYuRichAPINew("QueryAppropriation", ReqClass.salesNo, Form_No, p_JSON, TransactionId, _httpClient, isUpdDB);
                }
                if (APIResult.ResultCode == "000")
                {
                    Result_QA m_Result = JsonConvert.DeserializeObject<Result_QA>(APIResult.objResult);
                    if (m_Result.code == "S001")
                    {
                        resultClass.objResult = m_Result;
                        if (isUpdDB)
                        {
                            _KFData.InsertQueryAppropriation(Form_No, m_Uesr, m_Result);
                        }
                    }
                    else
                    {
                        resultClass.ResultCode = "999";
                        resultClass.ResultMsg = m_Result.msg;
                    }
                }
                else
                {
                    resultClass.ResultCode = "999";
                    resultClass.ResultMsg = $"{APIResult.ResultMsg}";
                }
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "999";
                resultClass.ResultMsg = $" response: {ex.Message}";
            }
            return Ok(resultClass);
        }



        public class ReqClass_RP
        {
            public BankInfo? BankInfo { get; set; }
            public RequestPayment? RequestPayment { get; set; }
            public PayInfo? PayInfo { get; set; }
            
        }


        [Route("RequestPayment")]
        [HttpPost]
        public async Task<ActionResult<string>> RequestPayment([FromBody] ReqClass_RP ReqClass,  string Form_No)
        {
            RequestPayment Class_RP= ReqClass.RequestPayment;
            BankInfo Class_BI = ReqClass.BankInfo;
            PayInfo Class_PI = ReqClass.PayInfo;

            ResultClass<BaseResult> resultClass = new ();
            DataTable tbReceive = _KFData.GetReceiveByform_no(Form_No);
            string TransactionId = DateTime.Now.ToString("yyyyMMddhhmmssffff");
            string m_RPUser = Class_RP.salesNo;
            if (tbReceive.Rows.Count != 0)
            {
                foreach (DataRow dr in tbReceive.Rows)
                {
                    Class_RP.branchNo = _branchNo;
                    Class_RP.dealerNo = _dealerNo;
                    Class_RP.examineNo = dr["ExamineNo"].ToString();
                    Class_RP.source = _source;
                    Class_RP.salesNo = "";
                }
               

                if (Class_RP.attachmentFile != null)
                {
                    if (Class_RP.attachmentFile.Length != 6)
                    {
                        resultClass.ResultCode = "999";
                        resultClass.ResultMsg = $"請款文件有遺漏,請確認";
                        return Ok(resultClass);
                    }
                    ResultClass<int> m_InsertFile = _KFData.InsertFile(Form_No, "RequestPayment", m_RPUser, TransactionId, Class_RP.attachmentFile);
                    if (m_InsertFile.ResultCode != "000")
                    {
                        resultClass.ResultCode = "999";
                        resultClass.ResultMsg = $"上傳檔案失敗!!";
                        return Ok(resultClass);
                    }
                    else
                    {
                        if (Class_RP.attachmentFile.Length == 6)
                        {
                            //移除-匯款存摺封面,這個不用傳給裕富
                            Class_RP.attachmentFile = Class_RP.attachmentFile.Take(Class_RP.attachmentFile.Length - 1).ToArray();
                        }
                    }
                }
            }
            if (Class_RP is null)
            {
                throw new ArgumentNullException(nameof(Class_RP));
            }
            try
            {
                string p_JSON = JsonConvert.SerializeObject(Class_RP);
                ResultClass<string> APIResult = new();
                if (_isCallTESTAPI)
                {
                    APIResult = _KFData.GetTestJsonByAPI("TEST0001", "RequestPayment", m_RPUser, Form_No, TransactionId);
                }
                else
                {
                    APIResult = await _Comm.CallYuRichAPINew("RequestPayment", Class_RP.salesNo, Form_No, p_JSON, TransactionId, _httpClient);
                }

                if (APIResult.ResultCode == "000")
                {
                    BaseResult m_Result = JsonConvert.DeserializeObject<BaseResult>(APIResult.objResult);
                    if (m_Result.code == "S001")
                    {
                        resultClass.objResult = m_Result;
                        /*存入匯款資訊*/
                        
                        _KFData.UpdByRequestPayment(Form_No, m_RPUser, TransactionId, Class_BI, Class_PI);
                    }
                    else
                    {
                        resultClass.ResultCode = "999";
                        resultClass.ResultMsg = m_Result.msg;
                        _Comm.InsertErrorLog("RequestPayment", TransactionId, m_Result.msg);

                    }
                }
                else
                {
                    resultClass.ResultCode = "999";
                    resultClass.ResultMsg = $"{APIResult.ResultMsg}";
                }
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "999";
                resultClass.ResultMsg = $" response: {ex.Message}";
            }
            return Ok(resultClass);
        }

        [Route("RequestforExam")]
        [HttpPost]
        public async Task<ActionResult<string>> RequestforExam([FromBody] RequestforExam ReqClass, string Form_No)
        {
            ResultClass<Result_RE> resultClass = new();
            string TransactionId = DateTime.Now.ToString("yyyyMMddhhmmssffff");
            if (ReqClass is null)
            {
                throw new ArgumentNullException(nameof(ReqClass));
            }
            try
            {
                DataTable tbReceive = _KFData.GetReceiveByform_no(Form_No);
                string m_REUser = ReqClass.salesNo;
               
                if (tbReceive.Rows.Count != 0)
                {
                    foreach (DataRow dr in tbReceive.Rows)
                    {
                        ReqClass.branchNo = _branchNo;
                        ReqClass.dealerNo= _dealerNo;
                        ReqClass.examineNo = dr["ExamineNo"].ToString();
                        ReqClass.source = _source;
                        ReqClass.salesNo = "";
                    }
                    if (ReqClass.attachmentFile != null)
                    {
                        ResultClass<int> m_InsertFile = _KFData.InsertFile(Form_No, "RequestforExam", m_REUser, TransactionId,ReqClass.attachmentFile);
                        if (m_InsertFile.ResultCode != "000")
                        {
                            resultClass.ResultCode = "999";
                            resultClass.ResultMsg = $"上傳檔案失敗!!";
                            return Ok(resultClass);
                        }
                    }
                }

                string p_JSON = JsonConvert.SerializeObject(ReqClass);
                ResultClass<string> APIResult = new();
                if (_isCallTESTAPI)
                {
                    APIResult = _KFData.GetTestJsonByAPI("TEST0001", "RequestforExam", m_REUser,Form_No, TransactionId, "TESTRE001");
                }
                else
                {
                    APIResult = await _Comm.CallYuRichAPINew("RequestforExam", ReqClass.salesNo, Form_No, p_JSON, TransactionId, _httpClient);
                }
                if (APIResult.ResultCode == "000")
                {
                    Result_RE m_Result = JsonConvert.DeserializeObject<Result_RE>(APIResult.objResult);
                    if (m_Result.code == "S001")
                    {
                        m_Result.TransactionId = APIResult.transactionId;
                        resultClass.objResult = m_Result;
                        string forceTryForExam = "";
                        if (ReqClass.forceTryForExam != null)
                        {
                            forceTryForExam = ReqClass.forceTryForExam;
                        }
                        _KFData.UpdByRequestforExam(Form_No, forceTryForExam, m_REUser, ReqClass.comment, TransactionId);
                    }
                    else
                    {
                        resultClass.ResultCode = "999";
                        resultClass.ResultMsg = m_Result.msg;
                        _Comm.InsertErrorLog("RequestforExam", TransactionId, m_Result.msg);
                    }
                }
                else
                {
                    resultClass.ResultCode = "999";
                    resultClass.ResultMsg = $"{APIResult.ResultMsg}";
                }
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "999";
                resultClass.ResultMsg = $" response: {ex.Message}";
            }
            return Ok(resultClass);
        }

        [Route("RequestSupplement")]
        [HttpPost]
        public async Task<ActionResult<string>> RequestSupplement([FromBody] RequestSupplement ReqClass, string Form_No)
        {
            ResultClass<BaseResult> resultClass = new ();
            string TransactionId = DateTime.Now.ToString("yyyyMMddhhmmssffff");
            if (ReqClass is null)
            {
                throw new ArgumentNullException(nameof(ReqClass));
            }
            try
            {
                DataTable tbReceive = _KFData.GetReceiveByform_no(Form_No);
                string m_RSUser = ReqClass.salesNo;
                if (tbReceive.Rows.Count != 0)
                {
                    foreach (DataRow dr in tbReceive.Rows)
                    {
                        ReqClass.branchNo = _branchNo;
                        ReqClass.dealerNo= _dealerNo;
                        ReqClass.examineNo = dr["ExamineNo"].ToString();
                        ReqClass.source = _source;
                        ReqClass.salesNo = "";
                    }
                    if (ReqClass.attachmentFile != null)
                    {
                        ResultClass<int> m_InsertFile = _KFData.InsertFile(Form_No, "RequestSupplement", m_RSUser, TransactionId, ReqClass.attachmentFile);
                        if (m_InsertFile.ResultCode != "000")
                        {
                            resultClass.ResultCode = "999";
                            resultClass.ResultMsg = $"上傳檔案失敗!!";
                            return Ok(resultClass);
                        }
                    }
                }

                string p_JSON = JsonConvert.SerializeObject(ReqClass);
                ResultClass<string> APIResult = new();
                if (_isCallTESTAPI)
                {
                    APIResult = _KFData.GetTestJsonByAPI("TEST0001", "RequestSupplement", m_RSUser, Form_No, TransactionId);
                }
                else
                {
                    APIResult = await _Comm.CallYuRichAPINew("RequestSupplement", ReqClass.salesNo, Form_No, p_JSON, TransactionId,_httpClient);
                }
                if (APIResult.ResultCode == "000")
                {
                    BaseResult m_Result = JsonConvert.DeserializeObject<BaseResult>(APIResult.objResult);
                    if (m_Result.code == "S001")
                    {
                        m_Result.TransactionId = APIResult.transactionId;
                        resultClass.objResult = m_Result;
                        List<string> m_Comments = new();
                        foreach (supplement m_supplement in ReqClass.supplement)
                        {
                            if (m_supplement.comment != "")
                            {
                                m_Comments.Add(m_supplement.comment);
                            }
                        }
                        _KFData.UpdByRequestSupplement(Form_No, m_RSUser, m_Comments, TransactionId);

                    }
                    else
                    {
                        resultClass.ResultCode = "999";
                        resultClass.ResultMsg = m_Result.msg;
                        _Comm.InsertErrorLog("RequestSupplement", TransactionId, m_Result.msg);
                    }
                }
                else
                {
                    resultClass.ResultCode = "999";
                    resultClass.ResultMsg = $"{APIResult.ResultMsg}";
                }
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "999";
                resultClass.ResultMsg = $" response: {ex.Message}";
            }
            return Ok(resultClass);
        }

        [Route("QueryCaseStatus")]
        [HttpPost]
        public async Task<ActionResult<string>> QueryCaseStatus([FromBody] QueryCaseStatus ReqClass, string Form_No,bool isUpdDB=true)
        {
            string TransactionId = DateTime.Now.ToString("yyyyMMddhhmmssffff");

            if (ReqClass is null)
            {
                throw new ArgumentNullException(nameof(ReqClass));
            }
            //紀錄呼叫API人
            string m_User = ReqClass.salesNo;
            //********注意**********************裕富這邊固定傳空值(有值會查不到...)******************************注意*********
            ReqClass.salesNo = "";
            //********注意**********************裕富這邊固定傳空值(有值會查不到...)******************************注意*********
            ResultClass<Result_QCS> resultClass = new ResultClass<Result_QCS>();
            try
            {
                string p_JSON = JsonConvert.SerializeObject(ReqClass);
                ResultClass<string> APIResult = new();
                if (_isCallTESTAPI)
                {
                    APIResult = _KFData.GetTestJsonByAPI("TEST0001", "QueryCaseStatus", m_User, Form_No, TransactionId, "TESTQC004");
                }
                else
                {
                    APIResult = await _Comm.CallYuRichAPINew("QueryCaseStatus", m_User, Form_No, p_JSON, TransactionId, _httpClient, isUpdDB);
                }
                if (APIResult.ResultCode == "000")
                {
                    Result_QCS m_Result = JsonConvert.DeserializeObject<Result_QCS>(APIResult.objResult);

                    if (m_Result.code == "S001")
                    {
                        resultClass.objResult = m_Result;
                        if (m_Result.TransactionId == "")
                        {
                            m_Result.TransactionId = TransactionId;
                        }
                        if (isUpdDB)
                        {
                            ResultClass<int> resultQCS = _KFData.InsertQCS(Form_No, m_User, m_Result);
                            if (resultQCS.ResultCode == "999")
                            {
                                resultClass.ResultCode = "999";
                                resultClass.ResultMsg = resultQCS.ResultMsg;
                                _KFData.DeltbAPILog(TransactionId);

                            }
                        }
                    }
                    else
                    {
                        resultClass.ResultCode = "999";
                        resultClass.ResultMsg = m_Result.msg;
                    }
                }
                else
                {
                    resultClass.ResultCode = "999";
                    resultClass.ResultMsg = $"{APIResult.ResultMsg}";
                }
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "999";
                resultClass.ResultMsg = $" response: {ex.Message}";
            }
            return Ok(resultClass);
        }

        [Route("reCallout")]
        [HttpPost]
        public async Task<ActionResult<string>> reCallout([FromBody] reCallout ReqClass, string Form_No)
        {
            string TransactionId = DateTime.Now.ToString("yyyyMMddhhmmssffff");

            if (ReqClass is null)
            {
                throw new ArgumentNullException(nameof(ReqClass));
            }

            ResultClass<BaseResult> resultClass = new ResultClass<BaseResult>();
            string m_RCUser = ReqClass.salesNo;
            try
            {
                string p_JSON = JsonConvert.SerializeObject(ReqClass);
                ResultClass<string> APIResult = new();
                if (_isCallTESTAPI)
                {
                    APIResult = _KFData.GetTestJsonByAPI("TEST0001", "reCallout", m_RCUser, Form_No, TransactionId);
                }
                else
                {
                    APIResult = await _Comm.CallYuRichAPINew("reCallout", ReqClass.salesNo, Form_No, p_JSON, TransactionId, _httpClient);
                }

                if (APIResult.ResultCode == "000")
                {
                    BaseResult m_Result = JsonConvert.DeserializeObject<BaseResult>(APIResult.objResult);
                    if (m_Result.code == "S001")
                    {
                        resultClass.objResult = m_Result;
                    }
                    else
                    {
                        resultClass.ResultCode = "999";
                        resultClass.ResultMsg = m_Result.msg;
                    }
                }
                else
                {
                    resultClass.ResultCode = "999";
                    resultClass.ResultMsg = $"{APIResult.ResultMsg}";
                }
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "999";
                resultClass.ResultMsg = $" response: {ex.Message}";
            }






            return Ok(resultClass);
        }

        [Route("PutFileToExamiePath")]
        [HttpPost]
        public async Task<ActionResult<string>> PutFileToExamiePath([FromBody] PutFileToExamiePath ReqClass, string Form_No)
        {
            string TransactionId = DateTime.Now.ToString("yyyyMMddhhmmssffff");

            ResultClass<BaseResult> resultClass = new();
            DataTable tbReceive = _KFData.GetReceiveByform_no(Form_No);
            string m_PTUser = ReqClass.salesNo;
            if (tbReceive.Rows.Count != 0)
            {
                foreach (DataRow dr in tbReceive.Rows)
                {
                    ReqClass.branchNo = _branchNo;
                    ReqClass.dealerNo = _dealerNo;
                    ReqClass.examineNo = dr["ExamineNo"].ToString();
                    ReqClass.source = _source;
                    ReqClass.salesNo = "";
                }
                if (ReqClass.attachmentFile != null)
                {
                    ResultClass<int> m_InsertFile = _KFData.InsertFile(Form_No, "PutFileToExamiePath", m_PTUser, TransactionId,ReqClass.attachmentFile);
                    if (m_InsertFile.ResultCode != "000")
                    {
                        resultClass.ResultCode = "999";
                        resultClass.ResultMsg = $"上傳檔案失敗!!";
                        return Ok(resultClass);
                    }
                }
            }

            if (ReqClass is null)
            {
                throw new ArgumentNullException(nameof(ReqClass));
            }
            try
            {
                string p_JSON = JsonConvert.SerializeObject(ReqClass);
                ResultClass<string> APIResult = new();
                if (_isCallTESTAPI)
                {
                    APIResult = _KFData.GetTestJsonByAPI("TEST0001", "PutFileToExamiePath", m_PTUser, Form_No, TransactionId);
                }
                else
                {
                    APIResult = await _Comm.CallYuRichAPINew("RequestPayment", ReqClass.salesNo, Form_No, p_JSON, TransactionId, _httpClient);
                }

                if (APIResult.ResultCode == "000")
                {
                    BaseResult m_Result = JsonConvert.DeserializeObject<BaseResult>(APIResult.objResult);
                    if (m_Result.code == "S001")
                    {
                        resultClass.objResult = m_Result;
                       
                    }
                    else
                    {
                        resultClass.ResultCode = "999";
                        resultClass.ResultMsg = m_Result.msg;
                    }
                }
                else
                {
                    resultClass.ResultCode = "999";
                    resultClass.ResultMsg = $"{APIResult.ResultMsg}";
                }
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "999";
                resultClass.ResultMsg = $" response: {ex.Message}";
            }
            return Ok(resultClass);
        }

        #endregion

        

        #region  對外API

        [Route("NotifyAppropriation")]
        [HttpPost]
        public Task<ActionResult<string>> NotifyAppropriation([FromBody] YuRichAPI_Class ReqYRClass)
        {
            string TransactionId = "";
            string m_Form_No = "";
            BaseResult resultClass = new();
            if (string.IsNullOrEmpty(ReqYRClass.transactionId))
            {
                TransactionId = DateTime.Now.ToString("yyyyMMddhhmmssffff");
            }
            else
            {
                TransactionId = ReqYRClass.transactionId;
            }
            resultClass.TransactionId = TransactionId;
          
            ResultClass<CheckNotifyResult> m_NotifyResult = _Comm.CheckReqYRClass(ReqYRClass);

            if (m_NotifyResult.ResultCode == "F001")
            {
                resultClass.code = "F001";
                resultClass.msg = m_NotifyResult.ResultMsg;
            }
            else
            {
                try
                {
                    CheckNotifyResult m_CheckNotifyResult = m_NotifyResult.objResult;
                    m_Form_No = m_CheckNotifyResult.from_no;
                    string m_DecJSON = m_CheckNotifyResult.DecJSON;
                    NotifyAppropriation ReqClass = JsonConvert.DeserializeObject<NotifyAppropriation>(m_DecJSON);

                    Result_QA m_Result_QA = new();
                    m_Result_QA.TransactionId = TransactionId;
                    m_Result_QA.code = "000";
                    Appropriations[] m_arrAppropriations = new Appropriations[]
                    {
                        new Appropriations
                        {
                            appropriationAmt = ReqClass.appropriationAmt,
                            appropriationDate = ReqClass.appropriationDate,
                            repayKindName = ReqClass.repayKindName,
                            examineNo = ReqClass.examineNo,
                            status = ReqClass.status
                        }
                    };
                    m_Result_QA.appropriations = m_arrAppropriations;
                    ResultClass<int> resultClassQa = _KFData.InsertQueryAppropriation(m_Form_No, "YuRich", m_Result_QA);
                    if (resultClassQa.ResultCode == "000")
                    {
                        resultClass.code = "S001";
                        resultClass.msg = "成功";
                    }
                    else
                    {
                        resultClass.code = "F001";
                        resultClass.msg = "更新失敗!!" + resultClassQa.ResultMsg;
                    }
                }
                catch (Exception ex)
                {
                    resultClass.code = "F001";
                    resultClass.msg = $" Error: {ex.Message}";
                }
            }

            string ipAddress = "查無IP";

            try
            {
                ipAddress = HttpContext.Connection.RemoteIpAddress.ToString();
            }
            catch { }


            _Comm.NotifyLog("NotifyAppropriation", ReqYRClass.transactionId, ipAddress, m_Form_No, JsonConvert.SerializeObject(ReqYRClass), JsonConvert.SerializeObject(resultClass), resultClass.code);



            return Task.FromResult<ActionResult<string>>(Ok(resultClass));
        }

        [Route("NotifyCaseStatus")]
        [HttpPost]
        public async Task<ActionResult<string>> NotifyCaseStatus([FromBody] YuRichAPI_Class ReqYRClass)
        {
            string m_Form_No = "";
            BaseResult resultClass = new();
            resultClass.code = "S001";
            string TransactionId = "";
            if (string.IsNullOrEmpty(ReqYRClass.transactionId))
            {
                TransactionId = DateTime.Now.ToString("yyyyMMddhhmmssffff");
            }
            else
            {
                TransactionId = ReqYRClass.transactionId;
            }
            resultClass.TransactionId = TransactionId;



            ResultClass<CheckNotifyResult> m_NotifyResult = _Comm.CheckReqYRClass( ReqYRClass);

            if (m_NotifyResult.ResultCode == "F001")
            {
                resultClass.code = "F001";
                resultClass.msg = m_NotifyResult.ResultMsg;
            }
            else
            {
                CheckNotifyResult m_CheckNotifyResult = m_NotifyResult.objResult;
                m_Form_No = m_CheckNotifyResult.from_no;
                string m_DecJSON = m_CheckNotifyResult.DecJSON  ;
                NotifyCaseStatus ReqClass = JsonConvert.DeserializeObject<NotifyCaseStatus>(m_DecJSON);
                try
                {
                    QueryCaseStatus m_QueryCaseStatus = new();
                    m_QueryCaseStatus.dealerNo = ReqClass.dealerNo;
                    m_QueryCaseStatus.branchNo = ReqClass.branchNo;
                    m_QueryCaseStatus.salesNo = "";
                    m_QueryCaseStatus.examineNo = ReqClass.examineNo;
                    m_QueryCaseStatus.source = _source;
                    string p_JSON = JsonConvert.SerializeObject(m_QueryCaseStatus);
                    ResultClass<string> APIResult = new();
                    if (_isCallTESTAPI)
                    {
                        APIResult = _KFData.GetTestJsonByAPI("TEST0001", "QueryCaseStatus", "YuRich", m_Form_No, TransactionId, "TESTQC001");
                    }
                    else
                    {
                        APIResult = await _Comm.CallYuRichAPINew("QueryCaseStatus", "YuRich", m_Form_No, p_JSON, TransactionId, _httpClient);
                    }

                    if (APIResult.ResultCode == "000")
                    {
                        Result_QCS m_Result = JsonConvert.DeserializeObject<Result_QCS>(APIResult.objResult);
                        Boolean isUPD = true;

                        if (m_Result.code == "S001")
                        {
                            _KFData.InsertQCS(m_Form_No, "YuRich", m_Result);
                            _KFData.UpdReceiveStatus(m_Form_No, m_Result);
                            resultClass.code = "S001";
                            resultClass.msg = "成功";
                        }
                        else
                        {
                            resultClass.code = "F001";
                            resultClass.msg = m_Result.msg;
                        }
                    }
                    else
                    {
                        resultClass.code = "F001";
                        resultClass.msg = $"{APIResult.ResultMsg}";
                    }
                }
                catch (Exception ex)
                {
                    resultClass.code = "F001";
                    resultClass.msg = $" Error: {ex.Message}";
                }
            }

            string ipAddress = "查無IP";

            try {
                ipAddress = HttpContext.Connection.RemoteIpAddress.ToString();
            }
            catch { }   
            

            _Comm.NotifyLog("NotifyCaseStatus",  ReqYRClass.transactionId, ipAddress, m_Form_No, JsonConvert.SerializeObject(ReqYRClass), JsonConvert.SerializeObject(resultClass), resultClass.code);



            return Ok(resultClass);
        }


        #endregion




    }
}