using KF_WebAPI.BaseClass;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using KF_WebAPI.FunctionHandler;
using KF_WebAPI.DataLogic;
using System.Data;
using Microsoft.AspNetCore.Http;
using System.Data.SqlClient;
using static KF_WebAPI.Controllers.YuRichAPIController;
using System.Transactions;
using static KF_WebAPI.FunctionHandler.Common;

namespace KF_WebAPI.Controllers
{
    public class MyObject
    {
        /// <summary>
        /// 回覆代碼
        /// </summary>
        public string? p_Key { get; set; }

        /// <summary>
        /// 回覆訊息 
        /// </summary>
        public attachmentFile[]? attachmentFiles { get; set; }
    }

    [ApiController]
    public class YuRichAPIController : Controller
    {
        /// <summary>
        /// 裕富API測試模式
        /// </summary>
        public Boolean _isYRAPITest = true;

        private readonly HttpClient _httpClient;
        private readonly string _branchNo = "0001";
        private readonly string _dealerNo = "MM09";
        private readonly string _source = "22";
        private AesEncryption _AE = new();
        private Common _Comm = new();
        private KFDataADO _ADO = new();
        public YuRichAPIController(IHttpClientFactory httpClientFactory)
        {
            _httpClient = httpClientFactory.CreateClient();
        }


        /*
        #region  內部API
        [Route("GetFileBySeq")]
        [HttpPost]
        public async Task<ActionResult<string>> GetFileBySeq(string KeyID, string Type, string Index)
        {
            ResultClass<attachmentFile> resultClass = new();
            try
            {
                resultClass = _ADO.GetFile(KeyID, Type, Index);
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "999";
                resultClass.ResultMsg = ex.Message;
            }
            return Ok(resultClass);
        }



        [Route("InsertReceive")]
        [HttpPost]
        public ActionResult<ResultClass<BaseResult>> InsertReceive([FromBody] objInsertReceive objects, string salesNo)
        {
           
            ResultClass<string> resultClass = new ResultClass<string>();
            try
            {
                ResultClass<string> resultSeq = _ADO.GetSeqNo("KF");
                if (resultSeq is not null && resultSeq.ResultCode == "000") 
                {
                    if (objects._Receive1 != null)
                    {
                        objects._Receive1.Action = "Add";
                        objects._Receive1.Case_Company = "KF";
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
                                ResultClass<int> m_InsertFile = _ADO.InsertFile(resultSeq.objResult, "Receive", salesNo,"", objects._Receive1.attachmentFile);
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
            ResultClass<List<objReceive>> resultClass = _ADO.GetReceive(form_no);
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
                if (_isYRAPITest)
                {
                    APIResult = _ADO.GetTestJsonByAPI("TEST0004", "Receive", salesNo, Form_No, TransactionId);
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
                        _ADO.UpdReceive(Form_No, APIResult.transactionId, m_Result.examineNo);
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

        [Route("QueryAppropriation")]
        [HttpPost]
        public async Task<ActionResult<string>> QueryAppropriation([FromBody] QueryAppropriation ReqClass, string Form_No)
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
                if (_isYRAPITest)
                {
                    APIResult  = _ADO.GetTestJsonByAPI("TEST0004", "QueryAppropriation", m_Uesr, Form_No, TransactionId, "TESTQA_1");
                }
                else
                {
                    APIResult = await _Comm.CallYuRichAPINew("QueryAppropriation", ReqClass.salesNo, Form_No, p_JSON, TransactionId, _httpClient);
                }
                if (APIResult.ResultCode == "000")
                {
                    Result_QA m_Result = JsonConvert.DeserializeObject<Result_QA>(APIResult.objResult);
                    if (m_Result.code == "S001")
                    {
                        resultClass.objResult = m_Result;
                        _ADO.InsertQueryAppropriation(Form_No, m_Uesr, m_Result);
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

        [Route("RequestPayment")]
        [HttpPost]
        public async Task<ActionResult<string>> RequestPayment([FromBody] RequestPayment ReqClass, string Form_No)
        {
            ResultClass<BaseResult> resultClass = new ();
            DataTable tbReceive = _ADO.GetReceiveByform_no(Form_No);
            string TransactionId = DateTime.Now.ToString("yyyyMMddhhmmssffff");
            string m_RPUser = ReqClass.salesNo;
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
                    ResultClass<int> m_InsertFile = _ADO.InsertFile(Form_No, "RequestPayment", m_RPUser, TransactionId, ReqClass.attachmentFile);
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
                if (_isYRAPITest)
                {
                    APIResult = _ADO.GetTestJsonByAPI("TEST0004", "RequestPayment", m_RPUser, Form_No, TransactionId);
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
                        _ADO.UpdByRequestPayment(Form_No, m_RPUser, TransactionId);
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
                DataTable tbReceive = _ADO.GetReceiveByform_no(Form_No);
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
                        ResultClass<int> m_InsertFile = _ADO.InsertFile(Form_No, "RequestforExam", m_REUser, TransactionId,ReqClass.attachmentFile);
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
                if (_isYRAPITest)
                {
                    APIResult = _ADO.GetTestJsonByAPI("TEST0004", "RequestforExam", m_REUser,Form_No, TransactionId, "TESTRE_1");
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
                        _ADO.UpdByRequestforExam(Form_No, forceTryForExam, m_REUser, ReqClass.comment, TransactionId);
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
                DataTable tbReceive = _ADO.GetReceiveByform_no(Form_No);
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
                        ResultClass<int> m_InsertFile = _ADO.InsertFile(Form_No, "RequestSupplement", m_RSUser, TransactionId, ReqClass.attachmentFile);
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
                if (_isYRAPITest)
                {
                    APIResult = _ADO.GetTestJsonByAPI("TEST0004", "RequestSupplement", m_RSUser, Form_No, TransactionId);
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
                        _ADO.UpdByRequestSupplement(Form_No, m_RSUser, m_Comments, TransactionId);

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

        [Route("QueryCaseStatus")]
        [HttpPost]
        public async Task<ActionResult<string>> QueryCaseStatus([FromBody] QueryCaseStatus ReqClass, string Form_No)
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
                if (_isYRAPITest)
                {
                    APIResult = _ADO.GetTestJsonByAPI("TEST0004", "QueryCaseStatus", m_User, Form_No, TransactionId, "TESTQC_2");
                }
                else
                {
                    APIResult = await _Comm.CallYuRichAPINew("QueryCaseStatus", m_User, Form_No, p_JSON, TransactionId, _httpClient);
                }
                if (APIResult.ResultCode == "000")
                {
                    Result_QCS m_Result = JsonConvert.DeserializeObject<Result_QCS>(APIResult.objResult);

                    if (m_Result.code == "S001")
                    {
                        resultClass.objResult = m_Result;
                        _ADO.InsertQCS(Form_No, m_User, m_Result);
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
                if (_isYRAPITest)
                {
                    APIResult = _ADO.GetTestJsonByAPI("TEST0004", "reCallout", m_RCUser, Form_No, TransactionId);
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
            DataTable tbReceive = _ADO.GetReceiveByform_no(Form_No);
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
                    ResultClass<int> m_InsertFile = _ADO.InsertFile(Form_No, "PutFileToExamiePath", m_PTUser, TransactionId,ReqClass.attachmentFile);
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
                if (_isYRAPITest)
                {
                    APIResult = _ADO.GetTestJsonByAPI("TEST0004", "PutFileToExamiePath", m_PTUser, Form_No, TransactionId);
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

        */

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
                    ResultClass<int> resultClassQa = _ADO.InsertQueryAppropriation(m_Form_No, "YuRich", m_Result_QA);
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
                    if (_isYRAPITest)
                    {
                        APIResult = _ADO.GetTestJsonByAPI("TEST0004", "QueryCaseStatus", "YuRich", m_Form_No, TransactionId, "TESTQC1");
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
                            _ADO.InsertQCS(m_Form_No, "YuRich", m_Result);
                            _ADO.UpdReceiveStatus(m_Form_No, m_Result);
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