﻿
using System.Data;
using Microsoft.Data.SqlClient;
using System.IO.Compression;
using System.Reflection;
using System.Text;
using System.Transactions;
using KF_WebAPI.BaseClass;
using Newtonsoft.Json;
using System.Net.Http.Headers;
using Azure.Core;
using KF_WebAPI.BaseClass.Max104;
using KF_WebAPI.DataLogic;
using Microsoft.AspNetCore.Http;
using LoanSky.Model;
using LoanSky;
using System.Threading.Channels;
using Grpc.Net.Client;
using ProtoBuf.Grpc.Client;
using LoanSky.Model;
using KF_WebAPI.BaseClass.LoanSky;
using Microsoft.AspNetCore.DataProtection.KeyManagement;

namespace KF_WebAPI.FunctionHandler
{
    public class Common
    {
        private AesEncryption _AE = new();
        private ADOData _ADO = new();
        /// <summary>
        /// 定義裕富API的正式環境or測試環境
        /// </summary>
        public bool isYRTestAPI = false;
        /// <summary>
        /// 是否呼叫測試API,true:用資料庫模擬API,false:呼叫裕富API;
        /// </summary>
        public bool isCallTESTAPI = false;

        private readonly string _dealerNo = "MM09";
        private readonly string _source = "52611690";
        private readonly string _version = "2.0";


        public string Get104API_URL()
        {
            string m_104API_URL = "";


            m_104API_URL = "http://192.168.1.239:3001/api/";

            return m_104API_URL;
        }




        public async Task<ResultClass<string>> Call_104API(string p_USER, string p_APIName, string p_JSON, string SESSION_KEY, HttpClient p_HttpClient, string ACCESS_TOKEN = "")
        {
            ResultClass<string> resultClass = new();

            var m_CallTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
            try
            {
                using (var httpClient = p_HttpClient)
                {
                    var uri = new Uri(Get104API_URL() + p_APIName);
                    var request = new HttpRequestMessage(HttpMethod.Post, uri);
                    var content = new StringContent(p_JSON, Encoding.UTF8, "application/json");

                    if (ACCESS_TOKEN != "")
                    {
                        request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("*/*"));
                        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", ACCESS_TOKEN);

                    }

                    request.Content = content;
                    var response = await httpClient.SendAsync(request);
                    // ResultClass<string> resultCheckStatus = CheckStatusCode104(p_APIName,response);
                    var AsyncResult = await response.Content.ReadAsStringAsync();
                    resultClass.ResultCode = "200";
                    resultClass.objResult = (AsyncResult);
                    /*if (resultCheckStatus.ResultCode == "000")
                    {
                        var AsyncResult = await response.Content.ReadAsStringAsync();
                        resultClass.ResultCode = "000";
                        resultClass.objResult = (AsyncResult);

                    }
                    else
                    {
                        resultClass = resultCheckStatus;
                        resultClass.objResult = (p_JSON);
                    }*/
                }
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "999";
                resultClass.ResultMsg = "API Error:" + ex.Message;
            }
            try
            {
                ResultClass_104<object> resultClass_log = new ResultClass_104<object>();

                resultClass_log = JsonConvert.DeserializeObject<ResultClass_104<object>>(resultClass.objResult);
                if (resultClass_log is not null)
                {
                    resultClass.ResultCode = resultClass_log.code;
                    resultClass.ResultMsg = resultClass_log.msg;

                }
                else
                {
                    resultClass.ResultCode = "000";
                    resultClass.ResultMsg = "無資料";
                }

                if (ACCESS_TOKEN != "")
                {
                    External_API_Log _External_API_Log = new External_API_Log();
                    _External_API_Log.API_CODE = "104MAX";
                    _External_API_Log.API_NAME = p_APIName;
                    _External_API_Log.API_KEY = SESSION_KEY;
                    _External_API_Log.ACCESS_TOKEN = ACCESS_TOKEN;
                    _External_API_Log.PARAM_JSON = p_JSON;
                    _External_API_Log.RESULT_CODE = resultClass.ResultCode;
                    _External_API_Log.RESULT_MSG = resultClass.ResultMsg;
                    _External_API_Log.Add_User = p_USER;

                    AE_HR m_AE_HR = new AE_HR();
                    m_AE_HR.InsertExternal_API_Log(_External_API_Log);
                }

            }
            catch
            {

            }

            return resultClass;
        }


        /// <summary>
        /// 裕富APIURL
        /// </summary>
        /// <returns></returns>
        public string GetYuRichAPI_URL()
        {
            string m_YuRichAPI_URL = "";
            if (isYRTestAPI)
            {  //裕富測試
                m_YuRichAPI_URL = "https://egateway.tac.com.tw/uat/api/yrc/agent/";
            }
            else
            { //裕富正式
                m_YuRichAPI_URL = "https://egateway.tac.com.tw/production/api/yrc/agent/";

            }
            return m_YuRichAPI_URL;
        }


        public string CheckInt(string p_Param)
        {
            string Result = (p_Param == null) ? "0" : p_Param;

            return Result;
        }

        /// <summary>
        /// 字串處理NULL or Empty
        /// </summary>
        /// <param name="p_Param"></param>
        /// <returns></returns>
        public string CheckString(string p_Param)
        {
            string Result = (p_Param == null) ? "" : p_Param;

            return Result;
        }

        /// <summary>
        /// 字串解析
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="propNa"></param>
        /// <param name="p_JsonVal"></param>
        /// <returns></returns>
        public string CheckJsonType<T>(string propNa, T p_JsonVal)
        {
            string m_Result = "";

            switch (propNa)
            {
                case "mobilePhone":

                    if (p_JsonVal is T[])
                    {
                        T[] inputArray = p_JsonVal as T[];
                        foreach (var obj in inputArray)
                        {
                            m_Result += obj.GetType().GetProperty("number").GetValue(obj, null);
                        }
                    }
                    break;
                case "calloutResult":
                    foreach (var prop in p_JsonVal.GetType().GetProperties())
                    {
                        if (prop.GetValue(p_JsonVal) is not null)
                        {
                            m_Result += prop.GetValue(p_JsonVal).ToString();
                        }
                    }

                    break;
                default:
                    if (p_JsonVal.GetType().FullName == "System.String")
                    {
                        m_Result = p_JsonVal.ToString();
                    }
                    break;
            }

            return m_Result;
        }

        /// <summary>
        /// 裕富參數處理
        /// </summary>
        /// <param name="p_encryptEnterCase">加密字串</param>
        /// <param name="transactionId">交易代號</param>
        /// <returns></returns>
        public YuRichAPI_Class SetYuRichAPI_Class(string p_encryptEnterCase, string transactionId)
        {
            YuRichAPI_Class m_YuRichAPI_Class = new();
            m_YuRichAPI_Class.dealerNo = _dealerNo;
            m_YuRichAPI_Class.source = _source;
            m_YuRichAPI_Class.transactionId = transactionId;
            m_YuRichAPI_Class.encryptEnterCase = p_encryptEnterCase;
            m_YuRichAPI_Class.version = _version;
            return m_YuRichAPI_Class;
        }

        /// <summary>
        /// 外部呼叫紀錄
        /// </summary>
        /// <param name="API_Name"></param>
        /// <param name="TransactionId"></param>
        /// <param name="IP"></param>
        /// <param name="Form_No"></param>
        /// <param name="m_RequeJSON"></param>
        /// <param name="ResultJSON"></param>
        /// <param name="StatusCode"></param>
        public void NotifyLog(string API_Name, string TransactionId, string IP, string Form_No, string m_RequeJSON, string ResultJSON, string StatusCode)
        {
            var m_CallTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
            var m_APILog = new APILog
            {
                API_Name = API_Name,
                TransactionId = TransactionId,
                Form_No = Form_No,
                ParamJSON = m_RequeJSON,
                ResultJSON = ResultJSON,
                CallTime = m_CallTime,
                CallUser = IP,
                StatusCode = StatusCode
            };

            InsertDataByClass("tbAPILog", m_APILog);
        }


        public class CheckNotifyResult
        {
            /// <summary>
            /// 1.表單號碼
            /// </summary>
            public String from_no { get; set; }

            /// <summary>
            /// 2.解密後的JSON 
            /// </summary>
            public String DecJSON { get; set; }


        }


        public class chkYuRichBase
        {
            /// <summary>
            /// 1.通路商編號
            /// </summary>
            public String? dealerNo { get; set; }

            /// <summary>
            /// 統編
            /// </summary>
            public String? source { get; set; }
        }


        public ResultClass<CheckNotifyResult> CheckReqYRClass(YuRichAPI_Class ReqYRClass)
        {
            CheckNotifyResult m_CheckNotifyResult = new();
            ResultClass<CheckNotifyResult> resultClass = new();
            if (ReqYRClass is null)
            {
                resultClass.ResultCode = "F001";
                resultClass.ResultMsg = "查無輸入資料";
            }
            else
            {
                //解密encryptEnterCase   h
                ResultClass<string> m_DecResult = _AE.DecryptAES256(ReqYRClass.encryptEnterCase);
                if (m_DecResult.ResultCode != "000")
                {
                    resultClass.ResultCode = "F001";
                    resultClass.ResultMsg = m_DecResult.ResultMsg.ToString();
                }
                else
                {
                    m_CheckNotifyResult.DecJSON = m_DecResult.objResult;
                    //解密成功轉成 NotifyCaseStatus
                    string m_DecJSON = m_DecResult.objResult;
                    string examineNo = "";
                    YuRichBase ReqClass = JsonConvert.DeserializeObject<YuRichBase>(m_DecJSON);
                    string m_Error = "";
                    if (CheckString(ReqYRClass.dealerNo) != "MM09")
                    {
                        m_Error = "dealerNo 錯誤!!";
                    }
                    if (CheckString(ReqYRClass.source) != "52611690")
                    {
                        m_Error += "source 錯誤!!";
                    }

                    if (CheckString(ReqClass.branchNo) == "")
                    {
                        m_Error += "參數 branchNo  錯誤!!";
                    }
                    if (CheckString(ReqClass.dealerNo) == "")
                    {
                        m_Error += "參數 dealerNo  錯誤!!";
                    }
                    if (CheckString(ReqClass.examineNo) == "")
                    {
                        m_Error += "參數 examineNo  錯誤!!";
                    }

                    if (m_Error == "")
                    {
                        examineNo = ReqClass.examineNo;
                        //查詢 form_no
                        List<SqlParameter> Params = new List<SqlParameter>()
                        {
                            new SqlParameter() {ParameterName = "@ExamineNo", SqlDbType = SqlDbType.VarChar, Value= examineNo}
                        };

                        DataTable m_RtnDT = _ADO.ExecuteQuery("SELECT form_no FROM tbReceive WHERE ExamineNo = @ExamineNo  ", Params);
                        string m_Form_No = "";
                        foreach (DataRow dr in m_RtnDT.Rows)
                        {
                            m_Form_No = dr["form_no"].ToString();
                        }

                        if (m_Form_No == "")
                        {
                            resultClass.ResultCode = "F001";
                            resultClass.ResultMsg = "審件編號:" + ReqClass.examineNo + ";不存在";
                        }
                        else
                        {
                            m_CheckNotifyResult.from_no = m_Form_No;
                            resultClass.ResultCode = "S001";
                            resultClass.objResult = m_CheckNotifyResult;
                        }
                    }
                    else
                    {
                        resultClass.ResultCode = "F001";
                        resultClass.ResultMsg = m_Error;
                    }
                }
            }

            return resultClass;
        }




        public Boolean CheckClass<T>(T[] p_compareNews, T[] p_compareOlds)
        {
            Boolean isSame = true;
            try
            {
                if (p_compareOlds == null && p_compareNews == null)
                {
                    return true;
                }
                if (p_compareOlds == null && p_compareNews != null)
                {
                    return false;
                }
                if (p_compareNews.Length != p_compareOlds.Length)
                {
                    return false;
                }


                Int32 m_Count = 0;
                foreach (var compareNew in p_compareNews)
                {
                    Type OldProp = p_compareOlds[m_Count].GetType();
                    foreach (var prop in compareNew.GetType().GetProperties())
                    {
                        if (prop.Name != "attachmentFile" && prop.Name != "FileInfo" && prop.Name != "Action")
                        {
                            PropertyInfo OldPropInfo = OldProp.GetProperty(prop.Name);
                            string OldPropValue = CheckString((string)OldPropInfo.GetValue(p_compareOlds[m_Count]));

                            string value = "";
                            if (prop.GetValue(compareNew) is not null)
                            {
                                value = prop.GetValue(compareNew).ToString();
                            }
                            if (OldPropValue != value)
                            {
                                return false;
                            }
                        }
                    }
                    m_Count++;
                }

            }
            catch
            {
                isSame = false;
            }

            return isSame;
        }



        /// <summary>
        /// 組依傳入的CLASSSQL語法
        /// </summary>
        /// <typeparam name="T">動態物件</typeparam>
        /// <param name="p_tbName">TableName</param>
        /// <param name="p_ExistClassCol">不存在Class的欄位</param>
        /// <param name="m_ExcClassColumn">不需處理的Class屬性</param>
        /// <param name="p_Class">動態物件</param>
        /// <returns></returns>
        public string GetInsertColumnByClass<T>(string p_tbName, string[] p_ExistClassCol, string[] m_ExcClassColumn, T p_Class)
        {
            string m_SQL = "";
            if (p_Class is not null)
            {
                string m_Columns = "";
                string m_Values = "";

                foreach (var ColumnNa in p_ExistClassCol)
                {
                    if (m_Columns == "")
                    {
                        m_Columns += "[" + ColumnNa + "]";
                        m_Values += "@" + ColumnNa;
                    }
                    else
                    {
                        m_Columns += ",[" + ColumnNa + "]";
                        m_Values += ",@" + ColumnNa;
                    }
                }

                foreach (var prop in p_Class.GetType().GetProperties())
                {

                    if (Array.IndexOf(m_ExcClassColumn, prop.Name) == -1)
                    {
                        if (m_Columns == "")
                        {
                            m_Columns += "[" + prop.Name + "]";
                            m_Values += "@" + prop.Name;
                        }
                        else
                        {
                            m_Columns += ",[" + prop.Name + "]";
                            m_Values += ",@" + prop.Name;
                        }
                    }
                }
                m_SQL = " INSERT INTO " + p_tbName + " (" + m_Columns + ")  ";
                m_SQL += "  VALUES (" + m_Values + ")  ";
            }
            return m_SQL;

        }

        /// <summary>
        /// 呼叫API錯誤 Log
        /// </summary>
        /// <param name="p_API_Name"></param>
        /// <param name="p_TransactionId"></param>
        /// <param name="p_ErrMSG"></param>
        public void InsertErrorLog(string p_API_Name, string p_TransactionId, string p_ErrMSG)
        {
            var m_APIErrLog = new APIErrorLog
            {
                API_Name = p_API_Name,
                TransactionId = p_TransactionId,
                ErrMSG = p_ErrMSG
            };

            InsertDataByClass("tbAPIErrorLog", m_APIErrLog);
        }



        /// <summary>
        /// 寫入APILog檔
        /// </summary>
        /// <param name="p_API_Name"></param>
        /// <param name="p_TransactionId"></param>
        /// <param name="p_ParamJSON"></param>
        /// <param name="p_CallTime"></param>
        /// <param name="p_ResultClass"></param>
        public void InsertAPILog<T>(string p_API_Name, string p_TransactionId, string p_Form_No, string p_ParamJSON, string p_CallTime, string p_CallUser, ResultClass<T> p_ResultClass)
        {
            var m_APILog = new APILog
            {
                API_Name = p_API_Name,
                TransactionId = p_TransactionId,
                Form_No = p_Form_No,
                // ParamJSON = p_ParamJSON,
                ResultJSON = p_ResultClass.objResult.ToString(),
                CallTime = p_CallTime,
                CallUser = p_CallUser,
                StatusCode = p_ResultClass.ResultCode
            };

            InsertDataByClass("tbAPILog", m_APILog);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="p_tbName"></param>
        /// <param name="p_BaseClass"></param>
        /// <returns></returns>
        public ResultClass<int> InsertDataByClass<T>(string p_tbName, T p_BaseClass)
        {
            ResultClass<int> resultClass = new();
            int m_Execut = 0;
            using SqlConnection conn = new SqlConnection(_ADO.ConnStr);
            // 開啟資料庫連線
            conn.Open();
            SqlTransaction transaction = conn.BeginTransaction();
            try
            {
                if (p_BaseClass is not null)
                {
                    string m_Columns = "";
                    string m_Values = "";

                    m_Columns = "";
                    m_Values = "";
                    foreach (var prop in p_BaseClass.GetType().GetProperties())
                    {
                        if (prop.Name != "attachmentFile" && prop.Name != "FileInfo" && prop.Name != "Action" && prop.Name != "add_date")
                        {
                            if (m_Columns == "")
                            {
                                m_Columns += prop.Name;
                                m_Values += "@" + prop.Name;

                            }
                            else
                            {
                                m_Columns += "," + prop.Name;
                                m_Values += ",@" + prop.Name;
                            }
                        }
                    }
                    // 建立 SQL 命令
                    string m_SQL = " INSERT INTO " + p_tbName + " (" + m_Columns + ")  ";
                    m_SQL += "  VALUES (" + m_Values + ")  ";
                    using SqlCommand command = new SqlCommand(m_SQL, conn, transaction);
                    // 設定參數
                    foreach (var prop in p_BaseClass.GetType().GetProperties())
                    {
                        if (prop.Name != "attachmentFile" && prop.Name != "FileInfo" && prop.Name != "Action")
                        {
                            string value = "";
                            if (prop.GetValue(p_BaseClass) is not null)
                            {
                                value = prop.GetValue(p_BaseClass).ToString();
                            }
                            command.Parameters.AddWithValue("@" + prop.Name, value);
                        }
                    }
                    // 執行 SQL 命令
                    m_Execut += command.ExecuteNonQuery();
                    transaction.Commit();
                    // 關閉資料庫連線
                    conn.Close();
                    resultClass.ResultCode = "000";
                    resultClass.ResultMsg = "";
                    resultClass.objResult = m_Execut;
                }
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                resultClass.ResultCode = "999";
                resultClass.ResultMsg = ex.Message;
                resultClass.objResult = 0;
            }
            return resultClass;
        }

        public ResultClass<int> UpdateDataByClass<T>(string p_tbName, string Key, T p_BaseClass)
        {
            ResultClass<int> resultClass = new();
            int m_Execut = 0;
            using SqlConnection conn = new SqlConnection(_ADO.ConnStr);
            // 開啟資料庫連線
            conn.Open();
            SqlTransaction transaction = conn.BeginTransaction();
            try
            {
                if (p_BaseClass is not null)
                {
                    string m_SQL = "";

                    foreach (var prop in p_BaseClass.GetType().GetProperties())
                    {
                        if (prop.Name != "attachmentFile" && prop.Name != "FileInfo" && prop.Name != "Action" && prop.Name != "add_date" && prop.Name != "add_user" && prop.Name != Key)
                        {
                            if (m_SQL == "")
                            {
                                m_SQL = " Update " + p_tbName + " set " + prop.Name + "=@" + prop.Name;
                            }
                            else
                            {
                                m_SQL += ", " + prop.Name + "=@" + prop.Name;
                            }
                        }
                    }

                    m_SQL += " WHERE " + Key + "=@" + Key;

                    // 建立 SQL 命令

                    using SqlCommand command = new SqlCommand(m_SQL, conn, transaction);
                    // 設定參數
                    foreach (var prop in p_BaseClass.GetType().GetProperties())
                    {
                        if (prop.Name != "attachmentFile" && prop.Name != "FileInfo" && prop.Name != "Action" && prop.Name != "add_date" && prop.Name != "add_user")
                        {
                            string value = "";
                            if (prop.GetValue(p_BaseClass) is not null)
                            {
                                value = prop.GetValue(p_BaseClass).ToString();
                            }
                            command.Parameters.AddWithValue("@" + prop.Name, value);
                        }
                    }
                    // 執行 SQL 命令
                    m_Execut += command.ExecuteNonQuery();
                    transaction.Commit();
                    // 關閉資料庫連線
                    conn.Close();
                    resultClass.ResultCode = "000";
                    resultClass.ResultMsg = "";
                    resultClass.objResult = m_Execut;
                }
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                resultClass.ResultCode = "999";
                resultClass.ResultMsg = ex.Message;
                resultClass.objResult = 0;
            }
            return resultClass;
        }



        public async Task<ResultClass<string>> CallYuRichAPINew(string p_APIName, string p_CallUser, string p_Form_No, string p_JSON, string p_TransactionId, HttpClient p_HttpClient, bool isUpdDB = true)
        {
            ResultClass<string> resultClass = new();
            if (!string.IsNullOrEmpty(p_JSON))
            {
                string m_AE_Json = _AE.EncryptAES256(p_JSON);
                YuRichAPI_Class m_YuRichAPI_Class = SetYuRichAPI_Class(m_AE_Json, p_TransactionId);
                var EncJsonString = JsonConvert.SerializeObject(m_YuRichAPI_Class);
                var m_CallTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                try
                {
                    using (var httpClient = p_HttpClient)
                    {
                        var uri = new Uri(GetYuRichAPI_URL() + p_APIName);
                        var request = new HttpRequestMessage(HttpMethod.Post, uri);
                        var content = new StringContent(EncJsonString, Encoding.UTF8, "application/json");
                        request.Content = content;
                        var response = await httpClient.SendAsync(request);
                        ResultClass<string> resultCheckStatus = CheckStatusCodeNew(response);
                        if (resultCheckStatus.ResultCode == "000")
                        {
                            var AsyncResult = await response.Content.ReadAsStringAsync();
                            resultClass.ResultCode = "000";
                            resultClass.objResult = (AsyncResult);
                            resultClass.transactionId = p_TransactionId;
                        }
                        else
                        {
                            resultClass = resultCheckStatus;
                        }
                    }
                }
                catch (Exception ex)
                {
                    resultClass.ResultCode = "999";
                    resultClass.ResultMsg = "API Error:" + ex.Message;
                }
                if (isUpdDB)
                {
                    InsertAPILog(p_APIName, resultClass.transactionId, p_Form_No, EncJsonString, m_CallTime, p_CallUser, resultClass);
                }
            }
            else
            {
                resultClass.ResultCode = "999";
                resultClass.ResultMsg = "傳入參數為空!!";
            }
            return resultClass;
        }

        public ResultClass<string> CheckStatusCodeNew(HttpResponseMessage p_response)
        {
            ResultClass<string> resultClass = new() { ResultCode = "000", ResultMsg = "", objResult = "" };

            if (!p_response.IsSuccessStatusCode)
            {
                resultClass.objResult = "";
                resultClass.ResultCode = "999";
                if (p_response.StatusCode == System.Net.HttpStatusCode.NotFound)
                {
                    resultClass.ResultMsg = "API endpoint not found";
                }
                else if (p_response.StatusCode == System.Net.HttpStatusCode.Unauthorized)
                {
                    resultClass.ResultMsg = "API endpoint requires authentication";
                }
                else
                {
                    resultClass.ResultMsg = $"API call failed with status code {p_response.StatusCode}";
                }
            }
            return resultClass;
        }


        public ResultClass<string> CheckStatusCode104(string p_APIName, HttpResponseMessage p_response)
        {
            ResultClass<string> resultClass = new() { ResultCode = "000", ResultMsg = "", objResult = "" };

            if (!p_response.IsSuccessStatusCode)
            {
                resultClass.objResult = "";
                resultClass.ResultCode = "999";
                if (p_response.StatusCode == System.Net.HttpStatusCode.NotFound)
                {
                    resultClass.ResultMsg = "API endpoint not found";
                }
                else if (p_response.StatusCode == System.Net.HttpStatusCode.Unauthorized)
                {
                    resultClass.ResultMsg = "API Unauthorized ";
                }
                else if (p_response.StatusCode.ToString() == "490")
                {
                    resultClass.ResultMsg = "API 邏輯錯誤 ";
                }
                else
                {
                    switch (p_APIName)
                    {
                        case "auth/signIn":
                        case "auth/signOut":
                            switch (p_response.StatusCode.ToString())
                            {
                                case "401":
                                    resultClass.ResultMsg = "認證失敗 (請更新 accessToken 再試一次)";
                                    break;
                                case "490":
                                    resultClass.ResultMsg = "登入失敗";
                                    break;
                                case "440":
                                    resultClass.ResultMsg = "參數錯誤 (請檢查 parameters 或 request body 的欄位、格式是否完整及正確)";
                                    break;
                                case "500":
                                    resultClass.ResultMsg = "系統異常";
                                    break;
                            }
                            break;
                    }

                }
            }
            return resultClass;
        }


        public ResultClass<bool> CheckStatusCode(HttpResponseMessage p_response)
        {
            ResultClass<bool> resultClass = new() { ResultCode = "000", ResultMsg = "", objResult = false };

            if (!p_response.IsSuccessStatusCode)
            {
                resultClass.objResult = true;
                resultClass.ResultCode = "999";
                if (p_response.StatusCode == System.Net.HttpStatusCode.NotFound)
                {
                    resultClass.ResultMsg = "API endpoint not found";
                }
                else if (p_response.StatusCode == System.Net.HttpStatusCode.Unauthorized)
                {
                    resultClass.ResultMsg = "API endpoint requires authentication";
                }
                else
                {
                    resultClass.ResultMsg = $"API call failed with status code {p_response.StatusCode}";
                }
            }
            return resultClass;
        }


        /// <summary>
        /// 壓縮 base64File
        /// </summary>
        /// <param name="base64File"></param>
        /// <returns></returns>
        public string CompressFile(string base64File)
        {
            // 解碼base64編碼的檔案
            byte[] fileBytes = Convert.FromBase64String(base64File);

            // 創建新的MemoryStream對象
            MemoryStream memStream = new MemoryStream();

            // 使用GZipStream壓縮檔案到記憶流中
            using (var gzipStream = new GZipStream(memStream, CompressionMode.Compress, false))
            {
                gzipStream.Write(fileBytes, 0, fileBytes.Length);
            }

            // 轉換記憶流中的壓縮檔案為base64編碼的字符串
            string compressedString = Convert.ToBase64String(memStream.ToArray());

            // 釋放資源
            memStream.Close();

            // 返回base64編碼的壓縮檔案
            return compressedString;
        }
        /// <summary>
        /// 解壓縮 base64File
        /// </summary>
        /// <param name="compressedBase64String"></param>
        /// <returns></returns>
        public string DecompressFile(string compressedBase64String)
        {
            string base64String = "";
            // 解碼base64編碼的壓縮檔案
            byte[] compressedBytes = Convert.FromBase64String(compressedBase64String);

            // 創建新的MemoryStream對象，並將壓縮檔案寫入其中
            MemoryStream compressedStream = new MemoryStream(compressedBytes);

            // 創建新的MemoryStream對象，並使用GZipStream解壓縮檔案
            MemoryStream decompressedStream = new MemoryStream();
            using (GZipStream gzipStream = new GZipStream(compressedStream, CompressionMode.Decompress))
            {
                gzipStream.CopyTo(decompressedStream);
            }

            // 從記憶流中讀取解壓縮後的檔案
            byte[] decompressedBytes = decompressedStream.ToArray();
            base64String = Convert.ToBase64String(decompressedBytes);
            // 釋放資源
            compressedStream.Close();
            decompressedStream.Close();

            // 返回解壓縮後的檔案
            return base64String;
        }

        #region LoanSky 宏邦API
        //private string account = "testCGU"; // 測試用帳號:testCGU
        private OrderRealEstateAdapterRequest ORErequest = new OrderRealEstateAdapterRequest();
        


        public async Task<ResultClass<string>> Call_LoanSkyAPI(string p_USER, string p_APIName, OrderRealEstateRequest request, LoanSky_Req req)
        {
            string apiCode = "LoanSky";
            string apikey = req.HP_id; // JsonConvert.SerializeObject(req);
            FuncHandler _fun = new FuncHandler();
            ResultClass<string> resultClass = new ResultClass<string>();
            ORErequest.ApiKey = req.LoanSkyApiKey;
            
            ORErequest.OrderRealEstates.Add(request);
            GrpcChannel channel = GrpcChannel.ForAddress(req.LoanSkyUrl);
            try
            {
                
                var company = channel.CreateGrpcService<ICompany>();
                var reply = await company.CreateOrderRealEstateAsync(ORErequest);

                resultClass.ResultMsg = JsonConvert.SerializeObject(reply);
                ORErequest.OrderRealEstates[0].BusinessUserName = _fun.toNCR(request.BusinessUserName); // 處理中文亂碼: LoanSky(UTF8) -> API.Log(NCR)
                #region 附件內容“不在”API Log
                //ORErequest.OrderRealEstates[0].Attachments.Clear(); // 清除附件內容
                ORErequest.OrderRealEstates[0].Attachments.ForEach(attachment =>
                {
                    attachment.File = new byte[0]; // 清除附件內容
                    attachment.OrginalFileName = _fun.toNCR(attachment.OrginalFileName); // 處理中文亂碼: LoanSky(UTF8) -> API.Log(NCR)
                });
                #endregion

                resultClass.objResult = JsonConvert.SerializeObject(ORErequest);
                if(reply.Result)
                {
                    resultClass.ResultCode = "200";
                    _fun.ExtAPILogIns(apiCode, p_APIName, apikey, ORErequest.ApiKey,
                        resultClass.objResult, resultClass.ResultCode, resultClass.ResultMsg, p_USER);
                }
                else
                {
                    resultClass.ResultCode = "500";
                    _fun.ExtAPILogIns(apiCode, p_APIName, apikey, ORErequest.ApiKey,
                        resultClass.objResult, resultClass.ResultCode, resultClass.ResultMsg, p_USER);
                }

            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                resultClass.ResultMsg = $" response: {ex.Message}";
                _fun.ExtAPILogIns(apiCode, p_APIName, apikey, ORErequest.ApiKey,
                       JsonConvert.SerializeObject(req), "500", $" response: {ex.Message}", p_USER);
                
            }
            finally
            {
                await channel.ShutdownAsync();
            }
            return resultClass;
        }

        
        #endregion

    }
}
