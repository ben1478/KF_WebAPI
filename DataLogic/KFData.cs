using KF_WebAPI.FunctionHandler;
using System.Data;
using System.Data.SqlClient;
using KF_WebAPI.BaseClass;
using Newtonsoft.Json;
using System.Data.Common;
using System.Transactions;


namespace KF_WebAPI.DataLogic
{
    public class KFData
    {
        ADOData _ADO = new();
        Common _Common = new();

        public Int32 DeltbAPILog(string TransactionId)
        {
            Int32 m_Execut = 0;
            try
            {
                List<SqlParameter> ParamsR = new List<SqlParameter>()
                {
                    new SqlParameter() {ParameterName = "@TransactionId", SqlDbType = SqlDbType.VarChar, Value= TransactionId}
                };

                _ADO.ExecuteNonQuery("  delete tbAPILog where TransactionId=@TransactionId ", ParamsR);
            }
            catch
            {
                throw;
            }
            return m_Execut;
        }


        public DataTable UserVerify(UserLogin p_UserLogin)
        {
            DataTable m_RtnDT = new("Data");
            try
            {
               /* List<SqlParameter> Params = new List<SqlParameter>()
                {
                    new SqlParameter() {ParameterName = "@form_no", SqlDbType = SqlDbType.VarChar, Value= form_no}
                };

                m_RtnDT = _ADO.ExecuteQuery("SELECT * FROM User_M WHERE Username = @Username AND PasswordHash = @PasswordHash  ", Params);
               */
            }
            catch
            {
                throw;
            }
            return m_RtnDT;
        }

        public ResultClass<string> GetTestJsonByAPI(string TestID, string API_Name, string User, string form_no, string TransactionId, string subTestID = "")
        {
            ResultClass<string> m_Result = new();
            m_Result.transactionId = TransactionId;
            List<SqlParameter> Params = new();
            try
            {
                Params = new List<SqlParameter>()
                {
                    new SqlParameter() {ParameterName = "@form_no", SqlDbType = SqlDbType.VarChar, Value= form_no}
                };

                DataTable m_tbReceive = _ADO.ExecuteQuery("select * from tbReceive where  Form_No=@Form_No", Params);

                if (m_tbReceive.Rows.Count != 0)
                {
                    object m_APIResult = new { };
                    Params = new List<SqlParameter>()
                    {
                        new SqlParameter() {ParameterName = "@TestID", SqlDbType = SqlDbType.VarChar, Value= TestID},
                        new SqlParameter() {ParameterName = "@API_Name", SqlDbType = SqlDbType.VarChar, Value= API_Name}
                    };

                    string m_SQL = "select ResultJSON from tbAPILog where PreTransactionId=@TestID   and API_Name=@API_Name  ";
                    if (subTestID != "")
                    {
                        m_SQL += "  and TransactionId='" + subTestID + "' ";
                    }

                    DataTable m_RtnDT = _ADO.ExecuteQuery(m_SQL, Params);
                    foreach (DataRow dr in m_RtnDT.Rows)
                    {
                        m_Result.ResultCode = "000";
                        switch (API_Name)
                        {
                            case "Receive":
                                Result_R m_Result_R = JsonConvert.DeserializeObject<Result_R>(dr["ResultJSON"].ToString());
                                m_Result_R.examineNo = form_no.Replace("KF", "TE").Replace("YL", "TE");
                                m_Result_R.TransactionId = TransactionId;
                                m_Result.objResult = JsonConvert.SerializeObject(m_Result_R);

                                break;
                            case "QueryAppropriation":
                                Result_QA m_Result_QA = JsonConvert.DeserializeObject<Result_QA>(dr["ResultJSON"].ToString());
                                if (m_Result_QA.appropriations != null)
                                {
                                    foreach (Appropriations aps in m_Result_QA.appropriations)
                                    {
                                        aps.examineNo = form_no.Replace("KF", "TE");
                                    }
                                }
                                m_Result_QA.TransactionId = TransactionId;
                                m_Result.objResult = JsonConvert.SerializeObject(m_Result_QA);
                                break;
                            case "QueryCaseStatus":
                                Result_QCS m_Result_QCS = JsonConvert.DeserializeObject<Result_QCS>(dr["ResultJSON"].ToString());
                                m_Result_QCS.TransactionId = TransactionId;
                                m_Result.objResult = JsonConvert.SerializeObject(m_Result_QCS);
                                break;
                            case "RequestPayment":
                                BaseResult m_BaseResult = JsonConvert.DeserializeObject<BaseResult>(dr["ResultJSON"].ToString());
                                m_BaseResult.TransactionId = TransactionId;
                                m_Result.objResult = JsonConvert.SerializeObject(m_BaseResult);
                                break;
                            case "RequestforExam":
                                Result_RE m_Result_RE = JsonConvert.DeserializeObject<Result_RE>(dr["ResultJSON"].ToString());
                                m_Result_RE.TransactionId = TransactionId;
                                m_Result.objResult = JsonConvert.SerializeObject(m_Result_RE);
                                break;

                            case "RequestSupplement":
                                BaseResult m_RSResult = JsonConvert.DeserializeObject<BaseResult>(dr["ResultJSON"].ToString());
                                m_RSResult.TransactionId = TransactionId;
                                m_Result.objResult = JsonConvert.SerializeObject(m_RSResult);
                                break;
                        }
                    }
                    var m_CallTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");

                    ResultClass<string> resultClass = new();
                    resultClass.ResultCode = "000";
                    resultClass.objResult = (m_Result.objResult);
                    resultClass.transactionId = TransactionId;
                    _Common.InsertAPILog(API_Name, TransactionId, form_no, "TEST_JSON", m_CallTime, User, resultClass);
                }
                else
                {
                    m_Result.ResultCode = "999";
                    m_Result.ResultMsg = "查無資料";
                }
                return m_Result;
            }
            catch
            {
                throw;
            }
        }

        public DataTable GetReceiveByform_no(string form_no)
        {
            DataTable m_RtnDT = new("Data");
            try
            {
                List<SqlParameter> Params = new List<SqlParameter>()
                {
                    new SqlParameter() {ParameterName = "@form_no", SqlDbType = SqlDbType.VarChar, Value= form_no}
                };

                m_RtnDT = _ADO.ExecuteQuery("SELECT * FROM tbReceive WHERE form_no = @form_no  ", Params);

            }
            catch
            {
                throw;
            }
            return m_RtnDT;
        }

        public DataTable GetAPILogInfo(string Form_No, string API_Name, string TransactionId)
        {
            DataTable m_RtnDT = new("Data");
            try
            {
                List<SqlParameter> Params = new List<SqlParameter>()
                {
                    new SqlParameter() {ParameterName = "@Form_No", SqlDbType = SqlDbType.VarChar, Value= Form_No},
                     new SqlParameter() {ParameterName = "@API_Name", SqlDbType = SqlDbType.VarChar, Value= API_Name},
                    new SqlParameter() {ParameterName = "@TransactionId", SqlDbType = SqlDbType.VarChar, Value= TransactionId}

                };

                m_RtnDT = _ADO.ExecuteQuery("SELECT TOP 1 ResultJSON  FROM tbAPILog where Form_No =@Form_No and API_Name=@API_Name and TransactionId <> @TransactionId   order by CallTime desc  ", Params);

            }
            catch
            {
                throw;
            }
            return m_RtnDT;
        }

        public ResultClass<int> InsertQCS(string form_no, string QCS_Uesr, Result_QCS p_QCS)
        {
            ResultClass<int> resultClass = new();
            int m_Execut = 0;
            string transactionId_qcs = p_QCS.TransactionId;
            List<SqlParameter> Params = new()
             {
                new SqlParameter() {ParameterName = "@form_no", SqlDbType = SqlDbType.VarChar, Value= form_no},
                new SqlParameter() {ParameterName = "@transactionId_qcs", SqlDbType = SqlDbType.VarChar, Value= transactionId_qcs},

             };
            //確認狀態
            DataTable m_tbReceive = GetReceiveByform_no(form_no);
            string m_CheckCaseStatus = "";
            foreach (DataRow dr in m_tbReceive.Rows)
            {
                m_CheckCaseStatus = dr["CaseStatus"].ToString();
            }

            Boolean isUPD = false;
            DataTable m_DT = GetAPILogInfo(form_no, "QueryCaseStatus", transactionId_qcs);
            if (m_DT.Rows.Count != 0)
            {
                foreach (DataRow row in m_DT.Rows)
                {
                    //比對reasonSuggestionDetail是否一樣,不一樣的話才更新資訊
                    Result_QCS OldResult_QCS = JsonConvert.DeserializeObject<Result_QCS>(row["ResultJSON"].ToString());
                    Boolean m_isSame = false;
                    if (m_CheckCaseStatus == "RP")//請款時要檢查capitalApply
                    {
                        m_isSame = _Common.CheckClass(p_QCS.capitalApply, OldResult_QCS.capitalApply);
                    }
                    else//非請款時要檢查reasonSuggestionDetail
                    {
                        m_isSame = _Common.CheckClass(p_QCS.reasonSuggestionDetail, OldResult_QCS.reasonSuggestionDetail);
                    }

                    if (!m_isSame)
                    {
                        isUPD = true;
                    }
                }
            }
            else
            {
                isUPD = true;
            }


            if (isUPD)
            {
                using SqlConnection conn = new SqlConnection(_ADO.GetConnStr());
                // 開啟資料庫連線
                conn.Open();
                SqlTransaction transaction = conn.BeginTransaction();
                try
                {
                    if (p_QCS is not null)
                    {
                        ResultClass<string> resultSeq = GetSeqNo("QCS");
                        if (resultSeq is not null && resultSeq.ResultCode == "000")
                        {
                            string m_qcs_idx = resultSeq.objResult;

                            string[] m_ExcTableColumn = new[] { "form_no", "qcs_idx" };
                            string[] m_ExcTableColumn_tbQCS = new[] { "form_no", "qcs_idx", "Add_User" };

                            string[] m_ExcClassColumn = new[] { "code", "msg", "Customer", "payment", "reasonSuggestionDetail", "Payee", "capitalApply" };

                            // 建立 SQL 命令
                            string m_SQL = _Common.GetInsertColumnByClass("tbQCS", m_ExcTableColumn_tbQCS, m_ExcClassColumn, p_QCS);
                            using SqlCommand command = new SqlCommand(m_SQL, conn, transaction);
                            // 設定參數
                            command.Parameters.AddWithValue("@form_no", form_no);
                            command.Parameters.AddWithValue("@qcs_idx", m_qcs_idx);
                            command.Parameters.AddWithValue("@Add_User", QCS_Uesr);

                            foreach (var prop in p_QCS.GetType().GetProperties())
                            {
                                if (Array.IndexOf(m_ExcClassColumn, prop.Name) == -1)
                                {
                                    string value = "";
                                    if (prop.GetValue(p_QCS) is not null)
                                    {
                                        value = prop.GetValue(p_QCS).ToString();
                                    }
                                    command.Parameters.AddWithValue("@" + prop.Name, _Common.CheckString(value));
                                }
                            }
                            m_Execut += command.ExecuteNonQuery();
                            if (p_QCS.Customer is not null)
                            {
                                if (p_QCS.Customer.Length != 0)
                                {
                                    foreach (Customer m_Customer in p_QCS.Customer)
                                    {
                                        // 建立 SQL 命令
                                        m_SQL = _Common.GetInsertColumnByClass("tbQCS_Customer", m_ExcTableColumn, m_ExcClassColumn, m_Customer);
                                        using SqlCommand commandC = new SqlCommand(m_SQL, conn, transaction);
                                        // 設定參數
                                        commandC.Parameters.AddWithValue("@form_no", form_no);
                                        commandC.Parameters.AddWithValue("@qcs_idx", m_qcs_idx);
                                        foreach (var prop in m_Customer.GetType().GetProperties())
                                        {
                                            if (Array.IndexOf(m_ExcClassColumn, prop.Name) == -1)
                                            {
                                                string value = "";
                                                if (prop.GetValue(m_Customer) is not null)
                                                {
                                                    value = _Common.CheckJsonType(prop.Name, prop.GetValue(m_Customer));
                                                }
                                                commandC.Parameters.AddWithValue("@" + prop.Name, _Common.CheckString(value));
                                            }
                                        }
                                        m_Execut += commandC.ExecuteNonQuery();
                                    }
                                }
                            }

                            if (p_QCS.Payee is not null)
                            {
                                // 建立 SQL 命令
                                m_SQL = _Common.GetInsertColumnByClass("tbQCS_Payee", m_ExcTableColumn, m_ExcClassColumn, p_QCS.Payee);
                                using SqlCommand commandP = new SqlCommand(m_SQL, conn, transaction);
                                // 設定參數
                                commandP.Parameters.AddWithValue("@form_no", form_no);
                                commandP.Parameters.AddWithValue("@qcs_idx", m_qcs_idx);
                                foreach (var prop in p_QCS.Payee.GetType().GetProperties())
                                {
                                    if (Array.IndexOf(m_ExcClassColumn, prop.Name) == -1)
                                    {
                                        string value = "";
                                        if (prop.GetValue(p_QCS.Payee) is not null)
                                        {
                                            value = _Common.CheckJsonType(prop.Name, prop.GetValue(p_QCS.Payee));
                                        }
                                        commandP.Parameters.AddWithValue("@" + prop.Name, _Common.CheckString(value));
                                    }
                                }
                                m_Execut += commandP.ExecuteNonQuery();
                            }

                            if (p_QCS.payment is not null)
                            {
                                if (p_QCS.payment.Length != 0)
                                {
                                    foreach (payment m_payment in p_QCS.payment)
                                    {
                                        // 建立 SQL 命令
                                        m_SQL = _Common.GetInsertColumnByClass("tbQCS_payment", m_ExcTableColumn, m_ExcClassColumn, m_payment);
                                        using SqlCommand commandP1 = new SqlCommand(m_SQL, conn, transaction);
                                        // 設定參數
                                        commandP1.Parameters.AddWithValue("@form_no", form_no);
                                        commandP1.Parameters.AddWithValue("@qcs_idx", m_qcs_idx);
                                        foreach (var prop in m_payment.GetType().GetProperties())
                                        {
                                            if (Array.IndexOf(m_ExcClassColumn, prop.Name) == -1)
                                            {
                                                string value = "";
                                                if (prop.GetValue(m_payment) is not null)
                                                {
                                                    value = _Common.CheckJsonType(prop.Name, prop.GetValue(m_payment));
                                                }
                                                commandP1.Parameters.AddWithValue("@" + prop.Name, _Common.CheckString(value));
                                            }
                                        }
                                        m_Execut += commandP1.ExecuteNonQuery();
                                    }
                                }
                            }

                            if (p_QCS.reasonSuggestionDetail is not null)
                            {
                                if (p_QCS.reasonSuggestionDetail.Length != 0)
                                {
                                    foreach (reasonSuggestionDetail m_reasonSuggestionDetail in p_QCS.reasonSuggestionDetail)
                                    {
                                        // 建立 SQL 命令
                                        m_SQL = _Common.GetInsertColumnByClass("tbQCS_reasonSuggestionDetail", m_ExcTableColumn, m_ExcClassColumn, m_reasonSuggestionDetail);
                                        using SqlCommand commandRS = new SqlCommand(m_SQL, conn, transaction);
                                        // 設定參數
                                        commandRS.Parameters.AddWithValue("@form_no", form_no);
                                        commandRS.Parameters.AddWithValue("@qcs_idx", m_qcs_idx);
                                        foreach (var prop in m_reasonSuggestionDetail.GetType().GetProperties())
                                        {
                                            if (Array.IndexOf(m_ExcClassColumn, prop.Name) == -1)
                                            {
                                                string value = "";
                                                if (prop.GetValue(m_reasonSuggestionDetail) is not null)
                                                {
                                                    value = _Common.CheckJsonType(prop.Name, prop.GetValue(m_reasonSuggestionDetail));
                                                }
                                                commandRS.Parameters.AddWithValue("@" + prop.Name, _Common.CheckString(value));
                                            }
                                        }
                                        m_Execut += commandRS.ExecuteNonQuery();
                                    }
                                }
                            }


                            if (p_QCS.capitalApply is not null)
                            {
                                if (p_QCS.capitalApply.Length != 0)
                                {
                                    foreach (capitalApply m_capitalApply in p_QCS.capitalApply)
                                    {
                                        string[] m_ExcTableColumn_Apply = new[] { "form_no", "qcs_idx", "appropriateDate" };
                                        string[] m_ExcClassColumn_Apply = new[] { "code", "msg", "appropriateDate" };
                                        // 建立 SQL 命令
                                        m_SQL = _Common.GetInsertColumnByClass("tbQCS_capitalApply", m_ExcTableColumn_Apply, m_ExcClassColumn_Apply, m_capitalApply);
                                        using SqlCommand commandRS = new SqlCommand(m_SQL, conn, transaction);
                                        // 設定參數
                                        commandRS.Parameters.AddWithValue("@form_no", form_no);
                                        commandRS.Parameters.AddWithValue("@qcs_idx", m_qcs_idx);

                                        string m_appropriateDate = "";
                                        DateTime date1;
                                        if (DateTime.TryParseExact(m_capitalApply.appropriateDate, "yyyyMMdd", null, System.Globalization.DateTimeStyles.None, out date1))
                                        {
                                            m_appropriateDate = date1.ToString("yyyy-MM-dd");
                                        }
                                        commandRS.Parameters.AddWithValue("@appropriateDate", m_appropriateDate);


                                        foreach (var prop in m_capitalApply.GetType().GetProperties())
                                        {
                                            if (Array.IndexOf(m_ExcClassColumn_Apply, prop.Name) == -1)
                                            {
                                                string value = "";
                                                if (prop.GetValue(m_capitalApply) is not null)
                                                {
                                                    value = _Common.CheckJsonType(prop.Name, prop.GetValue(m_capitalApply));
                                                }
                                                commandRS.Parameters.AddWithValue("@" + prop.Name, _Common.CheckString(value));
                                            }
                                        }
                                        m_Execut += commandRS.ExecuteNonQuery();
                                    }
                                }
                            }

                            /*UpdReceiveStatus*/
                            Int32 m_QcsExecut = 0;

                            string m_CaseStatus = "";
                            string m_approveDate = _Common.CheckString(p_QCS.approveDate);
                            DateTime date;
                            if (DateTime.TryParseExact(m_approveDate, "yyyyMMdd", null, System.Globalization.DateTimeStyles.None, out date))
                            {
                                m_approveDate = date.ToString("yyyy-MM-dd");
                            }

                            switch (p_QCS.examStatusExplain)
                            {
                                case "收件":
                                    m_CaseStatus = "1";
                                    break;
                                case "核准":
                                    m_CaseStatus = "2";
                                    break;
                                case "婉拒":
                                    m_CaseStatus = "3";
                                    break;
                                case "附條件":
                                    m_CaseStatus = "4";
                                    break;
                                case "待補":
                                    m_CaseStatus = "5";
                                    break;
                                case "補件":
                                    m_CaseStatus = "6";
                                    break;
                                case "申覆":
                                    m_CaseStatus = "7";
                                    break;
                                case "自退":
                                    m_CaseStatus = "8";
                                    break;
                            }

                            string m_QA_SQL = "";
                            string m_SQLByStatus = "";
                            string m_RE_SQL = " update  tbReceive set caseStatus=@caseStatus where form_no=@form_no  ";

                            List<SqlParameter> ParamsRE = new()
                            {
                                new SqlParameter("@form_no", form_no),
                                new SqlParameter("@CaseStatus", m_CaseStatus)
                            };

                            switch (m_CheckCaseStatus)
                            {
                                case "1"://送件--更新tbReceive - transactionId_qcs
                                    m_RE_SQL = " update  tbReceive set transactionId_qcs=@transactionId_qcs,CaseStatus=@CaseStatus where form_no=@form_no and  transactionId_qcs is null    ";
                                    ParamsRE.Add(new SqlParameter("@transactionId_qcs", transactionId_qcs));
                                    break;
                                case "RE"://申覆--更新 tbRequestforExam - transactionId_qcs
                                    m_SQLByStatus = " update  tbRequestforExam set transactionId_qcs=@transactionId_qcs where form_no=@form_no and  transactionId_qcs is null  ";
                                    break;
                                case "RS"://補件--更新tbRequestSupplement - transactionId_qcs
                                    m_SQLByStatus = " update  tbRequestSupplement set transactionId_qcs=@transactionId_qcs where form_no=@form_no and  transactionId_qcs is null    ";
                                    break;
                                case "RP"://請款--tbRequestPayment - transactionId_qcs
                                    m_SQLByStatus = " update  tbRequestPayment set transactionId_qcs=@transactionId_qcs where form_no=@form_no     ";
                                    m_QA_SQL = " update  tbQueryAppropriation set transactionId_qcs=@transactionId_qcs ,status='A004' where form_no=@form_no and  transactionId_qcs is null    ";
                                    //請款的回覆狀態改為AP已撥款
                                    foreach (SqlParameter parameter in ParamsRE)
                                    {
                                        if (parameter.ParameterName == "@CaseStatus")
                                        {
                                            parameter.Value = "AP";
                                            break;
                                        }
                                    }

                                    break;
                            }

                            if (m_RE_SQL != "")
                            {
                                using SqlCommand commandByRE = new SqlCommand(m_RE_SQL, conn, transaction);
                                commandByRE.Parameters.AddRange(ParamsRE.ToArray());
                                m_Execut += commandByRE.ExecuteNonQuery();
                            }

                            if (m_SQLByStatus != "")
                            {
                                using SqlCommand commandByStatus = new SqlCommand(m_SQLByStatus, conn, transaction);
                                commandByStatus.Parameters.AddWithValue("@form_no", form_no);
                                commandByStatus.Parameters.AddWithValue("@transactionId_qcs", transactionId_qcs);
                                m_Execut += commandByStatus.ExecuteNonQuery();
                            }

                            if (m_QA_SQL != "")
                            {
                                using SqlCommand commandByQA = new SqlCommand(m_QA_SQL, conn, transaction);
                                commandByQA.Parameters.AddWithValue("@form_no", form_no);
                                commandByQA.Parameters.AddWithValue("@transactionId_qcs", transactionId_qcs);
                                m_Execut += commandByQA.ExecuteNonQuery();
                            }


                            /*UpdReceiveStatus*/
                        }
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
            }
            return resultClass;
        }

        public ResultClass<int> InsertQueryAppropriation(string form_no, string QA_Uesr, Result_QA p_QA)
        {
            ResultClass<int> resultClass = new();
            int m_Execut = 0;
            using SqlConnection conn = new SqlConnection(_ADO.GetConnStr());
            // 開啟資料庫連線
            conn.Open();
            SqlTransaction transaction = conn.BeginTransaction();
            try
            {
                if (p_QA is not null)
                {
                    ResultClass<string> resultSeq = GetSeqNo("QA");
                    if (resultSeq is not null && resultSeq.ResultCode == "000")
                    {
                        /*檢查tbRequestPayment transactionId_qcs 是否已被更新 */
                        List<SqlParameter> ParamsRP = new()
                        {
                            new SqlParameter() {ParameterName = "@form_no", SqlDbType = SqlDbType.VarChar, Value= form_no},
                        };
                        string transactionId_qcs = "";
                        DataTable tbRequestPayment = _ADO.ExecuteQuery("SELECT isnull(transactionId_qcs,'')transactionId_qcs   FROM tbRequestPayment where form_no=@form_no ", ParamsRP);
                        foreach (DataRow dr in tbRequestPayment.Rows)
                        {
                            transactionId_qcs = dr["transactionId_qcs"].ToString();
                        }

                        string m_qa_idx = resultSeq.objResult;
                        string m_SQL = "";

                        if (transactionId_qcs != "")
                        {
                            m_SQL = "Insert into tbQueryAppropriation (form_no,qa_idx,examineNo,transactionId,transactionId_qcs,appropriationDate,status,Add_User) ";
                            m_SQL += " values (@form_no,@qa_idx,@examineNo,@transactionId,@transactionId_qcs,@appropriationDate,@status,@Add_User) ";
                        }
                        else
                        {
                            m_SQL = "Insert into tbQueryAppropriation (form_no,qa_idx,examineNo,transactionId,appropriationDate,status,Add_User) ";
                            m_SQL += " values (@form_no,@qa_idx,@examineNo,@transactionId,@appropriationDate,@status,@Add_User) ";
                        }

                        string examineNo = "";
                        string appropriationDate = "";
                        string status = "";
                        using SqlCommand commandQA = new SqlCommand(m_SQL, conn, transaction);
                        commandQA.Parameters.AddWithValue("@form_no", form_no);
                        commandQA.Parameters.AddWithValue("@qa_idx", m_qa_idx);
                        if (transactionId_qcs != "")
                        {
                            commandQA.Parameters.AddWithValue("@transactionId_qcs", transactionId_qcs);
                        }

                        if (p_QA.appropriations != null)
                        {
                            foreach (var Appro in p_QA.appropriations)
                            {

                                string m_SQL_APinfoo = "Insert into tbQueryAppropriation_APinfo (form_no,qa_idx,appropriationAmt,repayKindName) ";
                                m_SQL_APinfoo += " values (@form_no,@qa_idx,@appropriationAmt,@repayKindName) ";

                                using SqlCommand commandAP = new SqlCommand(m_SQL_APinfoo, conn, transaction);
                                commandAP.Parameters.AddWithValue("@form_no", form_no);
                                commandAP.Parameters.AddWithValue("@qa_idx", m_qa_idx);
                                commandAP.Parameters.AddWithValue("@appropriationAmt", _Common.CheckString(Appro.appropriationAmt));
                                commandAP.Parameters.AddWithValue("@repayKindName", Appro.repayKindName);
                                m_Execut += commandAP.ExecuteNonQuery();

                                appropriationDate = "";
                                DateTime date;
                                if (DateTime.TryParseExact(Appro.appropriationDate, "yyyyMMddhhmm", null, System.Globalization.DateTimeStyles.None, out date))
                                {
                                    appropriationDate = date.ToString("yyyy-MM-dd hh:mm");
                                }
                                examineNo = Appro.examineNo;
                                status = Appro.status;
                            }
                        }
                        commandQA.Parameters.AddWithValue("@examineNo", examineNo);
                        commandQA.Parameters.AddWithValue("@transactionId", p_QA.TransactionId);
                        commandQA.Parameters.AddWithValue("@appropriationDate", appropriationDate);
                        commandQA.Parameters.AddWithValue("@status", status);
                        commandQA.Parameters.AddWithValue("@Add_User", QA_Uesr);
                        m_Execut += commandQA.ExecuteNonQuery();


                        m_SQL = "update tbRequestPayment set transactionId_qa=@transactionId_qa  where form_no=@form_no ";
                        using SqlCommand commandRP = new SqlCommand(m_SQL, conn, transaction);
                        commandRP.Parameters.AddWithValue("@form_no", form_no);
                        commandRP.Parameters.AddWithValue("@transactionId_qa", p_QA.TransactionId);


                        m_Execut += commandRP.ExecuteNonQuery();
                    }
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

        public Int32 UpdReceiveStatus(string Form_No, Result_QCS p_Result_QCS)
        {
            Int32 m_Execut = 0;
            Int32 m_QcsExecut = 0;
            string transactionId_qcs = p_Result_QCS.TransactionId;
            try
            {
                string m_CaseStatus = "";
                string m_approveDate = _Common.CheckString(p_Result_QCS.approveDate);

                DateTime date;
                if (DateTime.TryParseExact(m_approveDate, "yyyyMMdd", null, System.Globalization.DateTimeStyles.None, out date))
                {
                    m_approveDate = date.ToString("yyyy-MM-dd");
                }

                switch (p_Result_QCS.examStatusExplain)
                {
                    case "收件":
                        m_CaseStatus = "1";
                        break;
                    case "核准":
                        m_CaseStatus = "2";
                        break;
                    case "婉拒":
                        m_CaseStatus = "3";
                        break;
                    case "附條件":
                        m_CaseStatus = "4";
                        break;
                    case "待補":
                        m_CaseStatus = "5";
                        break;
                    case "補件":
                        m_CaseStatus = "6";
                        break;
                    case "申覆":
                        m_CaseStatus = "7";
                        break;
                    case "自退":
                        m_CaseStatus = "8";
                        break;
                }

                List<SqlParameter> Params = new List<SqlParameter>()
                {
                    new SqlParameter() {ParameterName = "@form_no", SqlDbType = SqlDbType.VarChar, Value= Form_No},
                    new SqlParameter() {ParameterName = "@transactionId_qcs", SqlDbType = SqlDbType.VarChar, Value= transactionId_qcs},

                };

                DataTable m_tbReceive = GetReceiveByform_no(Form_No);
                string m_CheckCaseStatus = "";
                foreach (DataRow dr in m_tbReceive.Rows)
                {
                    m_CheckCaseStatus = dr["CaseStatus"].ToString();
                }
                string m_SQL = "";
                string m_SQL1 = "";

                switch (m_CheckCaseStatus)
                {
                    case "1"://送件--更新tbReceive - transactionId_qcs
                        m_SQL = " update  tbReceive set transactionId_qcs=@transactionId_qcs where form_no=@form_no and  transactionId_qcs is null    ";
                        break;
                    case "RE"://申覆--更新 tbRequestforExam - transactionId_qcs
                        m_SQL = " update  tbRequestforExam set transactionId_qcs=@transactionId_qcs where form_no=@form_no and  transactionId_qcs is null  ";
                        break;
                    case "RS"://補件--更新tbRequestSupplement - transactionId_qcs
                        m_SQL = " update  tbRequestSupplement set transactionId_qcs=@transactionId_qcs where form_no=@form_no and  transactionId_qcs is null    ";
                        break;
                    case "RP"://請款--tbRequestPayment - transactionId_qcs
                        m_SQL = " update  tbQueryAppropriation set transactionId_qcs=@transactionId_qcs where form_no=@form_no and  transactionId_qcs is null    ";
                        m_SQL1 = " update  tbRequestPayment set transactionId_qcs=@transactionId_qcs where form_no=@form_no and  transactionId_qcs is null    ";
                        m_QcsExecut += _ADO.ExecuteNonQuery(m_SQL1, Params);
                        break;
                }
                m_QcsExecut += _ADO.ExecuteNonQuery(m_SQL, Params);


                List<SqlParameter> ParamsR = new List<SqlParameter>()
                {
                    new SqlParameter() {ParameterName = "@form_no", SqlDbType = SqlDbType.VarChar, Value= Form_No},
                    new SqlParameter() {ParameterName = "@approveDate", SqlDbType = SqlDbType.VarChar, Value= m_approveDate},
                    new SqlParameter() {ParameterName = "@CaseStatus", SqlDbType = SqlDbType.VarChar, Value= m_CaseStatus}
                };


                m_Execut = _ADO.ExecuteNonQuery("update tbReceive set " +
                    " approveDate=@approveDate," +
                    " CaseStatus=@CaseStatus " +
                    "  where  form_no=@form_no ", ParamsR);
            }
            catch
            {
                throw;
            }
            return m_Execut;
        }

        public string GetFromNoByExamineNo(string ExamineNo)
        {
            List<SqlParameter> Params = new List<SqlParameter>()
            {
                new SqlParameter() {ParameterName = "@ExamineNo", SqlDbType = SqlDbType.VarChar, Value= ExamineNo}
            };

            DataTable m_RtnDT = _ADO.ExecuteQuery("SELECT form_no FROM tbReceive WHERE ExamineNo = @ExamineNo  ", Params);
            string m_Form_No = "";
            foreach (DataRow dr in m_RtnDT.Rows)
            {
                m_Form_No = dr["form_no"].ToString();
            }
            return m_Form_No;
        }

        public Int32 UpdByRequestSupplement(string Form_No, string RS_User, List<string> REcomment, string transactionId)
        {
            Int32 m_Execut = 0;
            try
            {
                List<SqlParameter> Params = new List<SqlParameter>()
                {
                    new SqlParameter() {ParameterName = "@form_no", SqlDbType = SqlDbType.VarChar, Value= Form_No},
                    new SqlParameter() {ParameterName = "@RS_User", SqlDbType = SqlDbType.VarChar, Value= RS_User}
                };
                m_Execut = _ADO.ExecuteNonQuery("update tbReceive set RS_User=@RS_User,RS_Count=(RS_Count+1),CaseStatus='RS'  where  form_no=@form_no ", Params);
                if (m_Execut == 1)
                {
                    Params = new List<SqlParameter>()
                    {
                        new SqlParameter() {ParameterName = "@form_no", SqlDbType = SqlDbType.VarChar, Value= Form_No}
                    };

                    DataTable m_RtnDT = _ADO.ExecuteQuery("SELECT RS_Count FROM tbReceive WHERE form_no = @form_no  ", Params);
                    string m_rs_idx = "";
                    foreach (DataRow dr in m_RtnDT.Rows)
                    {
                        m_rs_idx = dr["RS_Count"].ToString();
                    }

                    Params = new List<SqlParameter>()
                    {
                        new SqlParameter() {ParameterName = "@form_no", SqlDbType = SqlDbType.VarChar, Value= Form_No},
                        new SqlParameter() {ParameterName = "@rs_idx", SqlDbType = SqlDbType.VarChar, Value= m_rs_idx},
                        new SqlParameter() {ParameterName = "@transactionId", SqlDbType = SqlDbType.VarChar, Value= transactionId},
                        new SqlParameter() {ParameterName = "@Add_User", SqlDbType = SqlDbType.VarChar, Value= RS_User},
                    };

                    Int32 m_CommCount = 1;
                    string m_AddColumn = "";
                    string m_AddValue = "";

                    foreach (string comment in REcomment)
                    {
                        string m_ColumnNa = "comment" + m_CommCount.ToString();
                        m_AddColumn += "," + m_ColumnNa;
                        m_AddValue += ",@" + m_ColumnNa;
                        SqlParameter param = new() { ParameterName = "@" + m_ColumnNa, SqlDbType = SqlDbType.NVarChar, Value = comment };
                        Params.Add(param);
                        m_CommCount++;
                    }

                    m_Execut += _ADO.ExecuteNonQuery("insert into tbRequestSupplement (form_no,rs_idx,transactionId,Add_User " + m_AddColumn + ")" +
                        " values(@form_no,@rs_idx,@transactionId,@Add_User" + m_AddValue + ")   ", Params);

                }
            }
            catch
            {
                throw;
            }
            return m_Execut;
        }

        public Int32 UpdByRequestPayment(string Form_No, string RP_User, string transactionId, BankInfo BankInfo, PayInfo PayInfo)
        {
            Int32 m_Execut = 0;
            try
            {
                if (PayInfo.remitAmount == "")
                {
                    PayInfo.remitAmount = "0";
                }
                List<SqlParameter> Params = new List<SqlParameter>()
                {
                    new SqlParameter() {ParameterName = "@instNo", SqlDbType = SqlDbType.Int, Value= PayInfo.instNo},
                    new SqlParameter() {ParameterName = "@instAmt", SqlDbType = SqlDbType.Int, Value= PayInfo.instAmt},
                    new SqlParameter() {ParameterName = "@remitAmount", SqlDbType = SqlDbType.Int, Value= PayInfo.remitAmount},
                    new SqlParameter() {ParameterName = "@instCap", SqlDbType = SqlDbType.Int, Value= PayInfo.instCap},
                    new SqlParameter() {ParameterName = "@BankCode", SqlDbType = SqlDbType.VarChar, Value= BankInfo.BankCode},
                    new SqlParameter() {ParameterName = "@BankName", SqlDbType = SqlDbType.VarChar, Value= BankInfo.BankName},
                    new SqlParameter() {ParameterName = "@BankID", SqlDbType = SqlDbType.VarChar, Value= BankInfo.BankID},
                    new SqlParameter() {ParameterName = "@AccountID", SqlDbType = SqlDbType.VarChar, Value= BankInfo.AccountID},
                    new SqlParameter() {ParameterName = "@AccountName", SqlDbType = SqlDbType.NVarChar, Value= BankInfo.AccountName},
                    new SqlParameter() {ParameterName = "@form_no", SqlDbType = SqlDbType.VarChar, Value= Form_No},
                    new SqlParameter() {ParameterName = "@RP_User", SqlDbType = SqlDbType.VarChar, Value= RP_User}
                };
                m_Execut = _ADO.ExecuteNonQuery("update tbReceive set instCap=@instCap, instNo=@instNo, instAmt=@instAmt, remitAmount=@remitAmount, BankCode=@BankCode, BankName=@BankName, BankID=@BankID, AccountID=@AccountID, AccountName=@AccountName, RP_User=@RP_User,RP_Count=(RP_Count+1),CaseStatus='RP'  where  form_no=@form_no ", Params);
                if (m_Execut == 1)
                {
                    Params = new List<SqlParameter>()
                    {
                        new SqlParameter() {ParameterName = "@form_no", SqlDbType = SqlDbType.VarChar, Value= Form_No}
                    };

                    DataTable m_RtnDT = _ADO.ExecuteQuery("SELECT RP_Count FROM tbReceive WHERE form_no = @form_no  ", Params);
                    string m_rp_idx = "";
                    foreach (DataRow dr in m_RtnDT.Rows)
                    {
                        m_rp_idx = dr["RP_Count"].ToString();
                    }

                    Params = new List<SqlParameter>()
                    {
                        new SqlParameter() {ParameterName = "@form_no", SqlDbType = SqlDbType.VarChar, Value= Form_No},
                        new SqlParameter() {ParameterName = "@rp_idx", SqlDbType = SqlDbType.VarChar, Value= m_rp_idx},
                        new SqlParameter() {ParameterName = "@transactionId", SqlDbType = SqlDbType.VarChar, Value= transactionId},
                         new SqlParameter() {ParameterName = "@Add_User", SqlDbType = SqlDbType.VarChar, Value= RP_User}
                    };


                    m_Execut += _ADO.ExecuteNonQuery("insert into tbRequestPayment (form_no,rp_idx,transactionId,Add_User)" +
                        " values(@form_no,@rp_idx,@transactionId,@Add_User)   ", Params);

                }
            }
            catch
            {
                throw;
            }
            return m_Execut;
        }

        public Int32 UpdByRequestforExam(string Form_No, string forceTryForExam, string RE_User, string comment, string transactionId)
        {
            Int32 m_Execut = 0;
            try
            {
                List<SqlParameter> Params = new List<SqlParameter>()
                {
                    new SqlParameter() {ParameterName = "@form_no", SqlDbType = SqlDbType.VarChar, Value= Form_No}
                };

                m_Execut = _ADO.ExecuteNonQuery("update tbReceive set RE_count=(RE_count+1),CaseStatus='RE'   where  form_no=@form_no ", Params);
                if (m_Execut == 1)
                {
                    Params = new List<SqlParameter>()
                    {
                        new SqlParameter() {ParameterName = "@form_no", SqlDbType = SqlDbType.VarChar, Value= Form_No}
                    };

                    DataTable m_RtnDT = _ADO.ExecuteQuery("SELECT RE_count FROM tbReceive WHERE form_no = @form_no  ", Params);
                    string m_re_idx = "";
                    foreach (DataRow dr in m_RtnDT.Rows)
                    {
                        m_re_idx = dr["RE_count"].ToString();
                    }

                    Params = new List<SqlParameter>()
                    {
                        new SqlParameter() {ParameterName = "@form_no", SqlDbType = SqlDbType.VarChar, Value= Form_No},
                        new SqlParameter() {ParameterName = "@re_idx", SqlDbType = SqlDbType.VarChar, Value= m_re_idx},
                        new SqlParameter() {ParameterName = "@comment", SqlDbType = SqlDbType.NVarChar, Value= comment},
                        new SqlParameter() {ParameterName = "@forceTryForExam", SqlDbType = SqlDbType.Char, Value= forceTryForExam},
                        new SqlParameter() {ParameterName = "@transactionId", SqlDbType = SqlDbType.VarChar, Value= transactionId},
                        new SqlParameter() {ParameterName = "@Add_User", SqlDbType = SqlDbType.VarChar, Value= RE_User},
                    };
                    m_Execut += _ADO.ExecuteNonQuery("insert into tbRequestforExam (form_no,re_idx,comment,forceTryForExam,transactionId,Add_User)" +
                        " values(@form_no,@re_idx,@comment,@forceTryForExam,@transactionId,@Add_User)   ", Params);
                }
            }
            catch
            {
                throw;
            }
            return m_Execut;
        }

        public Int32 UpdReceive(string Form_No, string TransactionId, string ExamineNo)
        {
            Int32 m_Execut = 0;
            try
            {
                List<SqlParameter> Params = new List<SqlParameter>()
                {
                    new SqlParameter() {ParameterName = "@form_no", SqlDbType = SqlDbType.VarChar, Value= Form_No},
                    new SqlParameter() {ParameterName = "@TransactionId", SqlDbType = SqlDbType.VarChar, Value= TransactionId},
                    new SqlParameter() {ParameterName = "@CaseStatus", SqlDbType = SqlDbType.VarChar, Value= "1"},
                    new SqlParameter() {ParameterName = "@ExamineNo", SqlDbType = SqlDbType.VarChar, Value= ExamineNo}
                };

                m_Execut = _ADO.ExecuteNonQuery("update tbReceive set  CaseStatus=@CaseStatus, ExamineNo=@ExamineNo,TransactionId=@TransactionId where  form_no=@form_no ", Params);

                Params = new List<SqlParameter>()
                {
                    new SqlParameter() {ParameterName = "@form_no", SqlDbType = SqlDbType.VarChar, Value= Form_No},
                    new SqlParameter() {ParameterName = "@Upload_Type", SqlDbType = SqlDbType.VarChar, Value= "Receive"},
                    new SqlParameter() {ParameterName = "@TransactionId", SqlDbType = SqlDbType.VarChar, Value= TransactionId}
                };


                m_Execut = _ADO.ExecuteNonQuery("update tbFiles set transactionId=@transactionId where  KeyID=@form_no  and Upload_Type=@Upload_Type", Params);

            }
            catch
            {
                throw;
            }
            return m_Execut;
        }

        public Int32 GetFileCount(string form_no, string Upload_Type)
        {
            Int32 m_FileCount = 0;
            try
            {
                List<SqlParameter> Params = new List<SqlParameter>()
                {
                    new SqlParameter() {ParameterName = "@KeyID", SqlDbType = SqlDbType.VarChar, Value= form_no},
                    new SqlParameter() {ParameterName = "@Upload_Type", SqlDbType = SqlDbType.VarChar, Value= Upload_Type}
                };
                DataTable m_dtFiles = _ADO.ExecuteQuery("select count(*) FileCount FROM tbFiles where KeyID=@KeyID and  Upload_Type=@Upload_Type ", Params);
                foreach (DataRow dr in m_dtFiles.Rows)
                {
                    m_FileCount = Convert.ToInt32(dr["FileCount"]);
                }
            }
            catch
            {
                throw;
            }
            return m_FileCount;
        }

        public ResultClass<List<tbQCS>> GetQCS(string form_no)
        {
            ResultClass<List<tbQCS>> resultClass = new();
            try
            {
                List<SqlParameter> Params = new List<SqlParameter>()
                {
                    new SqlParameter() {ParameterName = "@form_no", SqlDbType = SqlDbType.VarChar, Value= form_no}
                };

                string m_SQL = "  SELECT Q.examineComment,Q.form_no,Q.qcs_idx,case when CONVERT(varchar(16), Q.Add_date, 120) is null then '' else CONVERT(varchar(16), Q.Add_date, 120) end qcs_time,Q.examStatusExplain explain,Q.transactionId ,ResulType " +
                    " FROM tbQCS Q left join " +
                    "(" +
                      " select '送件'ResulType,  form_no, transactionId_qcs,'' transactionId_qa  from tbReceive union all " +
                      " select '申覆'ResulType,  form_no, transactionId_qcs,'' transactionId_qa  from tbRequestforExam union all  " +
                      " select '補件'ResulType,  form_no, transactionId_qcs,'' transactionId_qa  from tbRequestSupplement union all      " +
                      " select '請款'ResulType,  form_no, transactionId_qcs,isnull(transactionId_qa,'')transactionId_qa  from tbRequestPayment  " +
                    " ) ResulFrom on Q.form_no=ResulFrom.form_no and Q.transactionId=ResulFrom.transactionId_qcs   " +
                    " where Q.form_no =@form_no  and ResulType <> '' order by qcs_idx ";

                DataTable m_dtQCS = _ADO.ExecuteQuery(m_SQL, Params);
                List<tbQCS> m_listbQCS = new();
                string qcs_idx = "";

                tbQCS m_tbQCS = new();
                foreach (DataRow row in m_dtQCS.Rows)
                {
                    m_tbQCS = new();
                    m_tbQCS.form_no = row["form_no"].ToString();
                    m_tbQCS.qcs_idx = row["qcs_idx"].ToString();
                    m_tbQCS.qcs_time = row["qcs_time"].ToString();
                    m_tbQCS.explain = row["explain"].ToString();
                    m_tbQCS.transactionId = row["transactionId"].ToString();
                    m_tbQCS.resulType = row["resulType"].ToString();
                    qcs_idx = row["qcs_idx"].ToString();

                    List<SqlParameter> Paramsdtl = new();
                    List<SqlParameter> Paramsdt2 = new();
                    string comment = "";

                    if (row["examineComment"].ToString() != "")
                    {
                        comment = row["examineComment"].ToString() + "<br>";
                    }


                    if (m_tbQCS.resulType != "請款")
                    {
                        Paramsdtl = new List<SqlParameter>()
                        {
                            new SqlParameter() {ParameterName = "@form_no", SqlDbType = SqlDbType.VarChar, Value= form_no},
                            new SqlParameter() {ParameterName = "@qcs_idx", SqlDbType = SqlDbType.VarChar, Value= qcs_idx}
                        };
                        DataTable m_dtQCSdtl = _ADO.ExecuteQuery(" SELECT kind+':'+explain +','+comment comment " +
                         "    FROM tbQCS_reasonSuggestionDetail   where form_no =@form_no and qcs_idx =@qcs_idx order by detail_idx ", Paramsdtl);

                        foreach (DataRow row1 in m_dtQCSdtl.Rows)
                        {
                            comment += row1["comment"].ToString();
                        }

                        Paramsdt2 = new List<SqlParameter>()
                        {
                            new SqlParameter() {ParameterName = "@form_no", SqlDbType = SqlDbType.VarChar, Value= form_no},
                            new SqlParameter() {ParameterName = "@qcs_idx", SqlDbType = SqlDbType.VarChar, Value= qcs_idx}
                        };
                        DataTable m_dtQCSCus = _ADO.ExecuteQuery(" SELECT  name+':'+calloutResult comment FROM tbQCS_Customer    where form_no =@form_no and qcs_idx =@qcs_idx and calloutResult <>'' order by customer_idx ", Paramsdt2);

                        foreach (DataRow row1 in m_dtQCSCus.Rows)
                        {
                            comment += "<br>" + row1["comment"].ToString();
                        }

                    }
                    else
                    {
                        Paramsdtl = new List<SqlParameter>()
                        {
                            new SqlParameter() {ParameterName = "@form_no", SqlDbType = SqlDbType.VarChar, Value= form_no},
                            new SqlParameter() {ParameterName = "@qcs_idx", SqlDbType = SqlDbType.VarChar, Value= qcs_idx}
                        };
                        DataTable m_dtQCSdtl = _ADO.ExecuteQuery(" SELECT  pa.instNo,pa.instAmt,   QC.[form_no],QC.[qcs_idx],QC.transactionId,isnull([appropriateDate],'')appropriateDate,isnull( CONVERT(varchar(16),sum([remitAmount])),'')remitAmount" +
                            " , case when QA.status='A004' then '已撥款' when QA.status='A003' then '申請中' end  status,APCount " +
                         "     FROM　[tbQCS]　QC left join [tbQCS_capitalApply] QCA on QC.form_no=QCA.form_no and QC.qcs_idx=QCA.qcs_idx " +
                         "  left join (" +
                         "    select qa.*,APCount,status FROM   " +
                         "      (select  form_no,transactionId_qcs,max(qa.qa_idx)qa_idx from tbQueryAppropriation qa group by  form_no,transactionId_qcs)qa" +
                         "    left join" +
                         "    (select form_no,qa_idx ,count(form_no+qa_idx)APCount FROM tbQueryAppropriation_APinfo group by form_no,qa_idx) AP " +
                         "     on qa.form_no=AP.form_no and qa.qa_idx=Ap.qa_idx " +
                         "     left join " +
                         "    (select form_no,qa_idx,status  FROM tbQueryAppropriation ) QAS   " +
                         "   on qa.form_no=QAS.form_no and qa.qa_idx=QAS.qa_idx  " +
                         ") QA on QC.form_no=QA.form_no and QC.transactionId=QA.transactionId_qcs  " +
                         "   Left Join tbQCS_payment PA on  QC.[form_no]= PA.[form_no] and  QC.[qcs_idx]= PA.[qcs_idx]  " +
                         "   where   [appropriateDate] is not null and QC.[form_no]=@form_no and QC.[qcs_idx]=@qcs_idx  " +
                         " group by QC.[form_no]  ,pa.instNo,pa.instAmt  ,QC.transactionId,QC.[qcs_idx],[appropriateDate],QA.status,APCount   ", Paramsdtl);

                        foreach (DataRow row1 in m_dtQCSdtl.Rows)
                        {
                            if (row1["appropriateDate"].ToString() != "")
                            {
                                comment = "撥款日期:" + row1["appropriateDate"].ToString() + ";";
                            }
                            if (row1["remitAmount"].ToString() != "")
                            {
                                if (row1["APCount"].ToString() != "1")
                                {
                                    comment += "**";
                                }
                                comment += "撥款金額:" + row1["remitAmount"].ToString();
                            }

                            if (row1["instNo"].ToString() != "" && row1["instAmt"].ToString() != "")
                            {
                                comment += "期數:" + row1["instNo"].ToString() + ";分期應繳:" + row1["instAmt"].ToString();
                            }


                            if (row1["status"].ToString() != "")
                            {
                                m_tbQCS.explain = row1["status"].ToString();
                            }
                        }
                    }

                    m_tbQCS.comment = comment;
                    m_listbQCS.Add(m_tbQCS);
                }
                resultClass.objResult = m_listbQCS;
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "999";
                resultClass.ResultMsg = ex.Message;
                resultClass.objResult = null;
            }
            return resultClass;
        }

        public ResultClass<List<tbRE>> GetRE(string form_no)
        {
            ResultClass<List<tbRE>> resultClass = new();
            try
            {
                List<SqlParameter> Params = new List<SqlParameter>()
                {
                    new SqlParameter() {ParameterName = "@form_no", SqlDbType = SqlDbType.VarChar, Value= form_no}
                };

                DataTable m_dtQCS = _ADO.ExecuteQuery(" SELECT  re.transactionId ,re.form_no,re.re_idx,case when CONVERT(varchar(16), re.Add_date, 120) is null then '' else CONVERT(varchar(16), re.Add_date, 120) end re_time,comment ,transactionId_qcs ,case when CONVERT(varchar(16), Q.Add_date, 120) is null then '' else CONVERT(varchar(16), Q.Add_date, 120) end qcs_time,isnull(Q.examStatusExplain,'未回覆') explain ,qcs_idx  " +
                    "    FROM  tbRequestforExam re  left join tbQCS Q on re.form_no=Q.form_no  and re.transactionId_qcs=Q.transactionId  where re.form_no =@form_no order by re_idx asc  ", Params);
                List<tbRE> m_listbRE = new();
                foreach (DataRow row in m_dtQCS.Rows)
                {
                    tbRE m_tbRE = new();
                    m_tbRE.form_no = row["form_no"].ToString();
                    m_tbRE.re_idx = row["re_idx"].ToString();
                    m_tbRE.re_time = row["re_time"].ToString();
                    m_tbRE.comment = row["comment"].ToString();
                    m_tbRE.transactionId_qcs = row["transactionId_qcs"].ToString();
                    m_tbRE.explain = row["explain"].ToString();


                    m_tbRE.qcs_time = row["qcs_time"].ToString();
                    m_tbRE.transactionId = row["transactionId"].ToString();
                    string qcs_idx = row["qcs_idx"].ToString();
                    List<SqlParameter> Paramsdtl = new List<SqlParameter>()
                    {
                        new SqlParameter() {ParameterName = "@form_no", SqlDbType = SqlDbType.VarChar, Value= form_no},
                        new SqlParameter() {ParameterName = "@qcs_idx", SqlDbType = SqlDbType.VarChar, Value= qcs_idx}
                    };

                    DataTable m_dtQCSdtl = _ADO.ExecuteQuery(" SELECT kind+':'+explain +','+comment qcs_comment " +
                      "    FROM tbQCS_reasonSuggestionDetail   where form_no =@form_no and qcs_idx =@qcs_idx order by detail_idx ", Paramsdtl);
                    string qcs_comment = "";
                    foreach (DataRow row1 in m_dtQCSdtl.Rows)
                    {
                        qcs_comment += row1["qcs_comment"].ToString();
                    }
                    m_tbRE.qcs_comment = qcs_comment;

                    m_listbRE.Add(m_tbRE);
                }

                resultClass.objResult = m_listbRE;
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "999";
                resultClass.ResultMsg = ex.Message;
                resultClass.objResult = null;
            }
            return resultClass;
        }

        public ResultClass<List<tbRS>> GetRS(string form_no)
        {
            ResultClass<List<tbRS>> resultClass = new();
            try
            {
                List<SqlParameter> Params = new List<SqlParameter>()
                {
                    new SqlParameter() {ParameterName = "@form_no", SqlDbType = SqlDbType.VarChar, Value= form_no}
                };

                DataTable m_dtQCS = _ADO.ExecuteQuery(" SELECT  rs.transactionId ,rs.form_no,rs.rs_idx,  case when CONVERT(varchar(16), rs.Add_date, 120) is null then '' else CONVERT(varchar(16), rs.Add_date, 120) end rs_time,comment1,comment2,comment3,comment4,comment5 ,transactionId_qcs ,case when CONVERT(varchar(16), Q.Add_date, 120) is null then '' else CONVERT(varchar(16), Q.Add_date, 120) end qcs_time,isnull(Q.examStatusExplain,'未回覆') explain ,qcs_idx  " +
                    "    FROM  tbRequestSupplement rs  left join tbQCS Q on rs.form_no=Q.form_no  and rs.transactionId_qcs=Q.transactionId  where rs.form_no =@form_no order by rs_idx asc  ", Params);
                List<tbRS> m_listbRS = new();
                foreach (DataRow row in m_dtQCS.Rows)
                {
                    tbRS m_tbRS = new();
                    m_tbRS.form_no = row["form_no"].ToString();
                    m_tbRS.rs_idx = row["rs_idx"].ToString();
                    m_tbRS.rs_time = row["rs_time"].ToString();
                    m_tbRS.comment1 = row["comment1"].ToString();
                    m_tbRS.comment2 = row["comment2"].ToString();
                    m_tbRS.comment3 = row["comment3"].ToString();
                    m_tbRS.comment4 = row["comment4"].ToString();
                    m_tbRS.comment5 = row["comment5"].ToString();
                    m_tbRS.transactionId_qcs = row["transactionId_qcs"].ToString();
                    m_tbRS.qcs_time = row["qcs_time"].ToString();

                    m_tbRS.explain = row["explain"].ToString();

                    m_tbRS.transactionId = row["transactionId"].ToString();
                    string qcs_idx = row["qcs_idx"].ToString();
                    List<SqlParameter> Paramsdtl = new List<SqlParameter>()
                    {
                        new SqlParameter() {ParameterName = "@form_no", SqlDbType = SqlDbType.VarChar, Value= form_no},
                        new SqlParameter() {ParameterName = "@qcs_idx", SqlDbType = SqlDbType.VarChar, Value= qcs_idx}
                    };

                    DataTable m_dtQCSdtl = _ADO.ExecuteQuery(" SELECT kind+':'+explain +','+comment qcs_comment " +
                      "    FROM tbQCS_reasonSuggestionDetail   where form_no =@form_no and qcs_idx =@qcs_idx order by detail_idx ", Paramsdtl);
                    string qcs_comment = "";
                    foreach (DataRow row1 in m_dtQCSdtl.Rows)
                    {
                        qcs_comment += row1["qcs_comment"].ToString();
                    }
                    m_tbRS.qcs_comment = qcs_comment;

                    m_listbRS.Add(m_tbRS);
                }

                resultClass.objResult = m_listbRS;
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "999";
                resultClass.ResultMsg = ex.Message;
                resultClass.objResult = null;
            }
            return resultClass;
        }

        public ResultClass<List<tbRP>> GetRP(string form_no)
        {
            ResultClass<List<tbRP>> resultClass = new();
            try
            {
                List<SqlParameter> Params = new List<SqlParameter>()
                {
                    new SqlParameter() {ParameterName = "@form_no", SqlDbType = SqlDbType.VarChar, Value= form_no}
                };

                DataTable m_dtQCS = _ADO.ExecuteQuery(" SELECT  rp.transactionId  ,qcs_idx,rp.form_no,rp.rp_idx,rp.transactionId_qcs ," +
                    "case when CONVERT(varchar(16), rp.Add_date, 120) is null then '' else CONVERT(varchar(16), rp.Add_date, 120) end rp_time, " +
                    "case when CONVERT(varchar(16), Q.Add_date, 120) is null then '' else CONVERT(varchar(16), Q.Add_date, 120) end qcs_time, " +
                    "case when CONVERT(varchar(16), QA.Add_date, 120) is null then '' else CONVERT(varchar(16), QA.Add_date, 120) end qa_time, " +
                    "isnull(QA.status,'') status ,isnull( case  isnull(QA.status,'未回覆')    " +
                    "    when 'A001' then '未申請' when 'A002' then '申請中'  " +
                    "    when 'A003' then '撥款中' when 'A004' then '已撥款' " +
                    "  end,'未回覆') StatusDesc " +
                    "     FROM  tbRequestPayment rp  left join tbQCS Q on rp.form_no=Q.form_no  and rp.transactionId_qcs=Q.transactionId " +
                    " left join tbQueryAppropriation QA on rp.form_no=QA.form_no  and rp.transactionId_qa=QA.transactionId " +
                    " where rp.form_no =@form_no order by rp_idx asc  ", Params);
                List<tbRP> m_listbRP = new();
                foreach (DataRow row in m_dtQCS.Rows)
                {
                    tbRP m_tbRP = new();
                    m_tbRP.form_no = row["form_no"].ToString();
                    m_tbRP.rp_idx = row["rp_idx"].ToString();
                    m_tbRP.rp_time = row["rp_time"].ToString();

                    m_tbRP.transactionId_qcs = row["transactionId_qcs"].ToString();
                    if (row["qcs_time"].ToString() == "")
                    {
                        m_tbRP.resq_time = row["qa_time"].ToString();
                    }
                    else
                    {
                        m_tbRP.resq_time = row["qcs_time"].ToString();
                    }


                    m_tbRP.transactionId = row["transactionId"].ToString();

                    m_tbRP.statusDesc = row["StatusDesc"].ToString();

                    string qcs_idx = row["qcs_idx"].ToString();
                    if (qcs_idx != "")
                    {
                        List<SqlParameter> Paramsdtl = new List<SqlParameter>()
                        {
                            new SqlParameter() {ParameterName = "@form_no", SqlDbType = SqlDbType.VarChar, Value= form_no},
                            new SqlParameter() {ParameterName = "@qcs_idx", SqlDbType = SqlDbType.VarChar, Value= qcs_idx}
                        };
                        if (row["status"].ToString() == "A004")
                        {
                            DataTable m_dtQCSdtl = _ADO.ExecuteQuery(" SELECT appropriateDate,remitAmount,payeeTypeName " +
                              "    FROM tbQCS_capitalApply   where form_no =@form_no and qcs_idx =@qcs_idx order by capitalApply_idx ", Paramsdtl);
                            if (m_dtQCSdtl.Rows.Count != 0)
                            {
                                capitalApply[] m_arrCapitalApply = new capitalApply[m_dtQCSdtl.Rows.Count];
                                Int32 Count = 0;
                                foreach (DataRow row1 in m_dtQCSdtl.Rows)
                                {
                                    capitalApply m_capitalApply = new();
                                    m_capitalApply.appropriateDate += row1["appropriateDate"].ToString();
                                    m_capitalApply.remitAmount += row1["remitAmount"].ToString();
                                    m_capitalApply.payeeTypeName += row1["payeeTypeName"].ToString();
                                    m_arrCapitalApply[Count] = m_capitalApply;
                                    Count++;
                                }
                                m_tbRP.capitalApply = m_arrCapitalApply;
                            }
                        }
                    }

                    m_listbRP.Add(m_tbRP);
                }

                resultClass.objResult = m_listbRP;
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "999";
                resultClass.ResultMsg = ex.Message;
                resultClass.objResult = null;
            }
            return resultClass;
        }

        public ResultClass<List<objReceive>> GetReceive(string form_no)
        {
            ResultClass<List<objReceive>> resultClass = new();
            try
            {

                Int32 m_FileCount = GetFileCount(form_no, "Receive");
                Int32 m_RSFileCount = GetFileCount(form_no, "RequestSupplement");
                Int32 m_REFileCount = GetFileCount(form_no, "RequestforExam");
                Int32 m_RPFileCount = GetFileCount(form_no, "RequestPayment");

                List<SqlParameter> Params = new List<SqlParameter>()
                {
                    new SqlParameter() {ParameterName = "@form_no", SqlDbType = SqlDbType.VarChar, Value= form_no}
                };

                DataTable m_dtReceive = _ADO.ExecuteQuery("select  *,case when CONVERT(varchar(10), Add_date, 120) is null then '' else CONVERT(varchar(10), Add_date, 120) end Add_time_dis " +
                    ", case  isnull(RP_Status,'')   " +
                    "   when '' then " +
                    "       case  isnull(CaseStatus,'0') " +
                    "           when 'RP' then '請款中' " +
                    "           when '2' then '待請款' " +
                    "       end " +
                    "    when 'A001' then '未申請' when 'A002' then '申請中'  " +
                    "    when 'A003' then '撥款中' when 'A004' then '已撥款' " +
                    "  end RP_StatusDesc  " +
                    ", case  isnull(CaseStatus,'0')   " +
                    "    when '0' then '未轉裕富' when '1' then '裕富收件'   " +
                    "    when '2' then '核准'     when '3' then '婉拒'   " +
                    "    when '4' then '附條件'   when '5' then '待補'  " +
                    "    when '6' then '補件'     when '7' then '申覆' " +
                    "    when '8' then '自退'  when 'RE' then '申覆中' " +
                     "   when 'RS' then '補件中'  when 'RP' then '請款中'  when 'AP' then '已撥款' " +
                    "  end CaseStatusDesc " +
                    " FROM tbReceive where form_no =@form_no and status='1' ", Params);
                List<objReceive> receiveList = new List<objReceive>();
                // 將資料表轉換為 List<Receive> 類型
                foreach (DataRow row in m_dtReceive.Rows)
                {
                    objReceive receive = new objReceive();
                    foreach (DataColumn column in m_dtReceive.Columns)
                    {
                        var Key = column.ColumnName.ToLower();
                        var Value = row[Key].ToString();
                        IDictionary<string, string> NewPV = new Dictionary<string, string>() { { Key, Value } };
                        if (NewPV.TryGetValue(Key, out var m_value))
                        {
                            if (receive.GetType().GetProperty(Key) != null)
                            {
                                receive.GetType().GetProperty(Key).SetValue(receive, m_value.ToString());
                            }
                        }
                    }
                    receive.add_date = row["Add_time_dis"].ToString();
                    if (m_FileCount != 0)
                    {
                        receive.attachmentFile = GetFilesByKeyID(form_no, "Receive", m_FileCount);
                    }
                    if (m_REFileCount != 0)
                    {
                        receive.rEattachmentFile = GetFilesByKeyID(form_no, "RequestforExam", m_REFileCount);
                    }
                    if (m_RSFileCount != 0)
                    {
                        receive.rSattachmentFile = GetFilesByKeyID(form_no, "RequestSupplement", m_RSFileCount);
                    }
                    if (m_RPFileCount != 0)
                    {
                        receive.rPattachmentFile = GetFilesByKeyID(form_no, "RequestPayment", m_RPFileCount);
                    }

                    ResultClass<List<tbQCS>> m_QCS = GetQCS(form_no);
                    if (m_QCS.ResultCode == "000")
                    {
                        receive.lisQCS = m_QCS.objResult;
                    }
                    else
                    {
                        resultClass.ResultCode = m_QCS.ResultCode;
                        resultClass.ResultMsg = m_QCS.ResultMsg;
                        return resultClass;
                    }

                    ResultClass<List<tbRE>> m_RE = GetRE(form_no);
                    if (m_RE.ResultCode == "000")
                    {
                        receive.lisRE = m_RE.objResult;
                    }
                    else
                    {
                        resultClass.ResultCode = m_RE.ResultCode;
                        resultClass.ResultMsg = m_RE.ResultMsg;
                        return resultClass;
                    }

                    ResultClass<List<tbRS>> m_RS = GetRS(form_no);
                    if (m_RS.ResultCode == "000")
                    {
                        receive.lisRS = m_RS.objResult;
                    }
                    else
                    {
                        resultClass.ResultCode = m_RS.ResultCode;
                        resultClass.ResultMsg = m_RS.ResultMsg;
                        return resultClass;
                    }

                    ResultClass<List<tbRP>> m_RP = GetRP(form_no);
                    if (m_RS.ResultCode == "000")
                    {
                        receive.lisRP = m_RP.objResult;
                    }
                    else
                    {
                        resultClass.ResultCode = m_RP.ResultCode;
                        resultClass.ResultMsg = m_RP.ResultMsg;
                        return resultClass;
                    }
                    receiveList.Add(receive);
                }
                resultClass.ResultCode = "000";
                resultClass.ResultMsg = "";
                resultClass.objResult = receiveList;
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "999";
                resultClass.ResultMsg = ex.Message;
                resultClass.objResult = null;
            }
            return resultClass;
        }

        public ResultClass<string> GetSeqNo(string p_Seq_Type)
        {
            ResultClass<string> resultClass = new();
            string m_ProcedureName = "Get_Seq_NO";
            string m_Seq = "";
            using DataSet m_ds = new("Data");
            try
            {
                using SqlConnection connection = new(_ADO.GetConnStr());
                using SqlCommand command = new(m_ProcedureName, connection);
                command.CommandType = CommandType.StoredProcedure;
                command.Parameters.AddWithValue("@Seq_Type", p_Seq_Type);
                command.Parameters.AddWithValue("@Seq_Code", DateTime.Now.ToString("yyMM"));

                // 建立資料配接器
                using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                {
                    // 建立資料集
                    DataSet dataSet = new DataSet();

                    // 執行命令並將結果填充到資料集中
                    adapter.Fill(dataSet);

                    // 使用資料集中的資料
                    foreach (DataRow row in dataSet.Tables[0].Rows)
                    {
                        m_Seq = row["Seq_NO"].ToString();
                    }
                }
                resultClass.ResultCode = "000";
                resultClass.ResultMsg = "";
                resultClass.objResult = m_Seq;

            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "999";
                resultClass.ResultMsg = ex.Message;
                resultClass.objResult = null;
            }
            return resultClass;
        }

        public attachmentFile[] GetFilesByKeyID(string p_Key, string p_Upload_Type, Int32 p_fileCount)
        {
            Common _Comm = new();
            attachmentFile[] m_attachmentFiles = new attachmentFile[p_fileCount];
            try
            {
                using (SqlConnection connection = new SqlConnection(_ADO.GetConnStr()))
                {
                    connection.Open();

                    using (SqlCommand command = connection.CreateCommand())
                    {
                        command.CommandText = "SELECT * FROM tbFiles WHERE KeyID = @KeyID  and Upload_Type=@Upload_Type";
                        command.Parameters.AddWithValue("@KeyID", p_Key);
                        command.Parameters.AddWithValue("@Upload_Type", p_Upload_Type);

                        Int32 m_Count = 0;
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            if (reader.HasRows)
                            {
                                while (reader.Read())
                                {
                                    byte[] blob = (byte[])reader["file_body_encode"];
                                    string base64String = Convert.ToBase64String(blob);
                                    attachmentFile m_attachmentFile = new();
                                    m_attachmentFile.file_body_encode = _Comm.DecompressFile(base64String);
                                    m_attachmentFile.file_size = reader["file_size"].ToString();
                                    m_attachmentFile.file_index = reader["file_index"].ToString();
                                    m_attachmentFile.file_name = reader["file_name"].ToString();
                                    m_attachmentFile.content_type = reader["content_type"].ToString();
                                    m_attachmentFile.TransactionId = reader["TransactionId"].ToString();
                                    m_attachmentFiles[m_Count] = (m_attachmentFile);
                                    m_Count++;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return m_attachmentFiles;
        }

        public ResultClass<attachmentFile> GetFile(string Key, string Type, string file_index = "")
        {
            Common _Comm = new();
            ResultClass<attachmentFile> resultClass = new();
            attachmentFile m_attachmentFile = new();
            try
            {
                string m_TypeName = "";
                switch (Type)
                {
                    case "R":
                        m_TypeName = "Receive";
                        break;
                    case "RE":
                        m_TypeName = "RequestforExam";
                        break;
                    case "RS":
                        m_TypeName = "RequestSupplement";
                        break;
                    case "RP":
                        m_TypeName = "RequestPayment";
                        break;
                }


                using (SqlConnection connection = new SqlConnection(_ADO.GetConnStr()))
                {
                    connection.Open();

                    using (SqlCommand command = connection.CreateCommand())
                    {
                        command.CommandText = "SELECT * FROM tbFiles WHERE KeyID = @KeyID and  Upload_Type = @Upload_Type and file_index=@file_index ";
                        command.Parameters.AddWithValue("@KeyID", Key);
                        command.Parameters.AddWithValue("@Upload_Type", m_TypeName);

                        if (file_index != "")
                        {
                            command.Parameters.AddWithValue("@file_index", file_index);
                        }

                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                byte[] blob = (byte[])reader["file_body_encode"];
                                string base64String = Convert.ToBase64String(blob);
                                m_attachmentFile.file_body_encode = _Comm.DecompressFile(base64String);
                                m_attachmentFile.file_size = reader["file_size"].ToString();
                                m_attachmentFile.file_index = reader["file_index"].ToString();
                                m_attachmentFile.file_name = reader["file_name"].ToString();
                                m_attachmentFile.content_type = reader["content_type"].ToString();
                            }
                        }
                    }
                }
                resultClass.ResultCode = "000";
                resultClass.ResultMsg = "";
                resultClass.objResult = m_attachmentFile;
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "999";
                resultClass.ResultMsg = ex.Message;
                resultClass.objResult = null;
            }
            return resultClass;
        }

        public ResultClass<int> InsertFile(string p_Key, string p_Upload_Type, string p_User, string p_transactionId, attachmentFile[] p_attachmentFiles)
        {
            Common _Comm = new();
            ResultClass<int> resultClass = new();
            int m_Execut = 0;
            try
            {
                using SqlConnection conn = new SqlConnection(_ADO.GetConnStr());
                // 開啟資料庫連線
                conn.Open();
                SqlTransaction transaction = conn.BeginTransaction();
                try
                {
                    foreach (attachmentFile file in p_attachmentFiles)
                    {
                        // 建立 SQL 命令
                        string m_SQL = " INSERT INTO tbFiles (KeyID,Upload_Type,file_index,file_body_encode,file_size,content_type,file_name,Adduser,transactionId)  ";
                        m_SQL += "  VALUES (@KeyID,@Upload_Type,@file_index,@file_body_encode,@file_size,@content_type,@file_name,@Adduser,@transactionId)  ";
                        using SqlCommand command = new SqlCommand(m_SQL, conn, transaction);
                        // 設定參數
                        string base64String = _Comm.CompressFile(file.file_body_encode);
                        byte[] imageBytes = Convert.FromBase64String(base64String);
                        command.Parameters.AddWithValue("@KeyID", p_Key);
                        command.Parameters.AddWithValue("@Upload_Type", p_Upload_Type);

                        command.Parameters.AddWithValue("@file_index", file.file_index);
                        command.Parameters.AddWithValue("@file_body_encode", imageBytes);
                        command.Parameters.AddWithValue("@file_size", file.file_size);
                        command.Parameters.AddWithValue("@content_type", file.content_type);
                        command.Parameters.AddWithValue("@file_name", file.file_name);
                        command.Parameters.AddWithValue("@Adduser", p_User);
                        command.Parameters.AddWithValue("@transactionId", p_transactionId);
                        // 執行 SQL 命令
                        m_Execut += command.ExecuteNonQuery();
                    }
                    transaction.Commit();
                    // 關閉資料庫連線
                    conn.Close();
                    resultClass.ResultCode = "000";
                    resultClass.ResultMsg = "";
                    resultClass.objResult = m_Execut;
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    resultClass.ResultCode = "999";
                    resultClass.ResultMsg = ex.Message;
                    resultClass.objResult = 0;
                }
            }
            catch (Exception ex)
            {

                throw new Exception("檔案上傳失敗!!");
            }

            return resultClass;
        }

    }
}
