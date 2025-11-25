using KF_WebAPI.BaseClass;
using KF_WebAPI.BaseClass.WebRobot;
using KF_WebAPI.FunctionHandler;
using Microsoft.Data.SqlClient;
using Newtonsoft.Json;
using System.Data;
using System.Text.Json;
using static Azure.Core.HttpHeader;

namespace KF_WebAPI.DataLogic
{


    public class WebRobot
    {
        ADOData _ADO = new();
        public ResultClass<int> InsertWebRobot_M(WebRobot_M p_WebRobot_M)
        {
            ResultClass<int> resultClass = new();
            int m_Execut = 0;
            try
            {
                string m_SQL = "Select * from WebRobot_M where ComputerInfo=@ComputerInfo ";

                List<SqlParameter> Params = new List<SqlParameter>()
                {
                    new SqlParameter() {ParameterName = "@ComputerInfo", SqlDbType = SqlDbType.VarChar, Value= p_WebRobot_M.ComputerInfo}
                };

                DataTable m_RtnDT = new("Data");
                m_RtnDT = _ADO.ExecuteQuery(m_SQL, Params);
                if (m_RtnDT.Rows.Count > 0)
                {
                    m_SQL = "update WebRobot_M set  KeyWords=@KeyWords,RunTime_S=@RunTime_S,RunTime_E=@RunTime_E,numDelay=@numDelay,RunStatus=@RunStatus,Version=@Version," +
                    " edit_date=SYSDATETIME(),RunTime=@RunTime where ComputerInfo=@ComputerInfo";
                }
                else
                {
                    m_SQL = "insert into WebRobot_M ([ComputerInfo],[KeyWords],[RunTime_S],[RunTime_E],[numDelay],[RunStatus],edit_date,RunTime,Version)" +
                " values(@ComputerInfo,@KeyWords,@RunTime_S,@RunTime_E,@numDelay,@RunStatus,SYSDATETIME(),@RunTime,@Version)  ";
                }
                if (p_WebRobot_M.RunTime is null)
                {
                    p_WebRobot_M.RunTime = "";
                }

                if (p_WebRobot_M.RunTime_S is null)
                {
                    p_WebRobot_M.RunTime_S = 0;
                }
                if (p_WebRobot_M.RunTime_E is null)
                {
                    p_WebRobot_M.RunTime_E = 0;
                }
                if (p_WebRobot_M.Version is null)
                {
                    p_WebRobot_M.Version = "";
                }


                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@ComputerInfo", p_WebRobot_M.ComputerInfo));
                parameters.Add(new SqlParameter("@KeyWords", p_WebRobot_M.KeyWords));
                parameters.Add(new SqlParameter("@RunTime_S", p_WebRobot_M.RunTime_S));
                parameters.Add(new SqlParameter("@RunTime_E", p_WebRobot_M.RunTime_E));
                parameters.Add(new SqlParameter("@numDelay", p_WebRobot_M.numDelay));
                parameters.Add(new SqlParameter("@RunStatus", p_WebRobot_M.RunStatus));
                parameters.Add(new SqlParameter("@Version", p_WebRobot_M.Version));
                parameters.Add(new SqlParameter("@RunTime", p_WebRobot_M.RunTime));


                m_Execut = _ADO.ExecuteNonQuery(m_SQL, parameters);
                resultClass.ResultCode = "000";
                resultClass.ResultMsg = "";
                resultClass.objResult = m_Execut;
            }
            catch (Exception ex)
            {

                resultClass.ResultCode = "999";
                resultClass.ResultMsg = ex.Message;
                resultClass.objResult = 0;
            }

            return resultClass;
        }

        public ResultClass<int> InsertWebRobot_D(WebRobot_D p_WebRobot_D)
        {
            ResultClass<int> resultClass = new();
            int m_Execut = 0;
            try
            {
                string m_SQL = "insert into WebRobot_D ([ComputerInfo],[Run_num],[Run_DateTime],[UserAgentInfo],[KeyWord],[Run_Result],[Position],edit_date)" +
                " values(@ComputerInfo,@Run_num,SYSDATETIME(),@UserAgentInfo,@KeyWord,@Run_Result,@Position,SYSDATETIME())  ";
                

                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@ComputerInfo", p_WebRobot_D.ComputerInfo));
                parameters.Add(new SqlParameter("@Run_num", DateTime.Now.ToString("yyyyMMddHHmmssfff")));
                parameters.Add(new SqlParameter("@UserAgentInfo", p_WebRobot_D.UserAgentInfo));
                parameters.Add(new SqlParameter("@KeyWord", p_WebRobot_D.KeyWord));
                parameters.Add(new SqlParameter("@Run_Result", p_WebRobot_D.Run_Result));
                parameters.Add(new SqlParameter("@Position", p_WebRobot_D.Position));


                m_Execut = _ADO.ExecuteNonQuery(m_SQL, parameters);
                resultClass.ResultCode = "000";
                resultClass.ResultMsg = "";
                resultClass.objResult = m_Execut;
            }
            catch (Exception ex)
            {

                resultClass.ResultCode = "999";
                resultClass.ResultMsg = ex.Message;
                resultClass.objResult = 0;
            }

            return resultClass;
        }

        private const string API_LINK_TEMPLATE = "https://tq.lunaproxy.com/getflowip?neek=1807058&num=30&regions=tw&ip_si=1&level=1&sb=";
        //舊帳號
        //private const string API_LINK_TEMPLATE = "https://tq.lunaproxy.com/getflowip?neek=1804199&num=1&regions=tw&ip_si=1&level=1&sb=";
        public async Task<ResultClass<List<string>>> GetProxyFlowAsync()
        {
            ResultClass<List<string>> resultClass = new();
           
            try
            {
                string json = "";
                using (HttpClient client = new HttpClient())
                {
                    json = await client.GetStringAsync(API_LINK_TEMPLATE);
                }

                try
                {
                    var result = JsonConvert.DeserializeObject<ProxyResponse>(json);
                    string ErrMessage = "";
                    if (result != null)
                    {
                        ErrMessage = "ErrorCode：" + result.code.ToString() + ";ErrMsg" + result.msg;
                    }
                }
                catch (Newtonsoft.Json.JsonException)
                {
                    resultClass.ResultCode = "000";
                    resultClass.objResult = json.Split(new char[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                }
              
            }
            catch (Exception ex)
            {

                resultClass.ResultCode = "999";
                resultClass.ResultMsg = ex.Message;
            }

            return resultClass;
        }


        public ResultClass<RemoteAction> GetRemoteAction(string ComputerInfo)
        {
            ResultClass<RemoteAction> resultClass = new();
            RemoteAction m_RemoteAction=new RemoteAction();
            try
            {
                string m_SQL = "Select TOP 1 * from WebRobot_RemoteAction where ComputerInfo=@ComputerInfo and ActionResult=@ActionResult ";

                List<SqlParameter> Params = new List<SqlParameter>()
                {
                    new SqlParameter() {ParameterName = "@ComputerInfo", SqlDbType = SqlDbType.VarChar, Value= ComputerInfo}
                    , new SqlParameter() {ParameterName = "@ActionResult", SqlDbType = SqlDbType.VarChar, Value= "N"}
                };
                DataTable m_RtnDT = new("Data");
                m_RtnDT = _ADO.ExecuteQuery(m_SQL, Params);
                
                if (m_RtnDT.Rows.Count > 0)
                {

                    m_RemoteAction.UpdRunTimes = m_RtnDT.Rows[0]["UpdRunTimes"].ToString().Split(";").ToList() ;
                    m_RemoteAction.UpdKeyWords = m_RtnDT.Rows[0]["UpdKeyWords"].ToString().Split(";").ToList();

                    m_RemoteAction.ReStart = m_RtnDT.Rows[0]["ReStart"].ToString();
                    m_RemoteAction.Shutdown = m_RtnDT.Rows[0]["Shutdown"].ToString();

                    string ActionKey = m_RtnDT.Rows[0]["ActionKey"].ToString();
                    m_SQL = "update WebRobot_RemoteAction set ActionResult=@ActionResult where ActionKey=@ActionKey and ComputerInfo=@ComputerInfo   ";

                    var parameters = new List<SqlParameter>();
                    parameters.Add(new SqlParameter("@ComputerInfo", ComputerInfo));
                    parameters.Add(new SqlParameter("@ActionKey", ActionKey));
                    parameters.Add(new SqlParameter("@ActionResult", "Y"));
                    Int32 m_Execut = _ADO.ExecuteNonQuery(m_SQL, parameters);
                }
                

                resultClass.ResultCode = "000";
                resultClass.objResult = m_RemoteAction;
            }
            catch (Exception ex)
            {

                resultClass.ResultCode = "999";
                resultClass.ResultMsg = ex.Message;
            }

            return resultClass;
        }


        public ResultClass<List<string>> GetSysKeyWords()
        {
            ResultClass<List<string>> resultClass = new();
            List<string> m_KeyWords = new List<string>();
            try
            {
                string m_SQL = "select item_D_code,item_D_name KeyWord from Item_list where item_M_code=@item_M_code and item_D_type='Y' and show_tag=0 and del_tag=0 order by item_D_code";

                List<SqlParameter> Params = new List<SqlParameter>()
                {
                     new SqlParameter() {ParameterName = "@item_M_code", SqlDbType = SqlDbType.VarChar, Value= "KeyWords"}
                };
                DataTable m_RtnDT = new("Data");
                string KewWordResult = "";

                m_RtnDT = _ADO.ExecuteQuery(m_SQL, Params);
                if (m_RtnDT.Rows.Count > 0)
                {
                    foreach (DataRow item in m_RtnDT.Rows)
                    {
                        if (KewWordResult == "")
                        {
                            KewWordResult += item["KeyWord"].ToString() ;
                        }
                        else
                        {
                            KewWordResult += ","+ item["KeyWord"].ToString() ;
                        }
                       
                    }
                }

                m_KeyWords= KewWordResult.Split(",").ToList();
                resultClass.ResultCode = "000";
                resultClass.objResult = m_KeyWords;
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "999";
                resultClass.ResultMsg = ex.Message;
            }

            return resultClass;
        }


    }
}
