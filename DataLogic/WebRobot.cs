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
                    m_SQL = "update WebRobot_M set  KeyWords=@KeyWords,RunTime_S=@RunTime_S,RunTime_E=@RunTime_E,numDelay=@numDelay,RunStatus=@RunStatus," +
                    " edit_date=SYSDATETIME() where ComputerInfo=@ComputerInfo";
                }
                else
                {
                    m_SQL = "insert into WebRobot_M ([ComputerInfo],[KeyWords],[RunTime_S],[RunTime_E],[numDelay],[RunStatus],edit_date)" +
                " values(@ComputerInfo,@KeyWords,@RunTime_S,@RunTime_E,@numDelay,@RunStatus,SYSDATETIME())  ";
                }
                
                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@ComputerInfo", p_WebRobot_M.ComputerInfo));
                parameters.Add(new SqlParameter("@KeyWords", p_WebRobot_M.KeyWords));
                parameters.Add(new SqlParameter("@RunTime_S", p_WebRobot_M.RunTime_S));
                parameters.Add(new SqlParameter("@RunTime_E", p_WebRobot_M.RunTime_E));
                parameters.Add(new SqlParameter("@numDelay", p_WebRobot_M.numDelay));
                parameters.Add(new SqlParameter("@RunStatus", p_WebRobot_M.RunStatus));


                m_Execut= _ADO.ExecuteNonQuery(m_SQL, parameters);
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
                List<string> proxies = new List<string>();

                try
                {
                    var result = JsonConvert.DeserializeObject<ProxyResponse>(json);
                    string ErrMessage = "";
                    if (result != null)
                    {
                        ErrMessage = "ErrorCode：" + result.code.ToString() + ";ErrMsg" + result.msg;
                    }
                    resultClass.ResultCode = "999";
                    resultClass.ResultMsg = ErrMessage;
                    resultClass.objResult = null;
                }
                catch (Newtonsoft.Json.JsonException)
                {
                    proxies = json.Split(new char[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                }


                resultClass.ResultCode = "000";
                resultClass.ResultMsg = "";
                resultClass.objResult = proxies;
            }
            catch (Exception ex)
            {

                resultClass.ResultCode = "999";
                resultClass.ResultMsg = ex.Message;
                resultClass.objResult = null;
            }

            return resultClass;
        }


    }
}
