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
                resultClass.objResult = 0;
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
                string m_SQL = "insert into WebRobot_D ([ComputerInfo],[Run_num],[Run_DateTime],[UserAgentInfo],[KeyWord],[Run_Result],edit_date)" +
                " values(@ComputerInfo,@Run_num,SYSDATETIME(),@UserAgentInfo,@KeyWord,@Run_Result,SYSDATETIME())  ";
                

                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@ComputerInfo", p_WebRobot_D.ComputerInfo));
                parameters.Add(new SqlParameter("@Run_num", DateTime.Now.ToString("yyyyMMddHHmmssfff")));
                parameters.Add(new SqlParameter("@UserAgentInfo", p_WebRobot_D.UserAgentInfo));
                parameters.Add(new SqlParameter("@KeyWord", p_WebRobot_D.KeyWord));
                parameters.Add(new SqlParameter("@Run_Result", p_WebRobot_D.Run_Result));


                m_Execut = _ADO.ExecuteNonQuery(m_SQL, parameters);
                resultClass.ResultCode = "000";
                resultClass.ResultMsg = "";
                resultClass.objResult = 0;
            }
            catch (Exception ex)
            {

                resultClass.ResultCode = "999";
                resultClass.ResultMsg = ex.Message;
                resultClass.objResult = 0;
            }

            return resultClass;
        }

    }
}
