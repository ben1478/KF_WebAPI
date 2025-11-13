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
