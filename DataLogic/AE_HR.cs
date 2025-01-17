using KF_WebAPI.FunctionHandler;
using System.Data;
using Microsoft.Data.SqlClient;
using KF_WebAPI.BaseClass;
using Newtonsoft.Json;
using System.Data.Common;
using System.Transactions;
using KF_WebAPI.BaseClass.Max104;
using KF_WebAPI.BaseClass.AE;
using System;

namespace KF_WebAPI.DataLogic
{
    public class AE_HR
    {
        ADOData _ADO = new();
        Common _Common = new();

        public DataTable GetAE_OTData(string form_no)
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

        public batchLeaveNew GetAE_LEAVE_DATA(string SESSION_KEY, string FR_kind, string Date_S, string Date_E)
        {
            batchLeaveNew ReqClass = new batchLeaveNew();
            ReqClass.CO_ID = 1;
            
            ReqClass.SESSION_KEY = SESSION_KEY;
            ReqClass.WF_NO = "WO"+ SESSION_KEY;
            try
            {
                List<SqlParameter> Params = new List<SqlParameter>()
                {
                    new SqlParameter() {ParameterName = "@Date_S", SqlDbType = SqlDbType.VarChar, Value= Date_S},
                     new SqlParameter() {ParameterName = "@Date_E", SqlDbType = SqlDbType.VarChar, Value= Date_E}
                };
                string m_SQL = "SELECT top 100  F.FR_ID WF_NO,F.FR_cknum SESSION_KEY,M.ID_104,  format( F.FR_date_begin,'yyyy/MM/dd HH:mm') FR_date_begin,format( F.FR_date_end,'yyyy/MM/dd HH:mm')  FR_date_end," +
                    "F.FR_ot_compen, case when FR_ot_compen ='0' then '3' when FR_ot_compen ='1' then '2' end PAY_TYPE ,'0' IS_MEAL,'0' IS_CARDMATCH ,FR_note ,cancel_num,M.U_name,LEAVEITEM_ID, " +
                    " /*假設代理人到職日大於請假日,就指定主管*/ " +
                    "isnull(M1.ID_104, " +
                    " case when M2.U_arrive_date>F.FR_date_begin then 2 else M2.ID_104 end " +
                    " ) AGENT_IDS " +
                    "FROM Flow_rest F LEFT JOIN User_M M on F.FR_U_num=M.U_num LEFT JOIN Leaveitem_104 L on F.FR_kind=L.AE_FR_kind LEFT JOIN User_M M1 on F.FR_step_01_num=M1.U_num LEFT JOIN User_M M2 on M.U_agent_num=M2.U_num " +
                    "WHERE F.del_tag = '0' AND FR_sign_type = 'FSIGN002' and SESSION_KEY_104 is null and M.ID_104 is not null and M.U_leave_date is null  and FR_ID not in (149618,151692,152802,152431) ";
                //測試用
                 //m_SQL += " and M.U_num='K0101' ";

                if (FR_kind == "FRK005")
                {
                    m_SQL += "AND FR_kind IN ('FRK005') ";
                   
                }
                else
                {
                    m_SQL += "AND FR_kind not IN ('FRK005','FRK004','FRK021') ";
                }
                m_SQL += "AND format( F.FR_date_begin,'yyyy-MM-dd') between @Date_S and @Date_E  ";

                m_SQL += " ORDER BY F.FR_U_num,  F.FR_cknum ";
                DataTable m_RtnDT = new("Data");
                m_RtnDT = _ADO.ExecuteQuery(m_SQL, Params);
                List<LEAVE_DATA> m_lisLEAVE_DATA = new List<LEAVE_DATA>();
                foreach (DataRow dr in m_RtnDT.Rows)
                {
                    LEAVE_DATA m_LEAVE_DATA =
                    new LEAVE_DATA(dr["ID_104"].ToString(), dr["LEAVEITEM_ID"].ToString(), dr["FR_date_begin"].ToString(), dr["FR_date_end"].ToString(), dr["AGENT_IDS"].ToString(), dr["FR_note"].ToString());
                    m_lisLEAVE_DATA.Add(m_LEAVE_DATA);
                }
                ReqClass.LEAVE_DATA = m_lisLEAVE_DATA.ToArray();
            }
            catch
            {
                throw;
            }
            return ReqClass;
        }

        public int ModifyAE_SESSION_KEY(string TableName, string SESSION_KEY, string FR_kind, string Date_S, string Date_E)
        {
            int m_UPDCount = 0;
            try
            {
               
                List<SqlParameter> Params = new List<SqlParameter>()
                {
                    new SqlParameter() {ParameterName = "@SESSION_KEY", SqlDbType = SqlDbType.VarChar, Value= SESSION_KEY},
                    new SqlParameter() {ParameterName = "@Date_S", SqlDbType = SqlDbType.VarChar, Value= Date_S},
                     new SqlParameter() {ParameterName = "@Date_E", SqlDbType = SqlDbType.VarChar, Value= Date_E}
                };

                string m_SQL = "update  " + TableName + "  set SESSION_KEY_104=@SESSION_KEY " +
                   " WHERE FR_ID in  ";

                m_SQL += " (SELECT top 100  F.FR_ID " +
                    "FROM Flow_rest F LEFT JOIN User_M M on F.FR_U_num=M.U_num LEFT JOIN Leaveitem_104 L on F.FR_kind=L.AE_FR_kind LEFT JOIN User_M M1 on F.FR_step_01_num=M1.U_num LEFT JOIN User_M M2 on M.U_agent_num=M2.U_num " +
                    "WHERE F.del_tag = '0' AND FR_sign_type = 'FSIGN002' and SESSION_KEY_104 is null and M.ID_104 is not null and M.U_leave_date is null and FR_ID not in (149618,151692,152802,152431) ";

                //測試用
                // m_SQL += " and M.U_num='K0101' ";

                if (FR_kind == "FRK005")
                {
                    m_SQL += "AND FR_kind IN ('FRK005') ";
                }
                else
                {
                    m_SQL += "AND FR_kind not IN ('FRK005','FRK004','FRK021') ";
                }

                m_SQL += "AND format( F.FR_date_begin,'yyyy-MM-dd') between @Date_S and @Date_E  ";

                m_SQL += "ORDER BY F.FR_U_num,  F.FR_cknum )";

                m_UPDCount = _ADO.ExecuteNonQuery(m_SQL, Params);

                
            }
            catch
            {
                throw;
            }
            return m_UPDCount;
        }

        public int ModifyAE_Sign(string TableName, string SESSION_KEY)
        {
            int m_UPDCount = 0;
            try
            {
                List<SqlParameter> Params = new List<SqlParameter>()
                {
                    new SqlParameter() {ParameterName = "@SESSION_KEY", SqlDbType = SqlDbType.VarChar, Value= SESSION_KEY}
                };

                string m_SQL = "update  " + TableName + "  set isSign_104='Y' " +
                   " WHERE SESSION_KEY_104 =@SESSION_KEY and isSign_104 is null ";

              

                m_UPDCount = _ADO.ExecuteNonQuery(m_SQL, Params);


            }
            catch
            {
                throw;
            }
            return m_UPDCount;
        }


        public void InsertAE_Calendar_day(arrResultClass_104<calendar_day> APIResult,string m_Year)
        {
            
            try
            {
                string TableName = "calendar_day_104";
                List<SqlParameter> Params = new List<SqlParameter>()
                {
                    new SqlParameter() {ParameterName = "@Year", SqlDbType = SqlDbType.VarChar, Value= m_Year}
                };
                _ADO.ExecuteNonQuery("Delete FROM dbo.calendar_day_104 where year([CALENDAR_DATE])=@Year ", Params);
                _ADO.DataTableToSQL(TableName, ((calendar_day[])APIResult.data), _ADO.ConnStr);
            }
            catch
            {
                throw;
            }
           
        }


        public void InsertAE_Leaveitem(arrResultClass_104<leaveitem> APIResult)
        {

            try
            {
                string TableName = "Leaveitem_104";
             
                _ADO.ExecuteNonQuery("Delete FROM dbo.Leaveitem_104 ", new List<SqlParameter>());
                _ADO.DataTableToSQL(TableName, ((leaveitem[])APIResult.data), _ADO.ConnStr);
            }
            catch
            {
                throw;
            }

        }

    }
}
