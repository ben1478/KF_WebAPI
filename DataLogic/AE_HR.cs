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
using System.Collections.Generic;
using System.Reflection;

namespace KF_WebAPI.DataLogic
{
    public class AE_HR
    {
        ADOData _ADO = new();
        Common _Common = new();

        public batchLeaveNew GetLate15To104(string SESSION_KEY, string Date_S, string Date_E)
        {
            batchLeaveNew ReqClass = new batchLeaveNew();
            ReqClass.CO_ID = 1;
            ReqClass.SESSION_KEY = SESSION_KEY;
            ReqClass.WF_NO = "WO" + SESSION_KEY;

            try
            {
                int iResult = 0;
                iResult = insertToLate15To104(Date_S, Date_E);
                if (iResult != 0)
                {
                    string m_SQL = "SELECT F.[ID_104],[LEAVEITEM_ID],F.[U_num],[Date_S]  ,[Date_E]  , " +
                            " case when M2.U_arrive_date < F.[Date_S] then case when M2.U_leave_date < F.[Date_S] then 2 else M2.ID_104 end  else 2  end " +
                        "  AGENT_IDS " +
                        " FROM[dbo].[Late15To104] F LEFT JOIN User_M M1 on F.u_num = M1.u_num " +
                        " LEFT JOIN User_M M2 on M1.U_agent_num = M2.U_num " +
                        " where[SESSION_KEY_104] is null and   cast( [Date_S] as datetime) between @Date_S+' 00:00' and @Date_E +' 23:59' ";

                    List<SqlParameter> Params = new List<SqlParameter>()
                    {
                        new SqlParameter() {ParameterName = "@Date_S", SqlDbType = SqlDbType.VarChar, Value= Date_S},
                         new SqlParameter() {ParameterName = "@Date_E", SqlDbType = SqlDbType.VarChar, Value= Date_E}
                    };
                    DataTable m_RtnDT = new("Data");
                    m_RtnDT = _ADO.ExecuteQuery(m_SQL, Params);
                    List<LEAVE_DATA> m_lisLEAVE_DATA = new List<LEAVE_DATA>();
                    foreach (DataRow dr in m_RtnDT.Rows)
                    {
                        LEAVE_DATA m_LEAVE_DATA =
                        new LEAVE_DATA(dr["ID_104"].ToString(), dr["LEAVEITEM_ID"].ToString(), dr["Date_S"].ToString(), dr["Date_E"].ToString(), dr["AGENT_IDS"].ToString(), "系統自動產生遲到假");
                        m_lisLEAVE_DATA.Add(m_LEAVE_DATA);
                    }
                    ReqClass.LEAVE_DATA = m_lisLEAVE_DATA.ToArray();
                }
                
            }
            catch
            {
                throw;
            }
            return ReqClass;
           

               


            
          
        }


        public int insertToLate15To104( string Date_S, string Date_E)
        {
            
            string m_SQL = "insert into Late15To104 ([ID_104],[LEAVEITEM_ID],[U_num],[Date_S],[Date_E])" +
                " select F.[ID_104],[LEAVEITEM_ID],F.[U_num],format([Date_S],'yyyy/MM/dd HH:mm'),format([Date_E],'yyyy/MM/dd HH:mm') " +
                " from [fun_LateOver15_To104](@Date_S,@Date_E)F  " +
                " Left join  (select * from Flow_rest R  where  del_tag = '0' AND FR_cancel<>'Y' and format( FR_date_begin,'yyyy/MM/dd')   " +
                " between @Date_S and @Date_E and FR_kind <> 'FRK021' )R on F.[U_num]=R.FR_U_num and format([Date_E],'yyyy/MM/dd HH:mm') between R.FR_date_begin and R.FR_date_end  " +
                "   where R.FR_date_begin is null  ";
            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("@Date_S", Date_S));
            parameters.Add(new SqlParameter("@Date_E", Date_E));

          return  _ADO.ExecuteNonQuery(m_SQL, parameters);

        }

        public string GetAE_Flow_rest( string FR_kind)
        {
             string m_SQL = "SELECT  F.FR_ID WF_NO,F.FR_cknum SESSION_KEY,M.ID_104,  format( F.FR_date_begin,'yyyy/MM/dd HH:mm') FR_date_begin," +
                "format( F.FR_date_end,'yyyy/MM/dd HH:mm')  FR_date_end, F.FR_ot_compen," +
                    "  /*[CALENDAR_LEAVE_ID]:4例假日,類別為7*/  " +
                    "  case when [CALENDAR_LEAVE_ID]='4' then '7' else " +
                    "  case when FR_ot_compen ='0' then '3' when FR_ot_compen ='1' then '2' end   " +
                    " end PAY_TYPE " +
                    "  ,'0' IS_MEAL,'0' IS_CARDMATCH ,FR_note ,cancel_num,M.U_name,LEAVEITEM_ID, " +
                    " /*假設代理人到職日大於請假日,就指定主管*/ " +
                    "isnull(M1.ID_104, " +
                    " case when M2.U_arrive_date>F.FR_date_begin then 2 else M2.ID_104 end " +
                    " ) AGENT_IDS " +
                    "FROM Flow_rest F LEFT JOIN User_M M on F.FR_U_num=M.U_num " +
                    " LEFT JOIN calendar_day_104 C on format( FR_date_begin,'yyyy-MM-dd') = format( [CALENDAR_DATE],'yyyy-MM-dd')  " +
                    " LEFT JOIN Leaveitem_104 L on F.FR_kind=L.AE_FR_kind LEFT JOIN User_M M1 on F.FR_step_01_num=M1.U_num LEFT JOIN User_M M2 on M.U_agent_num=M2.U_num " +
                    "WHERE F.del_tag = '0' AND FR_sign_type = 'FSIGN002' and SESSION_KEY_104 is null and M.ID_104 is not null and M.U_leave_date is null  and FR_ID not in (149618,151692,152802,152431) AND format( FR_date_begin,'yyyy-MM-dd') >'2024-01-01' ";
            if (FR_kind == "TEST")
            {
                m_SQL += "  and M.U_num='K0108' AND FR_kind IN ('FRK021')  AND format( FR_date_begin,'yyyy-MM-dd') >'2024-12-20'   ";
            }
            else
            {
                if (FR_kind == "FRK005")//特休
                {
                    m_SQL += "AND FR_kind IN ('FRK005') ";
                }
                else if (FR_kind == "FRK004")//補休
                {
                    m_SQL += "AND FR_kind IN ('FRK004') ";
                }
                else if (FR_kind == "FRK021")//加班
                {
                    m_SQL += "  AND FR_kind IN ('FRK021')   ";
                }
                else if (FR_kind == "Other1")//喪假、產假、陪產假、產檢假
                {
                    m_SQL += "  AND FR_kind IN ('FRK007','FRK022','FRK023','FRK008','FRK009','FRK010')   ";
                }
                else//非特休,補休,加班,喪假、產假、陪產假、產檢假
                {
                    m_SQL += "AND FR_kind not IN ('FRK005','FRK004','FRK021','FRK007','FRK022','FRK023','FRK008','FRK009','FRK010')  ";
                }
                m_SQL += " AND  format( F.FR_step_HR_date,'yyyy-MM-dd') between @Date_S and @Date_E  ";
            }

            m_SQL += " ORDER BY F.FR_U_num,  F.FR_cknum ";
            return m_SQL;

        }

        public int ModifyAE_SESSION_KEY(string TableName, string SESSION_KEY, string FR_kind, string Date_S, string Date_E)
        {
            int m_UPDCount = 0;
            try
            {
                Date_S = ConvertYYYYMMDD(Date_S);
                Date_E = ConvertYYYYMMDD(Date_E);
                List<SqlParameter> Params = new List<SqlParameter>()
                {
                    new SqlParameter() {ParameterName = "@SESSION_KEY", SqlDbType = SqlDbType.VarChar, Value= SESSION_KEY},
                    new SqlParameter() {ParameterName = "@Date_S", SqlDbType = SqlDbType.VarChar, Value= Date_S},
                     new SqlParameter() {ParameterName = "@Date_E", SqlDbType = SqlDbType.VarChar, Value= Date_E}
                };
                string m_SQL = "update  " + TableName + "  set SESSION_KEY_104=@SESSION_KEY ";
                if (TableName == "Late15To104")
                {
                    m_SQL += " WHERE cast([Date_S] as datetime) between @Date_S+' 00:00' and @Date_E +' 23:59' and SESSION_KEY_104=@SESSION_KEY   ";
                }
                else
                {
                    m_SQL += " WHERE FR_ID in  ";

                    m_SQL += " (SELECT  F.FR_ID " +
                        "FROM Flow_rest F LEFT JOIN User_M M on F.FR_U_num=M.U_num LEFT JOIN Leaveitem_104 L on F.FR_kind=L.AE_FR_kind LEFT JOIN User_M M1 on F.FR_step_01_num=M1.U_num LEFT JOIN User_M M2 on M.U_agent_num=M2.U_num " +
                        "WHERE F.del_tag = '0' AND FR_sign_type = 'FSIGN002' and SESSION_KEY_104 is null and M.ID_104 is not null and M.U_leave_date is null and FR_ID not in (149618,151692,152802,152431) ";

                    if (FR_kind == "TEST")
                    {
                        m_SQL += "  and M.U_num='K0108' AND FR_kind IN ('FRK021')  AND format( FR_date_begin,'yyyy-MM-dd') >'2024-12-20'   ";
                    }
                    else
                    {
                        if (FR_kind == "FRK005")//特休
                        {
                            m_SQL += "AND FR_kind IN ('FRK005') ";
                        }
                        else if (FR_kind == "FRK004")//補休
                        {
                            m_SQL += "AND FR_kind IN ('FRK004') ";
                        }
                        else if (FR_kind == "FRK021")//加班
                        {
                            m_SQL += "  AND FR_kind IN ('FRK021')    ";
                        }
                        else if (FR_kind == "Other1")//喪假、產假、陪產假、產檢假
                        {
                            m_SQL += "  AND FR_kind IN ('FRK007','FRK022','FRK023','FRK008','FRK009','FRK010')   ";
                        }
                        else//非特休,補休,加班,喪假、產假、陪產假、產檢假
                        {
                            m_SQL += "AND FR_kind not IN ('FRK005','FRK004','FRK021','FRK007','FRK022','FRK023','FRK008','FRK009','FRK010')  ";
                        }
                        m_SQL += "AND format( F.FR_step_HR_date,'yyyy-MM-dd') between @Date_S and @Date_E  ";
                    }
                    m_SQL += " )";
                }





                m_UPDCount = _ADO.ExecuteNonQuery(m_SQL, Params);


            }
            catch
            {
                throw;
            }
            return m_UPDCount;
        }

        public string ConvertYYYYMMDD(string p_Date)
        {
            string m_date = Convert.ToDateTime(p_Date).ToString("yyyy-MM-dd");

            return m_date;
        }
        public batchOtNew GetAE_OT_DATA(string SESSION_KEY, string FR_kind, string Date_S, string Date_E)
        {
            batchOtNew ReqClass = new batchOtNew();
            ReqClass.CO_ID = 1;

            ReqClass.SESSION_KEY = SESSION_KEY;
            ReqClass.WF_NO = "WO" + SESSION_KEY;
            Date_S = ConvertYYYYMMDD(Date_S);
            Date_E = ConvertYYYYMMDD(Date_E);
            try
            {
                List<SqlParameter> Params = new List<SqlParameter>()
                {
                    new SqlParameter() {ParameterName = "@Date_S", SqlDbType = SqlDbType.VarChar, Value= Date_S},
                     new SqlParameter() {ParameterName = "@Date_E", SqlDbType = SqlDbType.VarChar, Value= Date_E}
                };
                string m_SQL = GetAE_Flow_rest(FR_kind);
                DataTable m_RtnDT = new("Data");
                m_RtnDT = _ADO.ExecuteQuery(m_SQL, Params);
                List<OT_DATA> m_lisOT_DATA = new List<OT_DATA>();
                foreach (DataRow dr in m_RtnDT.Rows)
                {
                    OT_DATA m_OT_DATA =
                    new OT_DATA(dr["ID_104"].ToString(), dr["FR_date_begin"].ToString(), dr["FR_date_end"].ToString(), dr["PAY_TYPE"].ToString(), "0", "0", dr["FR_note"].ToString());
                    m_lisOT_DATA.Add(m_OT_DATA);
                }
                ReqClass.OT_DATA = m_lisOT_DATA.ToArray();
            }
            catch
            {
                throw;
            }
            return ReqClass;
        }

        public batchLeaveNew GetAE_LEAVE_DATA(string SESSION_KEY, string FR_kind, string Date_S, string Date_E)
        {
            batchLeaveNew ReqClass = new batchLeaveNew();
            ReqClass.CO_ID = 1;
            ReqClass.SESSION_KEY = SESSION_KEY;
            ReqClass.WF_NO = "WO" + SESSION_KEY;

            Date_S = ConvertYYYYMMDD(Date_S);
            Date_E = ConvertYYYYMMDD(Date_E);
            try
            {
                List<SqlParameter> Params = new List<SqlParameter>()
                {
                    new SqlParameter() {ParameterName = "@Date_S", SqlDbType = SqlDbType.VarChar, Value= Date_S},
                     new SqlParameter() {ParameterName = "@Date_E", SqlDbType = SqlDbType.VarChar, Value= Date_E}
                };
                string m_SQL = GetAE_Flow_rest(FR_kind);
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

        public void InsertExternal_API_Log(External_API_Log p_External_API_Log)
        {

            try
            {
                string TableName = "External_API_Log";
                External_API_Log[] arrExternal_API_Log = new External_API_Log[1];
                arrExternal_API_Log[0] = p_External_API_Log;

                _ADO.DataTableToSQL(TableName, arrExternal_API_Log, _ADO.ConnStr);
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
