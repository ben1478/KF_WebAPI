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
                " Left join  (select * from Flow_rest R  where  del_tag = '0' AND FR_cancel<>'Y' and FR_date_begin   " +
                " between @Date_S+' 00:00' and @Date_E+' 23:59' and FR_kind <> 'FRK021' )R on F.[U_num]=R.FR_U_num and format([Date_E],'yyyy/MM/dd HH:mm') between R.FR_date_begin and R.FR_date_end  " +
                "   where R.FR_date_begin is null and u_susp_date is null  ";
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
                else if (FR_kind == "Other1")//婚假、產假、陪產假、產檢假、喪假(3,6,8天)
                {
                    m_SQL += "  AND FR_kind IN ('FRK006','FRK007','FRK022','FRK023','FRK008','FRK009','FRK010','FRK019')   ";
                }
                else//非特休,補休,加班,婚假、產假、陪產假、產檢假、喪假(3,6,8天)
                {
                    m_SQL += "AND FR_kind not IN ('FRK006','FRK005','FRK004','FRK021','FRK007','FRK022','FRK023','FRK008','FRK009','FRK010','FRK019')  ";
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
                    m_SQL += " WHERE cast([Date_S] as datetime) between @Date_S+' 00:00' and @Date_E +' 23:59'   ";
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
                        else if (FR_kind == "Other1")//婚假、產假、陪產假、產檢假、喪假(3,6,8天)
                        {
                            m_SQL += "  AND FR_kind IN ('FRK006','FRK007','FRK022','FRK023','FRK008','FRK009','FRK010')   ";
                        }
                        else//非特休,補休,加班,婚假、產假、陪產假、產檢假、喪假(3,6,8天)
                        {
                            m_SQL += "AND FR_kind not IN ('FRK006','FRK005','FRK004','FRK021','FRK007','FRK022','FRK023','FRK008','FRK009','FRK010')  ";
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


        public void InsertAE_Calendar_day(arrResultClass_104<calendar_day> APIResult,string m_Year, string UserID)
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
                InsertHolidays(m_Year, UserID);

            }
            catch
            {
                throw;
            }
        }


        public void InsertHolidays( string m_Year, string UserID)
        {

            try
            {
                //Holidays
                List<SqlParameter> Params = new List<SqlParameter>()
                {
                    new SqlParameter() {ParameterName = "@Year", SqlDbType = SqlDbType.VarChar, Value= m_Year}
                };
                _ADO.ExecuteNonQuery("delete Holidays where [HDate] in (SELECT [HDate] FROM Holidays_D where year(convert(datetime,[HDate]))=@Year and [Influence] is null) ", Params);

                Params = new List<SqlParameter>()
                {
                    new SqlParameter() {ParameterName = "@Year", SqlDbType = SqlDbType.VarChar, Value= m_Year},
                    new SqlParameter() {ParameterName = "@UserID", SqlDbType = SqlDbType.VarChar, Value= UserID}
                };
                string SQL = " INSERT INTO[Holidays] ";
                SQL += " SELECT format([CALENDAR_DATE],'yyyy/MM/dd'),SYSDATETIME(),@UserID ";
                SQL += " FROM calendar_day_104 where format([CALENDAR_DATE],'yyyy')= @Year and[CALENDAR_LEAVE_ID] <> 1 ";
                _ADO.ExecuteNonQuery(SQL, Params);

                //Holidays_D
                 Params = new List<SqlParameter>()
                {
                    new SqlParameter() {ParameterName = "@Year", SqlDbType = SqlDbType.VarChar, Value= m_Year}
                };
                _ADO.ExecuteNonQuery("delete Holidays_D where year(convert(datetime,[HDate]))=@Year and [Influence] is null ", Params);

                Params = new List<SqlParameter>()
                {
                    new SqlParameter() {ParameterName = "@Year", SqlDbType = SqlDbType.VarChar, Value= m_Year},
                    new SqlParameter() {ParameterName = "@UserID", SqlDbType = SqlDbType.VarChar, Value= UserID}
                };

                SQL = " INSERT INTO[Holidays_D]([HDate],[CreateDate],[CreateUser]) ";
                SQL += " SELECT format([CALENDAR_DATE],'yyyy/MM/dd'),SYSDATETIME(),@UserID ";
                SQL += " FROM calendar_day_104 where format([CALENDAR_DATE],'yyyy')= @Year and[CALENDAR_LEAVE_ID] <> 1 ";
                _ADO.ExecuteNonQuery(SQL, Params);

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


        public void InsertBankInfom(arrResultClass_104<emp_bank> APIResult)
        {

            try
            {
                string TableName = "BankInfo_104";

                _ADO.ExecuteNonQuery("Delete FROM dbo.BankInfo_104 ", new List<SqlParameter>());
                _ADO.DataTableToSQL(TableName, ((emp_bank[])APIResult.data), _ADO.ConnStr);
            }
            catch
            {
                throw;
            }

        }

        /// <summary>
        /// 出勤紀錄查詢
        /// </summary>
        /// <param name="YM"></param>
        /// <param name="AttStatus"></param>
        /// <param name="User_Num"></param>
        /// <param name="Type">1:個人查詢 2:部門整體查詢</param>
        /// <returns></returns>
        public List<Attendance_per_res> Attendance_Query(string YM, int? AttStatus, string User_Num,string Type,string U_BC)
        {
            try
            {
                //觸發新指紋機資料輸入
                if (Type == "1")
                {
                    AttendanceCardUpd(YM, User_Num);
                }

                var T_SQL = @"SELECT CASE WHEN ad.[attendance_date]= '2024/12/20' AND U.U_BC='BC0100' THEN 20 
                              WHEN ad.[attendance_date]= '2024/12/20' AND (U.U_BC like 'BC08%' OR U.U_BC='BC0900') THEN 10
                              WHEN ad.[attendance_date]= '2025/06/13' AND U.U_BC='BC0100' THEN 15 
                              WHEN ad.[attendance_date]= '2025/06/13' AND U.U_BC like 'BC08%' THEN 10
                              WHEN ad.[attendance_date]= '2025/12/19' AND U.U_BC='BC0100' THEN 20
                              WHEN ad.[attendance_date]= '2025/12/19' AND U.U_BC like 'BC08%' THEN 15
                              WHEN ad.[attendance_date]= '2026/02/13' THEN 120 
                              WHEN ad.[attendance_date]= '2026/06/05' AND U.U_BC='BC0100' THEN 20
                              WHEN ad.[attendance_date]= '2026/06/05' AND U.U_BC like 'BC08%' THEN 15 ELSE 0 END EarlyMin,
                              U_name,userID,ad.attendance_date,work_time,
                              CASE WHEN Holiday_NA IS NULL THEN CASE WHEN OT='N' THEN 
                              CASE WHEN isnull([work_time], '')='' THEN 0 WHEN [work_time] >= '12:00' AND [work_time] <= '13:00' THEN 180
                              WHEN [work_time] > '09:00' AND [work_time] < '12:00' THEN DATEDIFF(MINUTE, '09:00', [work_time])
                              WHEN [work_time] > '13:00' THEN DATEDIFF(MINUTE, '09:00', [work_time]) - DATEDIFF(MINUTE, '12:00', '13:00') ELSE 0 END ELSE 0 END ELSE 0 END Late,
                              CASE WHEN Holiday_NA IS NULL THEN CASE WHEN OT='N' THEN CASE WHEN isnull([work_time], '')='' THEN '未刷卡'
                              WHEN [work_time]>'09:00' THEN CASE WHEN isnull(U_num_NL, 'N')='N' THEN 
                              CASE WHEN convert(varchar, U_arrive_date, 112)= convert(varchar, CAST(ad.[attendance_date] AS datetime), 112)THEN '到職日' ELSE'遲到' END ELSE '' END ELSE '' END ELSE '加班' END ELSE Holiday_NA END work_status,
                              getoffwork_time,CASE WHEN Holiday_NA IS NULL THEN CASE WHEN OT='N' THEN CASE WHEN isnull([getoffwork_time], '')='' THEN 0
                              WHEN [getoffwork_time] <= '12:00' THEN DATEDIFF(MINUTE, [getoffwork_time], '18:00') - DATEDIFF(MINUTE, '12:00', '13:00')
                              WHEN [getoffwork_time] > '12:00' AND [getoffwork_time] <= '13:00' THEN DATEDIFF(MINUTE, '13:00', '18:00')
                              WHEN [getoffwork_time] > '13:00' AND [getoffwork_time] <= '18:00' THEN DATEDIFF(MINUTE, [getoffwork_time], '18:00') ELSE 0 END ELSE 0 END ELSE 0 END early,
                              CASE WHEN Holiday_NA IS NULL THEN CASE WHEN OT='N' THEN CASE WHEN isnull([getoffwork_time], '')='' THEN '未刷卡' 
                              WHEN [getoffwork_time]<'18:00' THEN CASE WHEN isnull(U_num_NL, 'N')='N' THEN '早退' ELSE '' END ELSE '' END ELSE '加班' END ELSE Holiday_NA END offwork_status,
                              U.U_Na,U.U_BC,isnull(RestCount, 0)RestCount 
                              FROM (
                                      SELECT FORMAT(cast(D.attendance_date AS datetime), 'yyyyMM') yyyymm,'N' OT,D.attendance_date,d.U_num userID,d.[user_name],
                                      isnull(ad.work_time, '') work_time,isnull(ad.getoffwork_time, '') getoffwork_time
                                      FROM (
                                             SELECT D.*,M.U_num,M.U_name user_name,M.U_BC FROM fun_GetWorkDays(@yyyy + '/' + @MM) D
                                             JOIN USER_M M ON 1=1 WHERE m.U_NUM NOT IN('K0001','K0002','K0000','AA999','K0999')
                                             AND (M.U_leave_date IS NULL OR FORMAT(cast(M.U_leave_date AS datetime), 'yyyyMM')=@yyyyMM)) D
                                      LEFT JOIN (SELECT userID,yyyymm,attendance_date,work_time,getoffwork_time FROM dbo.attendance WHERE yyyymm=@yyyyMM ) ad 
                                      ON D.U_num=userID AND D.attendance_date=convert(varchar,convert(datetime, @yyyy + '/' +ad.[attendance_date]), 111)
                                      UNION ALL SELECT FORMAT(cast(H.attendance_date AS datetime), 'yyyyMM')yyyymm,'Y' OT,H.attendance_date,ad.userID,
                                      ad.user_name,isnull(ad.work_time, '') work_time,isnull(ad.getoffwork_time, '') getoffwork_time FROM [dbo].[attendance] ad
                                      LEFT JOIN User_M U ON ad.userID=U.U_num
                                      JOIN (
                                             SELECT userID,yyyymm,FORMAT(cast(Hdate AS datetime), 'yyyy/MM/dd') attendance_date 
                                             FROM attendance A
                                             JOIN (
                                                    SELECT H.*,D.Influence FROM Holidays H
                                                    LEFT JOIN Holidays_D D ON H.HDate=D.HDate WHERE Holiday_kind <> 'Hk_04' OR Holiday_kind IS NULL) H
                                                    ON cast(A.yyyymm AS varchar(4))+'/'+attendance_date=FORMAT(cast(Hdate AS datetime), 'yyyy/MM/dd')
                                                    WHERE yyyymm = @yyyyMM) H ON ad.userID=H.userID AND ad.yyyymm=H.yyyymm AND cast(ad.yyyymm AS varchar(4))+'/'+ad.attendance_date=H.attendance_date
                                                    WHERE ad.yyyymm = @yyyyMM AND work_time <> '') ad
                                             LEFT JOIN (
                                                         SELECT U_PFT,U_BC,U_num,U_name,I.item_D_name U_Na,U_arrive_date FROM User_M U
                                                         LEFT JOIN (SELECT item_D_code,item_D_name FROM Item_list
                                                                    WHERE item_M_code='branch_company' AND item_D_type='Y' AND del_tag='0') I ON U.u_bc=I.item_D_code
                                                                    WHERE del_tag='0') U ON ad.userID=U.U_num
                                             LEFT JOIN ( /*特殊節日颱風天*/ 
                                                         SELECT FORMAT(convert(datetime, H.HDate), 'yyyy/MM/dd') HDate,isnull(HD.Influence, 'HD')U_BC,
                                                         isnull(HD.Holiday_kind, 'HD') Holiday_kind,isnull(Holiday_NA, '假日') Holiday_NA FROM Holidays H
                                                         LEFT JOIN Holidays_D HD ON H.HDate=HD.HDate
                                                         LEFT JOIN (
                                                                     SELECT item_D_code Holiday_kind,item_D_name Holiday_NA 
                                                                     FROM Item_list WHERE item_M_code='Holiday_kind'
                                                                     AND item_D_type='Y')I ON HD.Holiday_kind=I.Holiday_kind
                                                         WHERE HD.Holiday_kind='Hk_04' AND H.HDate like @yyyy + '%') HD ON ad.attendance_date=HD.HDate AND U.U_BC=HD.U_BC
                                             LEFT JOIN (
                                                         SELECT item_D_code U_num_NL FROM Item_list WHERE /*不計遲到人員*/ item_M_code = 'NonLate'
                                                         AND item_M_type='N')NL ON U.U_num=NL.U_num_NL
                                             LEFT JOIN (
                                                         SELECT FR_U_num,convert(varchar, FR_date_begin, 111) FR_date_S,convert(varchar, FR_date_end, 111) FR_date_E,
                                                         count(FR_U_num) RestCount FROM Flow_rest WHERE del_tag = '0' AND FR_cancel<>'Y' GROUP BY FR_U_num,
                                                         convert(varchar, FR_date_begin, 111),convert(varchar, FR_date_end, 111)) R ON ad.userID=R.FR_U_num
                                                         AND ad.attendance_date BETWEEN R.FR_date_S AND FR_date_E
                                             WHERE [work_time] < '24:00' AND Userid<>'' AND yyyymm = @yyyyMM";
                if(Type == "1")
                {
                    T_SQL += " AND userID = @user";
                }
                else
                {
                    T_SQL += " AND U.U_BC = @U_BC";
                }
                switch (AttStatus)
                {
                    case 1:
                        T_SQL += " AND (work_time>'09:00' or getoffwork_time<'18:00')";
                        break;
                    case 2:
                        T_SQL += " AND (work_time>'09:00' or getoffwork_time<'18:00') AND isnull(RestCount, 0) = 0";
                        break;
                    case 3:
                        T_SQL += " AND (work_time>'09:00' or getoffwork_time<'18:00') AND isnull(RestCount, 0) <> 0";
                        break;
                    case 4:
                        T_SQL += " AND work_time>'09:00'";
                        break;
                    case 5:
                        T_SQL += " AND getoffwork_time<'18:00'";
                        break;
                    case 6:
                        T_SQL += " AND OT='Y'";
                        break;
                    default:
                        break;
                }
                T_SQL += @"  ORDER BY u_BC,userID,attendance_date";
                
                var YYYY = YM.Substring(0, 4);
                var MM = YM.Substring(4, 2);

                var parameters = new List<SqlParameter>()
                {
                    new SqlParameter("@yyyy", YYYY),
                    new SqlParameter("@MM", MM),
                    new SqlParameter("@yyyyMM", YM),
                    new SqlParameter("@user", User_Num)
                };

                var result = _ADO.ExecuteQuery(T_SQL, parameters).AsEnumerable().Select( row => new Attendance_per_res 
                {
                    EarlyMin = row.Field<int>("EarlyMin"),
                    user_name = row.Field<string>("U_name"),
                    userID = row.Field<string>("userID"),
                    attendance_date = row.Field<string>("attendance_date"),
                    work_time = row.Field<string>("work_time"),
                    Late = row.Field<int>("Late"),
                    work_status = row.Field<string>("work_status"),
                    getoffwork_time = row.Field<string>("getoffwork_time"),
                    early = row.Field<int>("early"),
                    offwork_status = row.Field<string>("offwork_status"),
                    U_Na = row.Field<string>("U_Na"),
                    U_BC = row.Field<string>("U_BC"),
                    RestCount = row.Field<int>("RestCount")
                }).ToList();

                #region 請假資訊
                // 先取出所有需要查詢的日期，並過濾掉重複值
                var dateList = result.Where(row => row.RestCount != 0 && !string.IsNullOrEmpty(row.attendance_date)).Select(row => row.attendance_date).Distinct().ToList();
                if (dateList.Any())
                {
                    {
                        // 動態產生 SQL
                        var dateParams = dateList.Select((date, index) => new SqlParameter($"@date{index}", date)).ToList();
                        var inClause = string.Join(",", dateParams.Select(p => p.ParameterName));

                        var SQL_FR = $@"select case when FR_step_now = 1 then '代理人-'+FR_01_U_name when FR_step_now = 2 then '直屬主管-'+FR_02_U_name
                           when FR_step_now = 3 then '單位主管-'+FR_03_U_name when FR_step_now = 9 then '人資-' when FR_step_now = 0 then ''
                           end +  FR_sign_type_name as FR_sign_type_name_desc,* from (
                           select FR_U_num,convert(varchar, FR_date_begin, 111) FR_date,convert(varchar(16),FR_date_begin, 120) FR_date_begin,
                           convert(varchar(16),FR_date_end, 120) FR_date_end,FR_total_hour ,FR_step_now
                           ,(select top 1 U_name from User_M where u_num= FR_step_01_num)FR_01_U_name
                           ,(select top 1 U_name from User_M where u_num= FR_step_02_num)FR_02_U_name
                           ,(select top 1 U_name from User_M where u_num= FR_step_03_num)FR_03_U_name
                           ,(select item_D_name from Item_list where item_M_code = 'FR_kind' AND item_D_type='Y' AND item_D_code = Flow_rest.FR_kind AND del_tag='0') as FR_kind_show
                           ,(select item_D_name from Item_list where item_M_code = 'Flow_sign_type' AND item_D_type='Y' AND item_D_code = Flow_rest.FR_sign_type AND del_tag='0') as FR_sign_type_name
                           ,(select item_D_color from Item_list where item_M_code = 'Flow_sign_type' AND item_D_type='Y' AND item_D_code = Flow_rest.FR_sign_type AND del_tag='0') as FR_sign_type_color
                           from Flow_rest where del_tag = '0' and FR_cancel<>'Y' and FR_U_num=@user_FR
                           and convert(varchar, FR_date_begin, 111) IN ({inClause}) ) A"; // <-- 關鍵修改：改成 IN

                        // 加入使用者參數與動態生成的日期參數
                        var parameters_FR = new List<SqlParameter> { new SqlParameter("@user_FR", User_Num) };
                        parameters_FR.AddRange(dateParams);

                        // 用來作為後續對齊日期的依據
                        var allFlowsLookup = _ADO.ExecuteQuery(SQL_FR, parameters_FR)
                            .AsEnumerable()
                            .Select(row => new
                            {
                                FR_date = row.Field<string>("FR_date"),
                                Flow = new Attendance_Flow
                                {
                                    FR_kind_show = row.Field<string>("FR_kind_show"),
                                    FR_sign_type_name_desc = row.Field<string>("FR_sign_type_name_desc"),
                                    FR_sign_type_color = row.Field<string>("FR_sign_type_color"),
                                    FR_date_begin = row.Field<string>("FR_date_begin"),
                                    FR_date_end = row.Field<string>("FR_date_end"),
                                    FR_total_hour = row.Field<decimal>("FR_total_hour")
                                }
                            })
                            .ToLookup(x => x.FR_date, x => x.Flow);

                        result.Where(row => row.RestCount != 0).ToList().ForEach(row =>
                        {
                            row.attendance_Flows = allFlowsLookup[row.attendance_date].ToList();
                        });
                    }
                }
                #endregion

                return result;
            }
            catch (Exception)
            {

                throw;
            }
        }

        public void AttendanceCardUpd(string YM, string User_Num)
        {
            try
            {
                if (User_Num != "AA999")
                {
                    var SQL_NC = @"delete attendance where substring(@yyyyMM_NC,1,4)+'/'+attendance_date = format( SYSDATETIME(),'yyyy/MM/dd')  and userID = @user_NC";
                    var parameters_NC = new List<SqlParameter>()
                    {
                        new SqlParameter("@yyyyMM_NC", YM),
                        new SqlParameter("@user_NC", User_Num)
                    };
                    _ADO.ExecuteQuery(SQL_NC, parameters_NC);
                    var SQL_IN = @"insert into  dbo.attendance select U_name,userID,CONVERT(varchar(6), convert(datetime ,Date_Record), 112) yyyymm,U_BC user_apart
                                   ,CONVERT(varchar(5), convert(datetime ,Date_Record), 101) attendance_date,work_time, getoffwork_time,SYSDATETIME() inputdate,
                                   'Card_SYS'user_num,'N' ac_flag 
                                   from (
                                          select WT.userID,WT.Date_Record,CAST(WO.Time_Record as varchar(5)) work_time,CAST(WT.Time_Record as varchar(5)) getoffwork_time 
                                          from (
                                                 SELECT userID,Date_Record,MAX(Time_Record) Time_Record FROM Card_Machine group by userID,Date_Record ) WT
                                          Join (
                                                 SELECT userID,Date_Record,min(Time_Record) Time_Record FROM Card_Machine 
                                                 where Time_Record between '04:00' and '17:00' group by userID,Date_Record ) WO on WT.userID=WO.userID AND WT.Date_Record=WO.Date_Record
                                        ) A
                                   Left Join User_M M on A.userID= M.u_num
                                   where  U_BC in ('BC0100','BC0200','BC0400','BC0500','BC0300','BC0600','BC0700','BC0800','BC0801','BC0802','BC0803','BC0804','BC0900')
                                   and [Date_Record]=format( SYSDATETIME(),'yyyy-MM-dd') and userID=@user_IN";
                    var parameters_IN = new List<SqlParameter>()
                    {
                        new SqlParameter("@user_IN", User_Num)
                    };
                    _ADO.ExecuteQuery(SQL_IN, parameters_IN);
                }
            }
            catch (Exception)
            {

                throw;
            }
        }
    }
}
