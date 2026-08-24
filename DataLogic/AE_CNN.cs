using KF_WebAPI.BaseClass;
using KF_WebAPI.FunctionHandler;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Newtonsoft.Json;
using System.Data;
using static KF_WebAPI.BaseClass.AE.Telemarketing;

namespace KF_WebAPI.DataLogic
{
    public class AE_CNN
    {
        ADOData _adoData = new ADOData();
        FuncHandler _Fun = new FuncHandler();

        public ResultClass<string> Recycle_M_Lquery(Recycle_req model)
        {
            var parameters = new List<SqlParameter>();
            ResultClass<string> resultClass = new();

            try
            {
                RecycleAuto();

                var T_SQL = @"WITH Latest_TD AS ( SELECT D.TM_id,D.TM_D_id,ISNULL(D.Fin_type, 'NA') AS Fin_type,ISNULL(D.Save_day, 0) AS Save_day,
                              ISNULL(D.Call_date, D.add_date) AS Call_date,ROW_NUMBER() OVER (PARTITION BY D.TM_id ORDER BY D.TM_D_id DESC) AS rn FROM Telemarketing_D D),
                              CallLog_Count AS ( SELECT TM_id,COUNT(1) AS CallCount FROM Telemarketing_Log GROUP BY TM_id )
                              SELECT DISTINCT ha.cs_name,ha.CS_ID,ha.cs_mtel1,um.U_name,CONVERT(VARCHAR(10), tm.Tel_Assign_date, 120) AS Tel_Assign_date,
                              ISNULL(tm.Tel_Assign, 'N') AS Tel_Assign,il_1.item_D_name AS Fin_type_name,
                              CONVERT(VARCHAR(10), DATEADD(DAY, td.Save_day + 60, tm.Mag_Assign_date), 120) AS REC_date,ISNULL(L.CallCount, 0) AS CallCount,
                              tm.tm_id,td.Fin_type,CASE WHEN tm.Mag_Assign_date IS NULL THEN 0 ELSE DATEDIFF(DAY, tm.Mag_Assign_date, GETDATE()) END AS AsDay_M 
                              FROM Telemarketing_M tm
                              LEFT JOIN Latest_TD td ON tm.TM_id = td.TM_id AND td.rn = 1
                              LEFT JOIN view_Telemarketing_source ha ON tm.ha_id = ha.HA_id AND tm.TM_type = ha.TM_type
                              LEFT JOIN user_m um ON um.u_num = tm.assign_num
                              LEFT JOIN Item_list il_1 ON td.Fin_type = il_1.item_D_code AND il_1.item_M_code = 'fin_type'
                              LEFT JOIN CallLog_Count L ON tm.TM_id = L.TM_id
                              WHERE tm.Mag_Assign = 'Y' AND Tel_Assign = 'Y' AND tm.TM_type = '1'";

                if(model.tbInfo.edit_num== "K0378")
                {
                    T_SQL += @" AND tm.Assign_num IN ('K0379','K0380')";
                }
                else
                {
                    T_SQL += @" AND tm.Assign_num NOT IN ('K0379','K0380')";
                }
                if(!string.IsNullOrEmpty(model.CS_Name))
                {
                    T_SQL += " AND cs_name = @cs_name";
                    parameters.Add(new SqlParameter("@cs_name", model.CS_Name));
                }
                if(!string.IsNullOrEmpty(model.U_num))
                {
                    T_SQL += " AND Assign_num = @Assign_num";
                    parameters.Add(new SqlParameter("@Assign_num", model.U_num));
                }
                if (!string.IsNullOrEmpty(model.finType))
                {
                    T_SQL += " AND Fin_type=@Fin_type";
                    parameters.Add(new SqlParameter("@Fin_type",model.finType));
                }
                if(!string.IsNullOrEmpty(model.telAsgDateS) && !string.IsNullOrEmpty(model.telAsgDateE))
                {
                    T_SQL += " AND Tel_Assign_date between @telAsgDateS and @telAsgDateE";
                    parameters.Add(new SqlParameter("@telAsgDateS", FuncHandler.ConvertROCToGregorian(model.telAsgDateS)));
                    parameters.Add(new SqlParameter("@telAsgDateE", FuncHandler.ConvertROCToGregorian(model.telAsgDateE)));
                }
                var result = _adoData.ExecuteQuery(T_SQL,parameters).AsEnumerable().Select(row => new
                {
                    cs_name = _Fun.DeCodeBNWords(row.Field<string>("cs_name")),
                    CS_ID = row.Field<string>("CS_ID"),
                    cs_mtel1 = row.Field<string>("cs_mtel1"),
                    U_name = row.Field<string>("U_name"),
                    Tel_Assign_date = row.Field<string>("Tel_Assign_date"),
                    Tel_Assign = row.Field<string>("Tel_Assign"),
                    Fin_type_name = row.Field<string>("Fin_type_name"),
                    REC_date = row.Field<string>("REC_date"),
                    CallCount = row.Field<int>("CallCount"),
                    tm_id = row.Field<decimal>("tm_id")
                }).ToList();
                resultClass.objResult = JsonConvert.SerializeObject(result);
                return resultClass;
            }
            catch (Exception)
            {

                throw;
            }
        }

        public ResultClass<string> Call_Detail_LQuery(decimal tmID)
        {
            ResultClass<string> resultClass = new();

            try
            {
                var T_SQL = @"select ha.cs_name,ha.cs_mtel1,Case When Call_type='CALL_T06' Then '未聯絡' Else '已聯絡' END AS Call_type_Show,
                              il_1.item_D_name AS Fin_type_name,tl.Memo,Call_date_S,Call_date_E,DATEDIFF(MINUTE, Call_date_S, Call_date_E) AS minute_diff
                              from Telemarketing_Log tl
                              LEFT JOIN Telemarketing_M tm ON tm.TM_id = tl.TM_id
                              LEFT JOIN view_Telemarketing_source ha ON tm.ha_id = ha.HA_id AND tm.TM_type = ha.TM_type
                              LEFT JOIN Item_list il_1 ON tl.Fin_type = il_1.item_D_code AND il_1.item_M_code = 'fin_type'
                              where tl.TM_id = @TM_id order by Call_date_S";
                var parameters = new List<SqlParameter>()
                {
                    new SqlParameter("@TM_id",tmID)
                };
                var resultTable = _adoData.ExecuteQuery(T_SQL, parameters);
                resultClass.objResult = JsonConvert.SerializeObject(resultTable);
                return resultClass;
            }
            catch (Exception)
            {

                throw;
            }
        }

        public ResultClass<string> Recycle_M_Upd(string user,List<string> ids)
        {
            ResultClass<string> resultClass = new();

            try
            {
                var sqlParams = new List<SqlParameter>();
                var parameterNames = new List<string>();
                for (int i = 0; i < ids.Count; i++)
                {
                    string paramName = $"@id{i}";
                    parameterNames.Add(paramName);
                    sqlParams.Add(new SqlParameter(paramName, ids[i]));
                }

                var inClause = string.Join(", ", parameterNames);

                var T_SQL_M = $@"UPDATE Telemarketing_M SET add_date = getdate(), assign_num = '', Tel_Assign = NULL, Tel_Assign_date = NULL WHERE tm_id IN ({inClause})";

                int intResultM = _adoData.ExecuteNonQuery(T_SQL_M, sqlParams);

                if (intResultM != 0)
                {
                    var T_SQL_D = $@"update Telemarketing_D set isRec = 'Y' where tm_id IN ({inClause}) and isRec is null";
                }

                return resultClass;
            }
            catch (Exception)
            {

                throw;
            }
        }

        public void RecycleAuto()
        {
            try
            {
                var T_SQL = @"WITH Latest_TD AS ( SELECT D.TM_id,D.TM_D_id,ISNULL(D.Fin_type, 'NA') AS Fin_type,ISNULL(D.Save_day, 0) AS Save_day,
                          ISNULL(D.Call_date, D.add_date) AS Call_date,ROW_NUMBER() OVER (PARTITION BY D.TM_id ORDER BY D.TM_D_id DESC) AS rn FROM Telemarketing_D D),
                          CallLog_Count AS ( SELECT TM_id,COUNT(1) AS CallCount FROM Telemarketing_Log GROUP BY TM_id )
                          SELECT DISTINCT ha.cs_name,ha.CS_ID,ha.cs_mtel1,um.U_name,CONVERT(VARCHAR(10), tm.Tel_Assign_date, 120) AS Tel_Assign_date,
                          ISNULL(tm.Tel_Assign, 'N') AS Tel_Assign,il_1.item_D_name AS Fin_type_name,
                          CONVERT(VARCHAR(10), DATEADD(DAY, td.Save_day + 60, tm.Mag_Assign_date), 120) AS REC_date,ISNULL(L.CallCount, 0) AS CallCount,
                          tm.tm_id,td.Fin_type,CASE WHEN tm.Mag_Assign_date IS NULL THEN 0 ELSE DATEDIFF(DAY, tm.Mag_Assign_date, GETDATE()) END AS AsDay_M 
                          FROM Telemarketing_M tm
                          LEFT JOIN Latest_TD td ON tm.TM_id = td.TM_id AND td.rn = 1
                          LEFT JOIN view_Telemarketing_source ha ON tm.ha_id = ha.HA_id AND tm.TM_type = ha.TM_type
                          LEFT JOIN user_m um ON um.u_num = tm.assign_num
                          LEFT JOIN Item_list il_1 ON td.Fin_type = il_1.item_D_code AND il_1.item_M_code = 'fin_type'
                          LEFT JOIN CallLog_Count L ON tm.TM_id = L.TM_id
                          WHERE tm.Mag_Assign = 'Y' AND Tel_Assign = 'Y' AND tm.TM_type = '1' AND Fin_type IN ('FIN_T28')";

                var result = _adoData.ExecuteSQuery(T_SQL).AsEnumerable().Select(row => new
                {
                    tm_id = row.Field<decimal>("tm_id"),
                    AsDay_M = row.Field<int>("AsDay_M")
                }).ToList();

                List<string> ids = result.Where(x => x.AsDay_M > 60).Select(x => x.tm_id.ToString()).ToList();

                if (ids.Any())
                {
                    Recycle_M_Upd("K0003", ids);
                }
            }
            catch (Exception)
            {

                throw;
            }
        }
    }
}
