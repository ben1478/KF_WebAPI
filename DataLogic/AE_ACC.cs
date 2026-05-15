using KF_WebAPI.BaseClass;
using KF_WebAPI.BaseClass.AE;
using KF_WebAPI.BaseClass.Winton;
using KF_WebAPI.Controllers;
using KF_WebAPI.FunctionHandler;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections;
using System.Data;
using System.Drawing;
using System.Dynamic;
using System.Net.NetworkInformation;
using System.Reflection;
using System.Reflection.PortableExecutable;
using System.Xml.Linq;

namespace KF_WebAPI.DataLogic
{
    public class AE_ACC
    {
        ADOData _adoData = new ADOData();
        FuncHandler _Fun = new FuncHandler();

        public ResultClass<string> GetHouseACH_LQuery(int yyyyMM, string type, string pjtype)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            try
            {
                var parameters = new List<SqlParameter>();
                var T_SQL = @"select rm.RCM_id,b.HS_id,b.Send_amount_date,ha.CS_name,um.U_name,li_1.item_D_name as U_BC,b.get_amount
                              ,b.get_amount_date,b.interest_rate_pass,li_2.item_D_name as pjName,rm.month_total,li_3.item_D_name as achState,Ach_Note
                              from view_HS_Base b
                              INNER JOIN Receivable_M rm ON rm.HS_id = b.HS_id AND rm.del_tag = 0
                              INNER JOIN House_apply ha ON ha.HA_id = b.HA_id AND ha.del_tag = 0
                              LEFT JOIN User_M um ON um.U_num = ha.plan_num
                              LEFT JOIN Item_list li_1 ON li_1.item_M_code = 'branch_company' and li_1.item_D_code = um.U_BC
                              LEFT JOIN Item_list li_2 ON li_2.item_M_code = 'project_title' and li_2.item_D_code = b.project_title
                              LEFT JOIN Item_list li_3 ON li_3.item_M_code = 'Ach_State' and li_3.item_D_code = rm.Ach_State
                              WHERE b.Send_result_type = 'SRT002' AND b.get_amount_type = 'GTAT002'";

                switch (pjtype)
                {
                    case "Moto":
                        T_SQL += @" AND b.project_title IN ('PJ00046','PJ00047')";
                        break;
                    case "Car":
                        T_SQL += @" AND b.project_title IN ('PJ00048')";
                        break;
                    default:
                        T_SQL += @" AND b.project_title NOT IN ('PJ00046','PJ00047','PJ00048')";
                        break;
                }

                switch (type)
                {
                    case "M":
                        T_SQL += @" AND (YEAR(b.get_amount_date)*100+MONTH(b.get_amount_date)) = @TargetMonth 
                               ORDER BY b.get_amount_date ,b.HS_id";
                        parameters.Add(new SqlParameter("@TargetMonth", yyyyMM));
                        break;
                    case "Y":
                        T_SQL += @" AND (YEAR(b.get_amount_date)*100) = @TargetYear 
                               ORDER BY b.get_amount_date ,b.HS_id";
                        int yyyy = (yyyyMM / 100) * 100;
                        parameters.Add(new SqlParameter("@TargetYear", yyyy));
                        break;
                    default:
                        T_SQL += @" ORDER BY b.get_amount_date ,b.HS_id";
                        break;
                }

                var result = _adoData.ExecuteQuery(T_SQL, parameters).AsEnumerable().Select(row=> new
                {
                    RCM_id = row.Field<decimal>("RCM_id"),
                    HS_id = row.Field<decimal>("HS_id"),
                    str_Send_amount_date = FuncHandler.ConvertGregorianToROC(row.Field<DateTime>("Send_amount_date").ToString("yyyy/MM/dd")),
                    CS_name = _Fun.DeCodeBNWords(row.Field<string>("CS_name")),
                    U_name = _Fun.DeCodeBNWords(row.Field<string>("U_name")),
                    U_BC = row.Field<string>("U_BC"),
                    get_amount = row.Field<string>("get_amount"),
                    str_get_amount_date = FuncHandler.ConvertGregorianToROC(row.Field<DateTime>("get_amount_date").ToString("yyyy/MM/dd")),
                    interest_rate_pass = row.Field<string>("interest_rate_pass"),
                    pjName = _Fun.DeCodeBNWords(row.Field<string>("pjName")),
                    month_total = row.Field<int>("month_total"),
                    achState = row.Field<string>("achState"),
                    Ach_Note = row.Field<string>("Ach_Note")
                }).ToList();

                resultClass.ResultCode = "000";
                resultClass.objResult = JsonConvert.SerializeObject(result);

                return resultClass;
            }
            catch (Exception)
            {

                throw;
            }
            
        }
    }
}
