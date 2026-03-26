using KF_WebAPI.BaseClass;
using KF_WebAPI.BaseClass.AE;
using KF_WebAPI.BaseClass.Max104;
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
using System.Reflection;
using System.Reflection.PortableExecutable;

namespace KF_WebAPI.DataLogic
{
    public class AE_HM
    {
        ADOData _adoData = new ADOData();
        FuncHandler _Fun = new FuncHandler();

        public ResultClass<string> Complaint_LQuery(Complaint_M_req model)
        {
            var parameters = new List<SqlParameter>();
            ResultClass<string> resultClass = new();
            try
            {
                var T_SQL = @"SELECT Comp_Id,CompDate,ub.item_D_name as CompSou_show,ContractID,CS_name,CS_PID,Um.U_name,
                              Remark,Cs.item_D_name as Risk_show,ResponsWay,ResponsStates,CASE WHEN IsClose='N' THEN '尚未結案' ELSE '結案' END as Close_show,CloseDate 
                              FROM Complaint Ct 
                              LEFT JOIN User_M Um ON Ct.Sales_num = Um.u_num
                              LEFT JOIN Item_list Ub ON ub.item_M_code = 'CompSou'  AND ub.item_D_code = Ct.CompSou
                              LEFT JOIN Item_list Cs ON Cs.item_M_code = 'RiskLevel'  AND Cs.item_D_code = Ct.RiskLevel
                              WHERE 1 = 1 AND IsClose = @IsClose";
                parameters.Add(new SqlParameter("@IsClose", model.IsClose));
                if (!string.IsNullOrWhiteSpace(model.CheckDateS) && !string.IsNullOrWhiteSpace(model.CheckDateE))
                {
                    T_SQL += @" AND CompDate BETWEEN @CheckDateS and @CheckDateE";
                    parameters.Add(new SqlParameter("@CheckDateS", FuncHandler.ConvertROCToGregorian(model.CheckDateS)));
                    parameters.Add(new SqlParameter("@CheckDateE", FuncHandler.ConvertROCToGregorian(model.CheckDateE)));
                }
                if (!string.IsNullOrWhiteSpace(model.CS_name))
                {
                    T_SQL += @" AND CS_name = @CS_name";
                    parameters.Add(new SqlParameter("@CS_name", model.CS_name));
                }
                if (!string.IsNullOrWhiteSpace(model.CS_PID))
                {
                    T_SQL += @" AND CS_PID = @CS_PID";
                    parameters.Add(new SqlParameter("@CS_PID", model.CS_PID));
                }
                if (!string.IsNullOrWhiteSpace(model.Sales_num))
                {
                    T_SQL += @" AND Sales_num = @Sales_num";
                    parameters.Add(new SqlParameter("@Sales_num", model.Sales_num));
                }
                if (!string.IsNullOrWhiteSpace(model.RiskLevel))
                {
                    T_SQL += @" AND RiskLevel = @RiskLevel";
                    parameters.Add(new SqlParameter("@RiskLevel", model.RiskLevel));
                }
                if (!string.IsNullOrWhiteSpace(model.CompSou))
                {
                    T_SQL += @" AND CompSou = @CompSou";
                    parameters.Add(new SqlParameter("@CompSou", model.CompSou));
                }
                if (model.RoleType == "RoleB")
                {
                    T_SQL += @" AND C.Sales_num = @UserNum";
                    parameters.Add(new SqlParameter("@UserNum", model.User));
                }
                if (model.RoleType == "RoleC")
                {
                    T_SQL += @" AND M1.U_BC = @UserBC";
                    parameters.Add(new SqlParameter("@UserBC", model.U_BC));
                }

                var result = _adoData.ExecuteQuery(T_SQL, parameters).AsEnumerable().Select(row => new
                {
                    Comp_Id = row.Field<decimal>("Comp_Id"),
                    CompDate = row.Field<string>("CompDate"),
                    CompSou_show = row.Field<string>("CompSou_show"),
                    ContractID = row.Field<string>("ContractID"),
                    CS_name = _Fun.DeCodeBNWords(row.Field<string>("CS_name")),
                    CS_PID = row.Field<string>("CS_PID"),
                    U_name = row.Field<string>("U_name"),
                    Remark = row.Field<string>("Remark"),
                    Risk_show = row.Field<string>("Risk_show"),
                    ResponsWay = row.Field<string>("ResponsWay"),
                    ResponsStates = row.Field<string>("ResponsStates"),
                    Close_show = row.Field<string>("Close_show"),
                    CloseDate = row.Field<string>("CloseDate")
                }).ToList();

                resultClass.objResult = JsonConvert.SerializeObject(result);
                return resultClass;
            }
            catch (Exception)
            {
                throw;
            }
        }
        public int Complaint_Ins(Complaint_M model)
        {
            int result = 0;
            try
            {
                var T_SQL = @"Insert into Complaint (CS_name,CS_PID,Sales_num,PassID,ContractID,CompDate,CompSou,RiskLevel,Complaint,ResponsBC,Remark,
                              ResponsWay,ResponsStates,CloseDate,IsClose,add_date,add_num) 
                              Values (@CS_name,@CS_PID,@Sales_num,@PassID,@ContractID,@CompDate,
                              @CompSou,@RiskLevel,@Complaint,@ResponsBC,@Remark,@ResponsWay,@ResponsStates,@CloseDate,@IsClose,getdate(),@add_num)";
                var parameters = new List<SqlParameter>
                {
                    new SqlParameter("@CS_name",model.CS_name),
                    new SqlParameter("@CS_PID",model.CS_PID),
                    new SqlParameter("@Sales_num",string.IsNullOrWhiteSpace(model.Sales_num)? (object)DBNull.Value:model.Sales_num),
                    new SqlParameter("@PassID",string.IsNullOrWhiteSpace(model.PassID) ?(object) DBNull.Value : model.PassID),
                    new SqlParameter("@ContractID",string.IsNullOrWhiteSpace(model.ContractID) ?(object) DBNull.Value : model.ContractID),
                    new SqlParameter("@CompDate",FuncHandler.ConvertROCToGregorian(model.CompDate)),
                    new SqlParameter("@CompSou",model.CompSou),
                    new SqlParameter("@RiskLevel",model.RiskLevel),
                    new SqlParameter("@Complaint",model.Complaint),
                    new SqlParameter("@ResponsBC",string.IsNullOrWhiteSpace(model.ResponsBC)?(object)DBNull.Value : model.ResponsBC),
                    new SqlParameter("@Remark",model.Remark),
                    new SqlParameter("@ResponsWay",string.IsNullOrWhiteSpace(model.ResponsWay) ?(object) DBNull.Value : model.ResponsWay),
                    new SqlParameter("@ResponsStates",string.IsNullOrWhiteSpace(model.ResponsStates) ?(object) DBNull.Value : model.ResponsStates),
                    new SqlParameter("@CloseDate",string.IsNullOrWhiteSpace(model.CloseDate)? (object)DBNull.Value: FuncHandler.ConvertROCToGregorian(model.CloseDate)),
                    new SqlParameter("@IsClose",model.IsClose),
                    new SqlParameter("@add_num",model.tbInfo.add_num)
                };
                result = _adoData.ExecuteNonQuery(T_SQL, parameters);
                return result;
            }
            catch (Exception)
            {
                throw;
            }
        }

        public ResultClass<string> Complaint_SQuery(string Comp_Id)
        {
            ResultClass<string> resultClass = new();

            try
            {
                var T_SQL = @"SELECT Comp_Id,CompDate,CompSou,ContractID,CS_name,CS_PID,Um.U_name,Remark,RiskLevel,
                              ResponsWay,ResponsStates,IsClose,Um.U_name,Sales_num,CloseDate,Complaint
                              FROM Complaint Ct 
                              LEFT JOIN User_M Um ON Ct.Sales_num = Um.u_num
                              WHERE Comp_Id = @Comp_Id";
                var parameters = new List<SqlParameter>()
                {
                    new SqlParameter("@Comp_Id",Comp_Id)
                };
                var result = _adoData.ExecuteQuery(T_SQL, parameters).AsEnumerable().Select(row => new
                {
                    Comp_Id = row.Field<decimal>("Comp_Id"),
                    CompDate = row.Field<string>("CompDate"),
                    Complaint = row.Field<string>("Complaint"),
                    CompSou = row.Field<string>("CompSou"),
                    ContractID = row.Field<string>("ContractID"),
                    CS_name = _Fun.DeCodeBNWords(row.Field<string>("CS_name")),
                    CS_PID = row.Field<string>("CS_PID"),
                    U_num = row.Field<string>("Sales_num"),
                    U_name = row.Field<string>("U_name"),
                    Remark = row.Field<string>("Remark"),
                    RiskLevel = row.Field<string>("RiskLevel"),
                    ResponsWay = row.Field<string>("ResponsWay"),
                    ResponsStates = row.Field<string>("ResponsStates"),
                    IsClose = row.Field<string>("IsClose"),
                    CloseDate = row.Field<string>("CloseDate")
                }).FirstOrDefault();

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
