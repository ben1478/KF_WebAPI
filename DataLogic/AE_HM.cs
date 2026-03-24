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
using System.Reflection;
using System.Reflection.PortableExecutable;

namespace KF_WebAPI.DataLogic
{
    public class AE_HM
    {
        ADOData _adoData = new ADOData();
        FuncHandler _Fun = new FuncHandler();

        //public string Complaint_LQuery(Complaint_M_req model)
        //{
        //    var parameters = new List<SqlParameter>();
        //    try
        //    {
        //        var T_SQL = @"SELECT * FROM Complaint C 
        //                      LEFT JOIN User_M M1 ON C.Sales_num = M1.u_num
        //                      LEFT JOIN Item_list ub ON ub.item_M_code = 'branch_company'  AND ub.item_D_code = M1.U_BC
        //                      LEFT JOIN User_M M2 on C.Add_num = M2.u_num
        //                      WHERE 1 = 1 AND IsClose = @IsClose";
        //        parameters.Add(new SqlParameter("@IsClose", model.IsClose));
        //        if (!string.IsNullOrWhiteSpace(model.CheckDateS) && !string.IsNullOrWhiteSpace(model.CheckDateE))
        //        {
        //            T_SQL += @" AND CompDate BETWEEN @CheckDateS and @CheckDateE";
        //            parameters.Add(new SqlParameter("@CheckDateS",FuncHandler.ConvertROCToGregorian(model.CheckDateS)));
        //            parameters.Add(new SqlParameter("@CheckDateE", FuncHandler.ConvertROCToGregorian(model.CheckDateE)));
        //        }
        //        if (!string.IsNullOrWhiteSpace(model.CS_name))
        //        {
        //            T_SQL += @" AND CS_name = @CS_name";
        //            parameters.Add(new SqlParameter("@CS_name",model.CS_name));
        //        }
        //        if (!string.IsNullOrWhiteSpace(model.CS_PID))
        //        {
        //            T_SQL += @" AND CS_PID = @CS_PID";
        //            parameters.Add(new SqlParameter("@CS_PID", model.CS_PID));
        //        }
        //        if (!string.IsNullOrWhiteSpace(model.Sales_num))
        //        {
        //            T_SQL += @" AND Sales_num = @Sales_num";
        //            parameters.Add(new SqlParameter("@Sales_num", model.Sales_num));
        //        }
        //        if (!string.IsNullOrWhiteSpace(model.RiskLevel))
        //        {
        //            T_SQL += @" AND RiskLevel = @RiskLevel";
        //            parameters.Add(new SqlParameter("@RiskLevel", model.RiskLevel));
        //        }
        //        if (!string.IsNullOrWhiteSpace(model.CompSou))
        //        {
        //            T_SQL += @" AND CompSou = @CompSou";
        //            parameters.Add(new SqlParameter("@CompSou", model.CompSou));
        //        }
        //        if(model.RoleType == "RoleA")
        //        {
        //            T_SQL += @" AND M1.U_BC = @BC_code";
        //        }
        //    }
        //    catch (Exception)
        //    {

        //        throw;
        //    }
        //}
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
    }
}
