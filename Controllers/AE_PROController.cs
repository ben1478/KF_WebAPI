using KF_WebAPI.BaseClass;
using KF_WebAPI.BaseClass.AE;
using KF_WebAPI.FunctionHandler;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Newtonsoft.Json;
using System.Data;
using System.Reflection;

namespace KF_WebAPI.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class AE_PROController : ControllerBase
    {
        /// <summary>
        /// 提供廠商資料列表 GetManufaList
        /// </summary>
        [HttpGet("GetManufaList")]
        public ActionResult<ResultClass<string>> GetManufaList(string Name)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"select Company_name as name from Manufacturer where Company_name like @Name";
                parameters.Add(new SqlParameter("@Name", "%" + Name + "%"));
                #endregion
                var dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
                resultClass.ResultCode = "000";
                resultClass.objResult = JsonConvert.SerializeObject(dtResult);
                return Ok(resultClass);
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass);
            }
        }

        /// <summary>
        /// 採購單資料新增 Procurement_M_Ins
        /// </summary>
        [HttpPost("Procurement_M_Ins")]
        public ActionResult<ResultClass<string>> Procurement_M_Ins(Procurement_Ins model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();
            try
            {
                ADOData _adoData = new ADOData();
                #region SQL_Procurement_M
                var parameters_m = new List<SqlParameter>();
                var T_SQL_M = @"Insert into Procurement_M (PM_ID,PM_type,PM_BC,PM_Pay_Type,PM_AppDate,PM_U_num,PM_Caption,PM_Amt,PM_Busin_Tax,PM_Tax_Amt
                    ,PM_Other,add_date,add_num,add_ip,edit_date,edit_num,edit_ip,PM_cknum)
                    Values (@PM_ID,@PM_type,@PM_BC,@PM_Pay_Type,@PM_AppDate,@PM_U_num,@PM_Caption,@PM_Amt,@PM_Busin_Tax,@PM_Tax_Amt
                    ,@PM_Other,GETDATE(),@add_num,@add_ip,GETDATE(),@edit_num,@edit_ip,@PM_cknum)";
                //取PM_ID
                var parameters_id = new List<SqlParameter>();
                var T_SQL_ID = @"exec GetFormID @formtype,@tablename,@New_ID output";
                parameters_id.Add(new SqlParameter("@formtype", "PO"));
                parameters_id.Add(new SqlParameter("@tablename", "Procurement_M"));
                SqlParameter newIdParameter = new SqlParameter("@New_ID", SqlDbType.VarChar, 20)
                {
                    Direction = ParameterDirection.Output
                };
                parameters_id.Add(newIdParameter);
                var result_id = _adoData.ExecuteQuery(T_SQL_ID, parameters_id);
                model.PM_ID = newIdParameter.Value.ToString();

                parameters_m.Add(new SqlParameter("@PM_ID", model.PM_ID));
                parameters_m.Add(new SqlParameter("@PM_type", "PO"));
                parameters_m.Add(new SqlParameter("@PM_BC", model.PM_BC));
                parameters_m.Add(new SqlParameter("@PM_Pay_Type", model.PM_Pay_Type));
                parameters_m.Add(new SqlParameter("@PM_AppDate", DateTime.Now.ToString("yyyy/MM/dd")));
                parameters_m.Add(new SqlParameter("@PM_U_num", model.PM_U_num));
                parameters_m.Add(new SqlParameter("@PM_Caption", model.PM_Caption));
                parameters_m.Add(new SqlParameter("@PM_Amt", model.PM_Amt));
                parameters_m.Add(new SqlParameter("@PM_Busin_Tax", model.PM_Busin_Tax));
                parameters_m.Add(new SqlParameter("@PM_Tax_Amt", model.PM_Tax_Amt));
                if (!string.IsNullOrEmpty(model.PM_Other))
                {
                    parameters_m.Add(new SqlParameter("@PM_Other", model.PM_Other));
                }
                else
                {
                    parameters_m.Add(new SqlParameter("@PM_Other", DBNull.Value));
                }
                parameters_m.Add(new SqlParameter("@add_num", model.PM_U_num));
                parameters_m.Add(new SqlParameter("@add_ip", clientIp));
                parameters_m.Add(new SqlParameter("@edit_num", model.PM_U_num));
                parameters_m.Add(new SqlParameter("@edit_ip", clientIp));
                parameters_m.Add(new SqlParameter("@PM_cknum", FuncHandler.GetCheckNum()));
                #endregion
                int result_m = _adoData.ExecuteNonQuery(T_SQL_M, parameters_m);
                if (result_m == 0)
                {
                    resultClass.ResultCode = "400";
                    resultClass.ResultMsg = "主檔新增失敗";
                    return BadRequest(resultClass);
                }
                else
                {
                    foreach (var item in model.PD_Ins_List) 
                    {
                        #region Procurement_D
                        var parameters_d = new List<SqlParameter>();
                        var T_SQL_D = @"Insert into Procurement_D (PM_ID,PD_Pro_name,PD_Unit,PD_Count,PD_Date,PD_Univalent,PD_Amt,PD_Company_name,PD_Est_Cost
                            ,add_date,add_num,add_ip,edit_date,edit_num,edit_ip) 
                            values (@PM_ID,@PD_Pro_name,@PD_Unit,@PD_Count,@PD_Date,@PD_Univalent,@PD_Amt,@PD_Company_name,@PD_Est_Cost,GETDATE()
                            ,@add_num,@add_ip,GETDATE(),@edit_num,@edit_ip)";
                        parameters_d.Add(new SqlParameter("@PM_ID", model.PM_ID));
                        parameters_d.Add(new SqlParameter("@PD_Pro_name", item.PD_Pro_name));
                        parameters_d.Add(new SqlParameter("@PD_Unit", item.PD_Unit));
                        parameters_d.Add(new SqlParameter("@PD_Count", item.PD_Count));
                        parameters_d.Add(new SqlParameter("@PD_Date", FuncHandler.ConvertROCToGregorian(item.PD_Date)));
                        parameters_d.Add(new SqlParameter("@PD_Univalent", item.PD_Univalent));
                        parameters_d.Add(new SqlParameter("@PD_Amt", item.PD_Amt));
                        parameters_d.Add(new SqlParameter("@PD_Company_name", item.PD_Company_name));
                        parameters_d.Add(new SqlParameter("@PD_Est_Cost", item.PD_Est_Cost));
                        parameters_d.Add(new SqlParameter("@add_num", model.PM_U_num));
                        parameters_d.Add(new SqlParameter("@add_ip", clientIp));
                        parameters_d.Add(new SqlParameter("@edit_num", model.PM_U_num));
                        parameters_d.Add(new SqlParameter("@edit_ip", clientIp));
                        #endregion
                        int result_d = _adoData.ExecuteNonQuery(T_SQL_D, parameters_d);
                        if (result_d == 0)
                        {
                            resultClass.ResultCode = "400";
                            resultClass.ResultMsg = "明細檔新增失敗";
                            return BadRequest(resultClass);
                        }
                    }

                    //寫入審核流程
                    bool resultFlow = FuncHandler.AuditFlow(model.PM_U_num, model.PM_BC, "PO", model.PM_ID, clientIp);
                    if (resultFlow)
                    {
                        resultClass.ResultCode = "000";
                        resultClass.ResultMsg = "新增成功";
                        return Ok(resultClass);
                    }
                    else
                    {
                        resultClass.ResultCode = "400";
                        resultClass.ResultMsg = "審核新增失敗,請洽資訊人員";
                        return BadRequest(resultClass);
                    }
                }
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass); 
            }
           
        }

        /// <summary>
        /// 採購單列表查詢 Procurement_M_LQuery
        /// </summary>
        [HttpGet("Procurement_M_LQuery")]
        public ActionResult<ResultClass<string>> Procurement_M_LQuery(string User)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"select PM.PM_ID,AM.FM_Step,LI.item_D_name,LI.item_D_name AS FM_Step_SignType,PM.PM_cknum,PM.PM_Cancel
                    ,(SELECT COUNT(*) FROM ASP_UpLoad WHERE cknum = PM.PM_cknum and del_tag='0') AS PM_cknum_count 
                    from Procurement_M PM
                    INNER JOIN AuditFlow_M AM ON AM.FM_Source_ID = PM.PM_ID and AM.AF_ID = PM.PM_type
                    LEFT JOIN Item_list LI ON LI.item_D_code = AM.FM_Step_SignType AND LI.item_M_code = 'Flow_sign_type'
                    where PM.add_num = @User ";
                parameters.Add(new SqlParameter("@User", User));
                #endregion
                var dtResult=_adoData.ExecuteQuery(T_SQL, parameters);

                resultClass.ResultCode = "000";
                resultClass.objResult = JsonConvert.SerializeObject(dtResult);
                return Ok(resultClass);
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass); // 返回 500 錯誤碼
            }
        }

        /// <summary>
        /// 採購單單筆查詢 Procurement_M_SQuery
        /// </summary>
        [HttpGet("Procurement_M_SQuery")]
        public ActionResult<ResultClass<string>> Procurement_M_SQuery(string PM_ID)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"select PM_BC,PM_Pay_Type,PM_Caption,PM_Amt,PM_Busin_Tax,PM_Tax_Amt,PM_Other　from Procurement_M where PM_ID=@PM_ID ";
                parameters.Add(new SqlParameter("@PM_ID", PM_ID));
                #endregion
                var dtResult = _adoData.ExecuteQuery(T_SQL, parameters);

                if (dtResult.Rows.Count > 0)
                {
                    var model = dtResult.AsEnumerable().Select(row => new Procurement_Res {
                        PM_BC = row.Field<string>("PM_BC"),
                        PM_Pay_Type = row.Field<string>("PM_Pay_Type"),
                        PM_Caption = row.Field<string>("PM_Caption"),
                        PM_Amt = row.Field<decimal>("PM_Amt"),
                        PM_Busin_Tax = row.Field<decimal>("PM_Busin_Tax"),
                        PM_Tax_Amt = row.Field<decimal>("PM_Tax_Amt"),
                        PM_Other = row.Field<string>("PM_Other")
                    }).FirstOrDefault();

                    #region SQL Procurement_D
                    var parameters_d = new List<SqlParameter>();
                    var T_SQL_D = @"select PD_ID,PD_Pro_name,PD_Unit,PD_Count,PD_Date,PD_Univalent,PD_Amt,PD_Company_name,PD_Est_Cost
                        from Procurement_D where PM_ID=@PM_ID ";
                    parameters_d.Add(new SqlParameter("@PM_ID", PM_ID));
                    #endregion
                    var dtResult_d = _adoData.ExecuteQuery(T_SQL_D, parameters_d);
                    var modelist_d = dtResult_d.AsEnumerable().Select(row => new Procurement_D_Res {
                        PD_ID = row.Field<int>("PD_ID"),
                        PD_Pro_name = row.Field<string>("PD_Pro_name"),
                        PD_Unit = row.Field<string>("PD_Unit"),
                        PD_Count = row.Field<string>("PD_Count"),
                        PD_Date = FuncHandler.ConvertGregorianToROC(row.Field<string>("PD_Date")),
                        PD_Univalent = row.Field<decimal>("PD_Univalent"),
                        PD_Amt = row.Field<decimal>("PD_Amt"),
                        PD_Company_name = row.Field<string>("PD_Company_name"),
                        PD_Est_Cost = row.Field<decimal>("PD_Est_Cost")
                    });

                    model.Procurement_D = modelist_d.ToList();

                    resultClass.ResultCode = "000";
                    resultClass.objResult = JsonConvert.SerializeObject(model);
                    return Ok(resultClass);
                }
                else
                {
                    resultClass.ResultCode = "400";
                    resultClass.objResult = JsonConvert.SerializeObject(resultClass);
                    return BadRequest(resultClass);
                }
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass); // 返回 500 錯誤碼
            }
        }


        /// <summary>
        /// 採購單抽單 Procurement_Canel
        /// </summary>
        [HttpGet("Procurement_Canel")]
        public ActionResult<ResultClass<string>> Procurement_Canel(string PM_ID, string User)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"update Procurement_M set PM_Cancel='Y',cancel_date=GETDATE(),cancel_num=@cancel_num,candel_ip=@candel_ip where PM_ID=@PM_ID";
                parameters.Add(new SqlParameter("@cancel_num", User));
                parameters.Add(new SqlParameter("@candel_ip", clientIp));
                parameters.Add(new SqlParameter("@PM_ID", PM_ID));
                #endregion
                int result = _adoData.ExecuteNonQuery(T_SQL, parameters);
                if (result == 0)
                {
                    resultClass.ResultCode = "400";
                    resultClass.ResultMsg = "取消失敗";
                    return BadRequest(resultClass);
                }
                else
                {
                    resultClass.ResultCode = "000";
                    resultClass.ResultMsg = "取消成功";
                    return Ok(resultClass);
                }
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass); // 返回 500 錯誤碼
            }
        }

        /// <summary>
        /// 採購單單筆修改 Procurement_Upd
        /// </summary>
        [HttpPost("Procurement_M_Upd")]
        public ActionResult<ResultClass<string>> Procurement_M_Upd(Procurement_Ins model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();
            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"update Procurement_M set PM_Pay_Type=@PM_Pay_Type,PM_Caption=@PM_Caption,PM_Amt=@PM_Amt,PM_Busin_Tax=@PM_Busin_Tax,
                    PM_Tax_Amt=@PM_Tax_Amt,PM_Other=@PM_Other,edit_num=@User,edit_date=GETDATE(),edit_ip=@IP where PM_ID=@PM_ID ";
                parameters.Add(new SqlParameter("@PM_ID", model.PM_ID));
                parameters.Add(new SqlParameter("@PM_Pay_Type", model.PM_Pay_Type));
                parameters.Add(new SqlParameter("@PM_Caption", model.PM_Caption));
                parameters.Add(new SqlParameter("@PM_Amt", model.PM_Amt));
                parameters.Add(new SqlParameter("@PM_Busin_Tax", model.PM_Busin_Tax));
                parameters.Add(new SqlParameter("@PM_Tax_Amt", model.PM_Tax_Amt));
                if (!string.IsNullOrEmpty(model.PM_Other))
                {
                    parameters.Add(new SqlParameter("@PM_Other", model.PM_Other));
                }
                else
                {
                    parameters.Add(new SqlParameter("@PM_Other", DBNull.Value));
                }
                parameters.Add(new SqlParameter("@User", model.PM_U_num));
                parameters.Add(new SqlParameter("@IP", clientIp));
                #endregion
                int result = _adoData.ExecuteNonQuery(T_SQL, parameters);
                if (result == 0)
                {
                    resultClass.ResultCode = "400";
                    resultClass.ResultMsg = "變更失敗";
                    return BadRequest(resultClass);
                }
                else
                {
                    var parameters_de = new List<SqlParameter>();
                    var T_SQL_DE = @"Delete Procurement_D where PM_ID=@PM_ID";
                    parameters_de.Add(new SqlParameter("@PM_ID", model.PM_ID));
                    _adoData.ExecuteQuery(T_SQL_DE, parameters_de);
                    foreach (var item in model.PD_Ins_List)
                    {
                        #region Procurement_D
                        var parameters_d = new List<SqlParameter>();
                        var T_SQL_D = @"Insert into Procurement_D (PM_ID,PD_Pro_name,PD_Unit,PD_Count,PD_Date,PD_Univalent,PD_Amt,PD_Company_name,PD_Est_Cost
                            ,add_date,add_num,add_ip,edit_date,edit_num,edit_ip) 
                            values (@PM_ID,@PD_Pro_name,@PD_Unit,@PD_Count,@PD_Date,@PD_Univalent,@PD_Amt,@PD_Company_name,@PD_Est_Cost,GETDATE()
                            ,@add_num,@add_ip,GETDATE(),@edit_num,@edit_ip)";
                        parameters_d.Add(new SqlParameter("@PM_ID", model.PM_ID));
                        parameters_d.Add(new SqlParameter("@PD_Pro_name", item.PD_Pro_name));
                        parameters_d.Add(new SqlParameter("@PD_Unit", item.PD_Unit));
                        parameters_d.Add(new SqlParameter("@PD_Count", item.PD_Count));
                        parameters_d.Add(new SqlParameter("@PD_Date", FuncHandler.ConvertROCToGregorian(item.PD_Date)));
                        parameters_d.Add(new SqlParameter("@PD_Univalent", item.PD_Univalent));
                        parameters_d.Add(new SqlParameter("@PD_Amt", item.PD_Amt));
                        parameters_d.Add(new SqlParameter("@PD_Company_name", item.PD_Company_name));
                        parameters_d.Add(new SqlParameter("@PD_Est_Cost", item.PD_Est_Cost));
                        parameters_d.Add(new SqlParameter("@add_num", model.PM_U_num));
                        parameters_d.Add(new SqlParameter("@add_ip", clientIp));
                        parameters_d.Add(new SqlParameter("@edit_num", model.PM_U_num));
                        parameters_d.Add(new SqlParameter("@edit_ip", clientIp));
                        #endregion
                        int result_d = _adoData.ExecuteNonQuery(T_SQL_D, parameters_d);
                        if (result_d == 0)
                        {
                            resultClass.ResultCode = "400";
                            resultClass.ResultMsg = "明細檔新增失敗";
                            return BadRequest(resultClass);
                        }
                    }
                    resultClass.ResultCode = "000";
                    resultClass.ResultMsg = "修改成功";
                    return Ok(resultClass);
                }
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass); // 返回 500 錯誤碼
            }
        }

        /// <summary>
        /// 取得請採購單列印資料 GetPrintData
        /// </summary>
        [HttpGet("GetPrintData")]
        public ActionResult<ResultClass<string>> GetPrintData(string PM_ID)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters_m = new List<SqlParameter>();
                var T_SQL_M = @"select LI.item_d_name as BC_Name,UM.U_name,PM.PM_AppDate,PM.PM_ID,PM.PM_Caption,Format(PM.PM_Amt,'N0') as PM_Amt,
                    Format(PM.PM_Busin_Tax,'N0') as PM_Busin_Tax,Format(PM.PM_Tax_Amt,'N0') as PM_Tax_Amt,PM.PM_cknum,
                    Case PM.PM_Pay_Type When 'GC' THEN '領現' When 'MT' THEN '匯款' END as PM_Pay_Type,PM.PM_Other,
                    FORMAT(PM.add_date,'yyyy/MM/dd HH:mm') as add_date
                    from Procurement_M PM
                    left join User_M UM on UM.U_num = PM.PM_U_num
                    left join Item_list LI on LI.item_D_code = PM.PM_BC
                    where PM.PM_ID=@PM_ID";
                parameters_m.Add(new SqlParameter("@PM_ID", PM_ID));
                #endregion
                var resultModel = _adoData.ExecuteQuery(T_SQL_M, parameters_m).AsEnumerable().Select(row => new Proc_Print
                {
                    BC_Name = row.Field<string>("BC_Name"),
                    U_name = row.Field<string>("U_name"),
                    PM_AppDate = row.Field<string>("PM_AppDate"),
                    PM_ID = row.Field<string>("PM_ID"),
                    PM_Caption = row.Field<string>("PM_Caption"),
                    PM_Amt = row.Field<string>("PM_Amt"),
                    PM_Busin_Tax = row.Field<string>("PM_Busin_Tax"),
                    PM_Tax_Amt = row.Field<string>("PM_Tax_Amt"),
                    PM_cknum = row.Field<string>("PM_cknum"),
                    PM_Pay_Type = row.Field<string>("PM_Pay_Type"),
                    PM_Other = row.Field<string>("PM_Other"),
                    add_date = row.Field<string>("add_date")
                }).FirstOrDefault();

                #region Proc_Print_Deatil
                var parameters_dt = new List<SqlParameter>();
                var T_SQL_DT = @"select PD_Pro_name,PD_Unit,PD_Count,PD_Date,Format(PD_Univalent,'N0') as PD_Univalent,
                    Format(PD_Amt,'N0') as PD_Amt,PD_Company_name,Format(PD_Est_Cost,'N0') as PD_Est_Cost 
                    from Procurement_D where PM_ID=@PM_ID";
                parameters_dt.Add(new SqlParameter("@PM_ID", PM_ID));
                var result_dt = _adoData.ExecuteQuery(T_SQL_DT, parameters_dt).AsEnumerable().Select(row => new Proc_Print_Deatil
                {
                    PD_Pro_name = row.Field<string>("PD_Pro_name"),
                    PD_Unit = row.Field<string>("PD_Unit"),
                    PD_Count = row.Field<string>("PD_Count"),
                    PD_Date = row.Field<string>("PD_Date"),
                    PD_Univalent = row.Field<string>("PD_Univalent"),
                    PD_Amt = row.Field<string>("PD_Amt"),
                    PD_Company_name = row.Field<string>("PD_Company_name"),
                    PD_Est_Cost = row.Field<string>("PD_Est_Cost")
                }).ToList();
                #endregion

                #region Proc_Print_File
                var parameters_fi = new List<SqlParameter>();
                var T_SQL_FI = @"select upload_name_show from ASP_UpLoad where cknum=@cknum";
                parameters_fi.Add(new SqlParameter("@cknum", resultModel.PM_cknum));
                var result_fi = _adoData.ExecuteQuery(T_SQL_FI, parameters_fi).AsEnumerable().Select(row => new Proc_Print_File
                {
                    upload_name_show = row.Field<string>("upload_name_show"),
                }).ToList();
                #endregion

                #region Proc_Print_Flow
                var parameters_fw = new List<SqlParameter>();
                var T_SQL_FW = @"select AD.FD_Sign_Countersign,AD.FD_Step,UM.U_name,
                    FORMAT(AD.FD_Step_date,'yyyy/MM/dd HH:mm') as FD_Step_date,AD.FD_Step_desc 
                    from AuditFlow_D AD
                    inner join User_M UM on UM.U_num = AD.FD_Step_num
                    where FM_Source_ID=@PM_ID order by AD.FD_Step";
                parameters_fw.Add(new SqlParameter("@PM_ID", PM_ID));
                var result_fw = _adoData.ExecuteQuery(T_SQL_FW, parameters_fw).AsEnumerable().Select(row => new Proc_Print_Flow
                {
                    FD_Sign_Countersign = row.Field<string>("FD_Sign_Countersign"),
                    FD_Step = row.Field<string>("FD_Step"),
                    U_name = row.Field<string>("U_name"),
                    FD_Step_date = row.Field<string>("FD_Step_date"),
                    FD_Step_desc = row.Field<string>("FD_Step_desc")
                }).ToList();
                #endregion

                resultModel.PT_Deatil_List = result_dt;
                resultModel.PT_File_List = result_fi;
                resultModel.PT_Flow_List = result_fw;

                resultClass.ResultCode = "000";
                resultClass.objResult = JsonConvert.SerializeObject(resultModel);

                return Ok(resultClass);
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass);
            }
        }

        //TODO 以下須跟著畫面異動

        /// <summary>
        /// 採購單報表 Procurement_Rpt
        /// </summary>
        [HttpPost("Procurement_Rpt")]
        public ActionResult<ResultClass<string>> Procurement_Rpt(string PM_Step)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"select PM_ID,PM_Step,(select item_D_name from Item_list where item_M_code = 'branch_company' and item_D_code = PM_BC) as PM_BC_Name
                    ,(select U_name from User_M where U_num = PM_U_num) as PM_Name
                    ,(select item_D_name from Item_list where item_M_code = 'Procurement_Pay' and item_D_code = PM_Pay_Type) as PM_Pay_Name
                    ,FORMAT(PM_Amt,'N0') as str_PM_Amt from Procurement_M";
                switch (PM_Step)
                {
                    case "N": 
                        T_SQL += " where 1=1";
                        break;
                    case "A":
                        T_SQL += " where PM_Step='A'";
                        break;
                    case "B":
                        T_SQL += " where PM_Step='B'";
                        break;
                    default:
                        break;
                }
                T_SQL += " order by PM_ID";
                #endregion
                var dtResult = _adoData.ExecuteSQuery(T_SQL);
                resultClass.ResultCode = "000";
                resultClass.objResult = JsonConvert.SerializeObject(dtResult);
                return Ok(resultClass);
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass); // 返回 500 錯誤碼
            }
        }

        /// <summary>
        /// 採購單報表明細 Procurement_DList_Rpt
        /// </summary>
        [HttpPost("Procurement_DList_Rpt")]
        public ActionResult<ResultClass<string>> Procurement_DList_Rpt(string PM_ID,string PM_Step)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"select * from Procurement_D where PM_ID=@PM_ID and PM_Step=@PM_Step";
                parameters.Add(new SqlParameter("@PM_ID", PM_ID));
                parameters.Add(new SqlParameter("@PM_Step", PM_Step));
                #endregion
                var dtResult=_adoData.ExecuteQuery(T_SQL, parameters);
                if (dtResult.Rows.Count > 0)
                {
                    resultClass.ResultCode = "000";
                    resultClass.objResult = JsonConvert.SerializeObject(dtResult);
                    return Ok(resultClass);
                }
                else
                {
                    resultClass.ResultCode = "400";
                    resultClass.ResultMsg = "查無資料";
                    return BadRequest(resultClass);
                }
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass); // 返回 500 錯誤碼
            }
        }

        /// <summary>
        /// 採購單報表Excel下載 Procurement_Excel
        /// </summary>
        [HttpGet("Procurement_Excel")]
        public IActionResult Procurement_Excel(string PM_Step)
        {
            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"select PM_ID,PM_Step,(select item_D_name from Item_list where item_M_code = 'branch_company' and item_D_code = PM_BC) as PM_BC_Name
                    ,(select U_name from User_M where U_num = PM_U_num) as PM_Name
                    ,(select item_D_name from Item_list where item_M_code = 'Procurement_Pay' and item_D_code = PM_Pay_Type) as PM_Pay_Name
                    ,FORMAT(PM_Amt,'N0') as str_PM_Amt from Procurement_M";
                switch (PM_Step)
                {
                    case "N":
                        T_SQL += " where 1=1";
                        break;
                    case "A":
                        T_SQL += " where PM_Step='A'";
                        break;
                    case "B":
                        T_SQL += " where PM_Step='B'";
                        break;
                    default:
                        break;
                }
                T_SQL += " order by PM_ID";
                #endregion
                var dtResult = _adoData.ExecuteSQuery(T_SQL);
                if (dtResult.Rows.Count > 0)
                {
                    var excelList = _adoData.ExecuteSQuery(T_SQL).AsEnumerable().Select(row => new Proc_M_Excel
                    {
                        PM_ID = row.Field<string>("PM_ID"),
                        PM_Step = row.Field<string>("PM_Step") == "A" ? "請購" : row.Field<string>("PM_Step") == "B" ? "請款" : row.Field<string>("PM_Step"),
                        PM_BC_Name = row.Field<string>("PM_BC_Name"),
                        PM_Name = row.Field<string>("PM_Name"),
                        PM_Pay_Name = row.Field<string>("PM_Pay_Name"),
                        str_PM_Amt = row.Field<string>("str_PM_Amt")
                    }).ToList();

                    var Excel_Headers = new Dictionary<string, string>
                    {
                        { "PM_ID","單號" },
                        { "PM_Step", "階段" },
                        { "PM_BC_Name", "部門" },
                        { "PM_Name", "請款人" },
                        { "PM_Pay_Name", "費用類別" },
                        { "str_PM_Amt","總價" }
                    };

                    var fileBytes = FuncHandler.ExportToExcel(excelList, Excel_Headers);
                    var fileName = "採購單資料" + DateTime.Now.ToString("yyyyMMddHHmm") + ".xlsx";
                    return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
                }
                else
                {
                    return NotFound();
                }
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex);
            }
        }

        
    }
}
