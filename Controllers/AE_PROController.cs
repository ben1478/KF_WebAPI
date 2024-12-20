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
        /// 採購單資料新增 Procurement_Ins
        /// </summary>
        [HttpPost("Procurement_Ins")]
        public ActionResult<ResultClass<string>> Procurement_Ins(Procurement_Ins model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL_Procurement_M
                var parameters_m = new List<SqlParameter>();
                var T_SQL_M = @"Insert into Procurement_M (PM_ID,PM_Step,PM_BC,PM_Pay_Type,PM_AppDate,PM_U_num,PM_Caption,PM_Amt,PM_Busin_Tax,PM_Tax_Amt
                    ,bnak_code,bank_name,branch_name,bank_account,payee_name,PM_Other,add_date,add_num,add_ip,edit_date,edit_num,edit_ip,PM_cknum)
                    Values (@PM_ID,@PM_Step,@PM_BC,@PM_Pay_Type,@PM_AppDate,@PM_U_num,@PM_Caption,@PM_Amt,@PM_Busin_Tax,@PM_Tax_Amt,@bnak_code,@bank_name
                    ,@branch_name,@bank_account,@payee_name,@PM_Other,GETDATE(),@add_num,@add_ip,GETDATE(),@edit_num,@edit_ip,@PM_cknum)";
                //取PM_ID
                var parameters_id = new List<SqlParameter>();
                var T_SQL_ID = @"exec GetFormID @formtype,@tablename";
                parameters_id.Add(new SqlParameter("@formtype", "PO"));
                parameters_id.Add(new SqlParameter("@tablename", "Procurement_M"));
                var result_id= _adoData.ExecuteQuery(T_SQL_ID, parameters_id);
                model.PM_ID = result_id.Rows[0]["resultPM_ID"].ToString();

                parameters_m.Add(new SqlParameter("@PM_ID", model.PM_ID));
                parameters_m.Add(new SqlParameter("@PM_Step", "A"));
                parameters_m.Add(new SqlParameter("@PM_BC", model.PM_BC));
                parameters_m.Add(new SqlParameter("@PM_Pay_Type", model.PM_Pay_Type));
                parameters_m.Add(new SqlParameter("@PM_AppDate", DateTime.Now.ToString("yyyy/MM/dd")));
                parameters_m.Add(new SqlParameter("@PM_U_num", model.PM_U_num));
                parameters_m.Add(new SqlParameter("@PM_Caption", model.PM_Caption));
                parameters_m.Add(new SqlParameter("@PM_Amt", model.PM_Amt));
                parameters_m.Add(new SqlParameter("@PM_Busin_Tax", model.PM_Busin_Tax));
                parameters_m.Add(new SqlParameter("@PM_Tax_Amt", model.PM_Tax_Amt));
                parameters_m.Add(new SqlParameter("@bnak_code", model.bnak_code));
                parameters_m.Add(new SqlParameter("@bank_name", model.bank_name));
                parameters_m.Add(new SqlParameter("@branch_name", model.branch_name));
                parameters_m.Add(new SqlParameter("@bank_account", model.bank_account));
                parameters_m.Add(new SqlParameter("@payee_name", model.payee_name));
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

                    //全部新增成功後要寫入請購審核流程
                    bool resultFlow = FuncHandler.AuditFlow(model.PM_U_num, "PO001", model.PM_ID, clientIp);
                    if(resultFlow)
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
                return StatusCode(500, resultClass); // 返回 500 錯誤碼
            }
           
        }
        /// <summary>
        /// 採購單列表查詢 Procurement_LQuery
        /// </summary>
        [HttpGet("Procurement_LQuery")]
        public ActionResult<ResultClass<string>> Procurement_LQuery()
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            var User_Num = HttpContext.Session.GetString("UserID");
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"select PM.PM_ID,PM_Step,LI.item_D_name,AD.FM_Step,LI.item_D_name as FM_Step_SignType,
                    (select COUNT(*) from ASP_UpLoad where cknum=PM.PM_cknum) as PM_cknum_count,PM.PM_cknum, PM.PM_Cancel,AD.FM_Step_Now
                    ,AD.FD_Step1_SignType
                    from Procurement_M PM 
                    inner join AuditFlow_D AD on AD.FD_Source_ID=PM.PM_ID AND ((PM.PM_Step = 'A' AND AD.FM_ID = 'PO001') OR (PM.PM_Step = 'B' AND AD.FM_ID = 'PO002'))
                    left join Item_list LI on LI.item_D_code=AD.FM_Step_SignType and LI.item_M_code='Flow_sign_type'
                    where PM.add_num = @User_Num ";
                parameters.Add(new SqlParameter("@User_Num", User_Num));
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
        /// 採購單單筆查詢Procurement_SQuery
        /// </summary>
        [HttpGet("Procurement_SQuery")]
        public ActionResult<ResultClass<string>> Procurement_SQuery(string PM_ID)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"select PM_BC,PM_Pay_Type,PM_Caption,PM_Amt,PM_Busin_Tax,PM_Tax_Amt,bnak_code,bank_name,branch_name,
                    bank_account,payee_name,PM_Other
                    from Procurement_M where PM_ID=@PM_ID";
                parameters.Add(new SqlParameter("@PM_ID", PM_ID));
                #endregion
                var dtResult = _adoData.ExecuteQuery(T_SQL, parameters);

                if (dtResult.Rows.Count > 0)
                {
                    var modellist = dtResult.AsEnumerable().Select(row => new Procurement_Res {
                        PM_BC = row.Field<string>("PM_BC"),
                        PM_Pay_Type = row.Field<string>("PM_Pay_Type"),
                        PM_Caption = row.Field<string>("PM_Caption"),
                        PM_Amt = row.Field<decimal>("PM_Amt"),
                        PM_Busin_Tax = row.Field<decimal>("PM_Busin_Tax"),
                        PM_Tax_Amt = row.Field<decimal>("PM_Tax_Amt"),
                        bnak_code = row.Field<string>("bnak_code"),
                        bank_name = row.Field<string>("bank_name"),
                        branch_name = row.Field<string>("branch_name"),
                        bank_account = row.Field<string>("bank_account"),
                        payee_name = row.Field<string>("payee_name"),
                        PM_Other = row.Field<string>("PM_Other")
                    }).FirstOrDefault();

                    #region SQL Procurement_D
                    var parameters_d = new List<SqlParameter>();
                    var T_SQL_D = @"select PD_ID,PD_Pro_name,PD_Unit,PD_Count,PD_Date,PD_Univalent,PD_Amt,PD_Company_name,PD_Est_Cost
                        from Procurement_D where PM_ID=@PM_ID";
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

                    modellist.Procurement_D = modelist_d.ToList();

                    resultClass.ResultCode = "000";
                    resultClass.objResult = JsonConvert.SerializeObject(modellist);
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
        public ActionResult<ResultClass<string>> Procurement_Canel(string PM_ID)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            var User_Num = HttpContext.Session.GetString("UserID");
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"update Procurement_M set PM_Cancel='Y',cancel_date=GETDATE(),cancel_num=@cancel_num,candel_ip=@candel_ip where PM_ID=@PM_ID";
                parameters.Add(new SqlParameter("@cancel_num", User_Num));
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
        /// 採購單明細單筆查詢 Procurement_D_SQuery
        /// </summary>
        [HttpGet("Procurement_D_SQuery")]
        public ActionResult<ResultClass<string>> Procurement_D_SQuery(string PD_ID)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL Procurement_D
                var parameters_d = new List<SqlParameter>();
                var T_SQL_D = @"select PD_ID,PD_Pro_name,PD_Unit,PD_Count,PD_Date,PD_Univalent,PD_Amt,PD_Company_name,PD_Est_Cost
                        from Procurement_D where PD_ID=@PD_ID";
                parameters_d.Add(new SqlParameter("@PD_ID", PD_ID));
                #endregion
                var dtResult_d = _adoData.ExecuteQuery(T_SQL_D, parameters_d);
                if(dtResult_d.Rows.Count > 0)
                {
                    var modelist_d = dtResult_d.AsEnumerable().Select(row => new Procurement_D_Res
                    {
                        PD_ID = row.Field<int>("PD_ID"),
                        PD_Pro_name = row.Field<string>("PD_Pro_name"),
                        PD_Unit = row.Field<string>("PD_Unit"),
                        PD_Count = row.Field<string>("PD_Count"),
                        PD_Date = FuncHandler.ConvertGregorianToROC(row.Field<string>("PD_Date")),
                        PD_Univalent = row.Field<decimal>("PD_Univalent"),
                        PD_Amt = row.Field<decimal>("PD_Amt"),
                        PD_Company_name = row.Field<string>("PD_Company_name"),
                        PD_Est_Cost = row.Field<decimal>("PD_Est_Cost")
                    }).FirstOrDefault();

                    resultClass.ResultCode = "000";
                    resultClass.objResult = JsonConvert.SerializeObject(modelist_d);
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
        /// 採購單明細單筆修改 Procurement_D_Upd
        /// </summary>
        [HttpPost("Procurement_D_Upd")]
        public ActionResult<ResultClass<string>> Procurement_D_Upd(Procurement_D_Res model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"update Procurement_D set PD_Pro_name=@PD_Pro_name,PD_Unit=@PD_Unit,PD_Count=@PD_Count,
                    PD_Date=@PD_Date,PD_Company_name=@PD_Company_name,PD_Univalent=@PD_Univalent,PD_Amt=@PD_Amt,PD_Est_Cost=@PD_Est_Cost where PD_ID=@PD_ID";
                parameters.Add(new SqlParameter("@PD_Pro_name", model.PD_Pro_name));
                parameters.Add(new SqlParameter("@PD_Unit", model.PD_Unit));
                parameters.Add(new SqlParameter("@PD_Count", model.PD_Count));
                parameters.Add(new SqlParameter("@PD_Date", FuncHandler.ConvertROCToGregorian(model.PD_Date)));
                parameters.Add(new SqlParameter("@PD_Company_name", model.PD_Company_name));
                parameters.Add(new SqlParameter("@PD_Univalent", model.PD_Univalent));
                parameters.Add(new SqlParameter("@PD_Amt", model.PD_Amt));
                parameters.Add(new SqlParameter("@PD_Est_Cost", model.PD_Est_Cost));
                parameters.Add(new SqlParameter("@PD_ID", model.PD_ID));
                #endregion
                int result_up = _adoData.ExecuteNonQuery(T_SQL, parameters);
                if (result_up == 0)
                {
                    resultClass.ResultCode = "400";
                    resultClass.ResultMsg = "變更失敗";
                    return BadRequest(resultClass);
                }
                else
                {
                    resultClass.ResultCode = "000";
                    resultClass.ResultMsg = "變更成功";
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
        [HttpPost("Procurement_Upd")]
        public ActionResult<ResultClass<string>> Procurement_Upd(Procurement_Res model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"update Procurement_M set PM_Pay_Type=@PM_Pay_Type,PM_Caption=@PM_Caption,PM_Amt=@PM_Amt,PM_Busin_Tax=@PM_Busin_Tax,PM_Tax_Amt=@PM_Tax_Amt
                    ,bnak_code=@bnak_code,bank_name=@bank_name,branch_name=@branch_name,bank_account=@bank_account,payee_name=@payee_name
                    ,PM_Other=@PM_Other where PM_ID=@PM_ID";
                parameters.Add(new SqlParameter("@PM_ID", model.PM_ID));
                parameters.Add(new SqlParameter("@PM_Pay_Type", model.PM_Pay_Type));
                parameters.Add(new SqlParameter("@PM_Caption", model.PM_Caption));
                parameters.Add(new SqlParameter("@PM_Amt", model.PM_Amt));
                parameters.Add(new SqlParameter("@PM_Busin_Tax", model.PM_Busin_Tax));
                parameters.Add(new SqlParameter("@PM_Tax_Amt", model.PM_Tax_Amt));
                parameters.Add(new SqlParameter("@bnak_code", model.bnak_code));
                parameters.Add(new SqlParameter("@bank_name", model.bank_name));
                parameters.Add(new SqlParameter("@branch_name", model.branch_name));
                parameters.Add(new SqlParameter("@bank_account", model.bank_account));
                parameters.Add(new SqlParameter("@payee_name", model.payee_name));
                if(!string.IsNullOrEmpty(model.PM_Other))
                {
                    parameters.Add(new SqlParameter("@PM_Other", model.PM_Other));
                }
                else
                {
                    parameters.Add(new SqlParameter("@PM_Other",DBNull.Value));
                }
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
                    resultClass.ResultCode = "000";
                    resultClass.ResultMsg = "變更成功";
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
        /// 採購單審核列表查詢 ProcFlow_LQuery
        /// </summary>
        [HttpPost("ProcFlow_LQuery")]
        public ActionResult<ResultClass<string>> ProcFlow_LQuery(ProcFlow_Req model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"select PM.PM_ID,PM_Step,LI.item_D_name,AD.FM_Step,LI.item_D_name as FM_Step_SignType,
                    PM.PM_cknum,(select COUNT(*) from ASP_UpLoad where cknum=PM.PM_cknum) as PM_cknum_count
                    ,PM.PM_Cancel
                    from Procurement_M PM 
                    inner join AuditFlow_D AD on AD.FD_Source_ID=PM.PM_ID AND ((PM.PM_Step = 'A' AND AD.FM_ID = 'PO001') OR (PM.PM_Step = 'B' AND AD.FM_ID = 'PO002'))
                    left join Item_list LI on LI.item_D_code=AD.FM_Step_SignType and LI.item_M_code='Flow_sign_type'
                    where PM.add_date between @PD_Date_S and @PD_Date_E";
                parameters.Add(new SqlParameter("@PD_Date_S", FuncHandler.ConvertROCToGregorian(model.PD_Date_S)));
                parameters.Add(new SqlParameter("@PD_Date_E", FuncHandler.ConvertROCToGregorian(model.PD_Date_E)));
                if (!string.IsNullOrEmpty(model.PM_BC))
                {
                    T_SQL += " and PM.PM_BC=@PM_BC";
                    parameters.Add(new SqlParameter("@PM_BC",model.PM_BC));
                }
                if(!string.IsNullOrEmpty(model.PM_U_num))
                {
                    T_SQL += " and PM_U_num=@PM_U_num";
                    parameters.Add(new SqlParameter("@PM_U_num", model.PM_U_num));
                }
                #endregion
                var dtResult=_adoData.ExecuteQuery(T_SQL,parameters);
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
        ///<summary>
        /// 採購單列印資料抓取
        /// </summary>
        [HttpPost("ProcForm_Query")]
        public ActionResult<ResultClass<string>> ProcForm_Query(string PM_ID)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"select (select item_D_name from Item_list where item_M_code='branch_company' and item_D_type='Y' and item_D_code=PM_BC) as U_BC_Name
                    ,PM_ID,(select item_D_name from Item_list where item_M_code='Procurement_Pay' and item_D_type='Y' and item_D_code=PM_Pay_Type) as PM_Pay_Name
                    ,CONVERT(VARCHAR, add_date, 111) as AppDate,PM_Caption,FORMAT(PM_Amt,'N0') as PM_Amt,FORMAT(PM_Busin_Tax,'N0') as PM_Busin_Tax
                    ,FORMAT(PM_Tax_Amt,'N0') as PM_Tax_Amt,'('+bnak_code+')'+bank_name as bank_name,branch_name,bank_account,payee_name,PM_Other
                    from Procurement_M where PM_ID=@PM_ID";
                parameters.Add(new SqlParameter("@PM_ID", PM_ID));
                #endregion
                var model = _adoData.ExecuteQuery(T_SQL,parameters).AsEnumerable().Select(row=> new ProcForm_M {
                    U_BC_Name = row.Field<string>("U_BC_Name"),
                    PM_ID = row.Field<string>("PM_ID"),
                    PM_Pay_Name = row.Field<string>("PM_Pay_Name"),
                    AppDate = row.Field<string>("AppDate"),
                    PM_Caption = row.Field<string>("PM_Caption"),
                    PM_Amt = row.Field<string>("PM_Amt"),
                    PM_Busin_Tax = row.Field<string>("PM_Busin_Tax"),
                    PM_Tax_Amt = row.Field<string>("PM_Tax_Amt"),
                    bank_name = row.Field<string>("bank_name"),
                    branch_name = row.Field<string>("branch_name"),
                    bank_account = row.Field<string>("bank_account"),
                    payee_name = row.Field<string>("payee_name"),
                    PM_Other = row.Field<string>("PM_Other")
                }).First();

                #region Procurement_D
                var parameters_d = new List<SqlParameter>();
                var T_SQL_D = @"select PD_Pro_name,PD_Unit,FORMAT(CAST(PD_Count AS INT), 'N0') as PD_Count,PD_Date,FORMAT(PD_Univalent,'N0') as PD_Univalent
                    ,FORMAT(PD_Amt,'N0') as PD_Amt,PD_Company_name,FORMAT(PD_Est_Cost,'N0') as PD_Est_Cost
                    from Procurement_D where PM_ID=@PM_ID";
                parameters_d.Add(new SqlParameter("@PM_ID", PM_ID));
                #endregion
                var modelList=_adoData.ExecuteQuery(T_SQL_D,parameters_d).AsEnumerable().Select(row=> new ProcForm_D {
                    PD_Pro_name = row.Field<string>("PD_Pro_name"),
                    PD_Unit = row.Field<string>("PD_Unit"),
                    PD_Count = row.Field<string>("PD_Count"),
                    PD_Date = row.Field<string>("PD_Date"),
                    PD_Univalent = row.Field<string>("PD_Univalent"),
                    PD_Amt = row.Field<string>("PD_Amt"),
                    PD_Company_name = row.Field<string>("PD_Company_name"),
                    PD_Est_Cost = row.Field<string>("PD_Est_Cost")
                }).ToList();

                model.ProcFormDList= modelList;
                resultClass.ResultCode = "000";
                resultClass.objResult = JsonConvert.SerializeObject(model);
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
        /// 請購轉請款
        /// </summary>
        [HttpGet("ProcTurn")]
        public ActionResult<ResultClass<string>> ProcTurn(string PM_ID,string User)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"update Procurement_M set PM_Step='B',edit_date=GETDATE(),edit_num=@edit_num,edit_ip=@edit_ip where PM_ID=@PM_ID";
                parameters.Add(new SqlParameter("@edit_num", User));
                parameters.Add(new SqlParameter("@edit_ip", clientIp));
                parameters.Add(new SqlParameter("@PM_ID", PM_ID));
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
                    //寫入請款審核流程
                    bool resultFlow = FuncHandler.AuditFlow(User, "PO002", PM_ID, clientIp);
                    if (resultFlow)
                    {
                        resultClass.ResultCode = "000";
                        resultClass.ResultMsg = "變更成功";
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
                return StatusCode(500, resultClass); // 返回 500 錯誤碼
            }
        }
    }
}
