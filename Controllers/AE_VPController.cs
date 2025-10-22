using KF_WebAPI.BaseClass;
using KF_WebAPI.BaseClass.AE;
using KF_WebAPI.BaseClass.Max104;
using KF_WebAPI.FunctionHandler;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Newtonsoft.Json;
using System;
using System.Data;
using System.Diagnostics.Eventing.Reader;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace KF_WebAPI.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class AE_VPController : ControllerBase
    {
        /// <summary>
        /// 查詢已核准的請款或請採購
        /// </summary>
        [HttpGet("Fina_form_LQuery")]
        public ActionResult<ResultClass<string>> Fina_form_LQuery(string User, string Type)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = "";
                switch (Type)
                {
                    case "PO":
                        T_SQL += @"select 'PO' as FormType,PM.PM_ID as FormID,PD.PD_Pro_name as FormCaption,PD.PD_Amt as FormMoney 
                                   from Procurement_M PM
                                   inner join Procurement_D PD on PM.PM_ID = PD.PM_ID
                                   inner join AuditFlow_M AM on AM.FM_Source_ID = PM.PM_ID and AM.FM_Step='9' and AM.FM_Step_SignType ='FSIGN002'
                                   where VP_ID is null and PM_U_num=@User";
                        break;
                    case "PS":
                        T_SQL += @"select VM.VP_type as FormType,VM.VP_ID as FormID,VD.VD_Fee_Summary as FormCaption,VD.VD_Fee as FormMoney    
                                   from InvPrepay_M VM
                                   inner join InvPrepay_D VD on VM.VP_ID=VD.VP_ID and not exists (select 1 from InvPrepay_D where Form_ID=VM.VP_ID)
                                   inner join AuditFlow_M AM on AM.FM_Source_ID = VM.VP_ID and AM.FM_Step='9' and AM.FM_Step_SignType ='FSIGN002'
                                   where VM.VP_type = 'PP' and VP_U_num =@User";
                        break;
                    default:
                        break;
                }
                var parameters = new List<SqlParameter> 
                {
                    new SqlParameter("@User", User)
                };
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
                return StatusCode(500, resultClass);
            }
        }

        /// <summary>
        /// 新增請款(預支)單 InvPrepay_M_Ins
        /// </summary>
        [HttpPost("InvPrepay_M_Ins")]
        public ActionResult<ResultClass<string>> InvPrepay_M_Ins(InvPrepay_Ins model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();
            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                //取VP_ID
                var T_SQL_ID = @"exec GetFormID @formtype,@tablename,@New_ID output";
                var str_type = "PA";
                if (model.VP_type != null && model.VP_type.Length > 0 )
                {
                    if (model.VP_type.Contains("PP"))
                        str_type = "PP";
                    if (model.VP_type.Contains("PS"))
                        str_type = "PS";
                }
                var parameters_id = new List<SqlParameter> 
                {
                    new SqlParameter("@formtype", str_type),
                    new SqlParameter("@tablename", "InvPrepay_M")
                };
                SqlParameter newIdParameter = new SqlParameter("@New_ID", SqlDbType.VarChar, 20)
                {
                    Direction = ParameterDirection.Output
                };
                parameters_id.Add(newIdParameter);

                var result_id = _adoData.ExecuteQuery(T_SQL_ID, parameters_id);
                var VP_ID = newIdParameter.Value.ToString();
                var T_SQL_M = @"Insert into InvPrepay_M (VP_ID,VP_type,VP_COM,VP_BC,VP_Pay_Type,VP_AppDate,VP_U_num,VP_Total_Money,bank_code,bank_name,
                    branch_name,bank_account,payee_name,VP_Summary,VP_MFG_Date,VP_Nsummary,VP_cknum,add_date,add_num,add_ip,edit_date,edit_num,edit_ip) 
                    Values (@VP_ID,@VP_type,@VP_COM,@VP_BC,@VP_Pay_Type,FORMAT(GETDATE(),'yyyy/MM/dd'),@VP_U_num,@VP_Total_Money,@bank_code,@bank_name,@branch_name,
                    @bank_account,@payee_name,@VP_Summary,@VP_MFG_Date,@VP_Nsummary,@VP_cknum,GETDATE(),@add_num,@add_ip,GETDATE(),@edit_num,@edit_ip)";
                var parameters_m = new List<SqlParameter> 
                {
                    new SqlParameter("@VP_ID", VP_ID),
                    new SqlParameter("@VP_type", str_type),
                    new SqlParameter("@VP_COM", model.VP_COM),
                    new SqlParameter("@VP_BC", model.VP_BC),
                    new SqlParameter("@VP_Pay_Type", model.VP_Pay_Type),
                    new SqlParameter("@VP_U_num", model.User),
                    new SqlParameter("@VP_Total_Money", model.VP_Total_Money),
                    new SqlParameter("@bank_code", string.IsNullOrEmpty(model.bank_code) ? DBNull.Value : model.bank_code),
                    new SqlParameter("@bank_name", string.IsNullOrEmpty(model.bank_name) ? DBNull.Value : model.bank_name),
                    new SqlParameter("@branch_name", string.IsNullOrEmpty(model.branch_name) ? DBNull.Value : model.branch_name),
                    new SqlParameter("@bank_account", string.IsNullOrEmpty(model.bank_account) ? DBNull.Value : model.bank_account),
                    new SqlParameter("@payee_name", string.IsNullOrEmpty(model.payee_name) ? DBNull.Value : model.payee_name),
                    new SqlParameter("@VP_Summary", string.IsNullOrEmpty(model.VP_Summary) ? DBNull.Value : model.VP_Summary),
                    new SqlParameter("@VP_MFG_Date", string.IsNullOrEmpty(model.VP_MFG_Date) ? DBNull.Value : model.VP_MFG_Date),
                    new SqlParameter("@VP_Nsummary", string.IsNullOrEmpty(model.VP_Nsummary) ? DBNull.Value : model.VP_Nsummary),
                    new SqlParameter("@VP_cknum", FuncHandler.GetCheckNum()),
                    new SqlParameter("@add_num", model.User),
                    new SqlParameter("@add_ip", clientIp),
                    new SqlParameter("@edit_num", model.User),
                    new SqlParameter("@edit_ip", clientIp)
                };
                #endregion
                int result_m = _adoData.ExecuteNonQuery(T_SQL_M, parameters_m);
                if(result_m == 0) 
                {
                    resultClass.ResultCode = "400";
                    resultClass.ResultMsg = "主檔新增失敗";
                    return BadRequest(resultClass);
                }
                else
                {
                    foreach (var item in model.Ins_List)
                    {
                        #region SQL
                        var T_SQL_D = @"Insert into InvPrepay_D (VP_ID,VP_type,Form_ID,VD_BC,VD_Fee_Summary,VD_Fee,VD_VAT,VD_Account_code,VD_Account,
                            add_date,add_num,add_ip,edit_date,edit_num,edit_ip,Change_reason,Change_num,Change_date,Change_ip) 
                            Values (@VP_ID,@VP_type,@Form_ID,@VD_BC,@VD_Fee_Summary,@VD_Fee,@VD_VAT,@VD_Account_code,@VD_Account,GETDATE(),@add_num,@add_ip,
                            GETDATE(),@edit_num,@edit_ip,@Change_reason,@Change_num,@Change_date,@Change_ip)";

                        var parameters_d = new List<SqlParameter> 
                        {
                            new SqlParameter("@VP_ID", VP_ID),
                            new SqlParameter("@VP_type", str_type),
                            new SqlParameter("@Form_ID", string.IsNullOrEmpty(item.FormID) ? DBNull.Value : item.FormID),
                            new SqlParameter("@VD_BC",  item.VD_BC),
                            new SqlParameter("@VD_Fee_Summary", item.FormCaption),
                            new SqlParameter("@VD_Fee", item.FormMoney),
                            new SqlParameter("@VD_VAT", item.VD_VAT),
                            new SqlParameter("@add_num", model.User),
                            new SqlParameter("@VD_Account_code", string.IsNullOrEmpty(item.VD_Account_code) ? DBNull.Value : item.VD_Account_code),
                            new SqlParameter("@VD_Account", string.IsNullOrEmpty(item.VD_Account) ? DBNull.Value : item.VD_Account),
                            new SqlParameter("@add_ip", clientIp),
                            new SqlParameter("@edit_num", model.User),
                            new SqlParameter("@edit_ip", clientIp)
                        };
                        if (!string.IsNullOrEmpty(item.ChangeReason))
                        {
                            parameters_d.Add(new SqlParameter("@Change_reason", item.ChangeReason));
                            parameters_d.Add(new SqlParameter("@Change_num", model.User));
                            parameters_d.Add(new SqlParameter("@Change_date", DateTime.Now));
                            parameters_d.Add(new SqlParameter("@Change_ip", clientIp));
                        }
                        else
                        {
                            parameters_d.Add(new SqlParameter("@Change_reason", DBNull.Value));
                            parameters_d.Add(new SqlParameter("@Change_num", DBNull.Value));
                            parameters_d.Add(new SqlParameter("@Change_date", DBNull.Value));
                            parameters_d.Add(new SqlParameter("@Change_ip", DBNull.Value));
                        }
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
                    bool resultFlow = FuncHandler.AuditFlow(model.User, model.VP_BC, str_type, VP_ID, clientIp);
                    if (resultFlow)
                    {
                        resultClass.ResultCode = "000";
                        resultClass.ResultMsg = "新增成功";

                        //勾稽請採購單或預支單
                        var distinctFormIDs = model.Ins_List.Where(item => !string.IsNullOrEmpty(item.FormID))
                                        .GroupBy(item => item.FormID).Select(Group => Group.Key).ToList();
                        if(distinctFormIDs !=null && distinctFormIDs.Count() > 0)
                        {
                            if(str_type =="PA" || str_type == "PP")
                            {
                                foreach (var item in distinctFormIDs)
                                {
                                    var T_SQL_P = @"Update Procurement_M set VP_ID=@VP_ID where PM_ID=@Form_ID";
                                    var parameters_p = new List<SqlParameter> 
                                    {
                                        new SqlParameter("@VP_ID", VP_ID),
                                        new SqlParameter("@Form_ID", item)
                                    };
                                    _adoData.ExecuteQuery(T_SQL_P, parameters_p);
                                }
                            }
                        }
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
        /// 請款(預支)單列表查詢 InvPrepay_M_LQuery
        /// </summary>
        [HttpGet("InvPrepay_M_LQuery")]
        public ActionResult<ResultClass<string>> InvPrepay_M_LQuery(string User)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"select VP.VP_ID,AM.FM_Step,LI.item_D_name,LI.item_D_name AS FM_Step_SignType,VP.VP_cknum,VP.VP_Cancel
                    ,(SELECT COUNT(*) FROM ASP_UpLoad WHERE cknum = VP.VP_cknum and del_tag='0') AS VP_cknum_count 
                    from InvPrepay_M VP
                    INNER JOIN AuditFlow_M AM ON AM.FM_Source_ID = VP.VP_ID and AM.AF_ID = VP.VP_type
                    LEFT JOIN Item_list LI ON LI.item_D_code = AM.FM_Step_SignType AND LI.item_M_code = 'Flow_sign_type'
                    where VP.add_num = @User order by VP.add_Date desc";
                var parameters = new List<SqlParameter> 
                {
                    new SqlParameter("@User", User)
                };
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
        /// 請款(預支)單單筆查詢 InvPrepay_M_SQuery
        /// </summary>
        [HttpGet("InvPrepay_M_SQuery")]
        public ActionResult<ResultClass<string>> InvPrepay_M_SQuery(string VP_ID)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"select VP_type,VP_COM,VP_BC,VP_Pay_Type,VP_Total_Money,bank_code,bank_name,branch_name,bank_account,
                    payee_name,VP_Summary,VP_MFG_Date,VP_Nsummary 
                    from InvPrepay_M where VP_ID=@VP_ID ";
                var parameters = new List<SqlParameter> 
                {
                    new SqlParameter("@VP_ID", VP_ID)
                };
                #endregion
                var dtResult = _adoData.ExecuteQuery(T_SQL, parameters);

                if (dtResult.Rows.Count > 0)
                {
                    var model = dtResult.AsEnumerable().Select(row => new InvPrepay_Ins
                    {
                        VP_type = new string[] { row.Field<string>("VP_type") },
                        VP_COM = row.Field<string>("VP_COM"),
                        VP_BC = row.Field<string>("VP_BC"),
                        VP_Pay_Type = row.Field<string>("VP_Pay_Type"),
                        VP_Total_Money = row.Field<int>("VP_Total_Money").ToString(),
                        bank_code = row.Field<string>("bank_code"),
                        bank_name = row.Field<string>("bank_name"),
                        branch_name = row.Field<string>("branch_name"),
                        bank_account = row.Field<string>("bank_account"),
                        payee_name = row.Field<string>("payee_name"),
                        VP_Summary = row.Field<string>("VP_Summary"),
                        VP_MFG_Date = row.Field<string>("VP_MFG_Date"),
                        VP_Nsummary = row.Field<string>("VP_Nsummary")
                    }).FirstOrDefault();

                    #region SQL Procurement_D
                    var T_SQL_D = @"select VD_ID,Form_ID,VD_BC,VD_Fee_Summary,VD_Fee,VD_VAT,VD_Account_code,VD_Account from InvPrepay_D where VP_ID=@VP_ID ";
                    var parameters_d = new List<SqlParameter> 
                    {
                        new SqlParameter("@VP_ID", VP_ID)
                    };
                    #endregion
                    var dtResult_d = _adoData.ExecuteQuery(T_SQL_D, parameters_d);
                    
                    model.Ins_List = dtResult_d.AsEnumerable().Select(row => new InvPrepay_D_Ins
                    {
                        VD_ID = row.Field<int>("VD_ID").ToString(),
                        FormID = row.Field<string>("Form_ID"),
                        VD_BC = row.Field<string>("VD_BC"),
                        FormCaption = row.Field<string>("VD_Fee_Summary"),
                        FormMoney = row.Field<int>("VD_Fee").ToString(),
                        VD_VAT = row.Field<string>("VD_VAT"),
                        VD_Account_code = row.Field<string>("VD_Account_code"),
                        VD_Account = row.Field<string>("VD_Account")
                    }).ToList();

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
                return StatusCode(500, resultClass); 
            }
        }

        /// <summary>
        /// 修改請款(預支)單 InvPrepay_M_Upd
        /// </summary>
        [HttpPost("InvPrepay_M_Upd")]
        public ActionResult<ResultClass<string>> InvPrepay_M_Upd(InvPrepay_Ins model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();
            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var str_type = "PA";
                if (model.VP_type != null && model.VP_type.Length > 0)
                {
                    if (model.VP_type.Contains("PP"))
                        str_type = "PP";
                    if (model.VP_type.Contains("PS"))
                        str_type = "PS";
                }
                var T_SQL_M = @"Update InvPrepay_M set VP_COM=@VP_COM,VP_type=@VP_type,VP_Pay_Type=@VP_Pay_Type,VP_Total_Money=@VP_Total_Money,bank_code=@bank_code,
                                bank_name=@bank_name,branch_name=@branch_name,bank_account=@bank_account,payee_name=@payee_name,VP_Summary=@VP_Summary,VP_MFG_Date=@VP_MFG_Date,
                                VP_Nsummary=@VP_Nsummary,edit_date=GETDATE(),edit_num=@edit_num,edit_ip=@edit_ip where VP_ID=@VP_ID";
                var parameters_m = new List<SqlParameter> 
                {
                    new SqlParameter("@VP_COM",model.VP_COM),
                    new SqlParameter("@VP_type",str_type),
                    new SqlParameter("@VP_Pay_Type", model.VP_Pay_Type),
                    new SqlParameter("@VP_Total_Money", model.VP_Total_Money),
                    new SqlParameter("@bank_code", string.IsNullOrEmpty(model.bank_code) ? DBNull.Value : model.bank_code),
                    new SqlParameter("@bank_name", string.IsNullOrEmpty(model.bank_name) ? DBNull.Value : model.bank_name),
                    new SqlParameter("@branch_name", string.IsNullOrEmpty(model.branch_name) ? DBNull.Value : model.branch_name),
                    new SqlParameter("@bank_account", string.IsNullOrEmpty(model.bank_account) ? DBNull.Value : model.bank_account),
                    new SqlParameter("@payee_name", string.IsNullOrEmpty(model.payee_name) ? DBNull.Value : model.payee_name),
                    new SqlParameter("@VP_Summary", string.IsNullOrEmpty(model.VP_Summary) ? DBNull.Value : model.VP_Summary),
                    new SqlParameter("@VP_MFG_Date", string.IsNullOrEmpty(model.VP_MFG_Date) ? DBNull.Value : model.VP_MFG_Date),
                    new SqlParameter("@VP_Nsummary", string.IsNullOrEmpty(model.VP_Nsummary) ? DBNull.Value : model.VP_Nsummary),
                    new SqlParameter("@edit_num", model.User),
                    new SqlParameter("@edit_ip", clientIp),
                    new SqlParameter("@VP_ID", model.VP_ID)
                };
                #endregion
                int result_m = _adoData.ExecuteNonQuery(T_SQL_M, parameters_m);
                if (result_m == 0)
                {
                    resultClass.ResultCode = "400";
                    resultClass.ResultMsg = "主檔修改失敗";
                    return BadRequest(resultClass);
                }
                else
                {
                    foreach (var item in model.Ins_List)
                    {
                        #region SQL
                        var parameters_d = new List<SqlParameter>();
                        var T_SQL_D = "";
                        if (!string.IsNullOrEmpty(item.VD_ID))
                        {
                            T_SQL_D += @"Update InvPrepay_D set Form_ID=@Form_ID,VD_BC=@VD_BC,VD_Fee_Summary=@VD_Fee_Summary,VD_Fee=@VD_Fee,VD_VAT=@VD_VAT, 
                            VD_Account_code=@VD_Account_code,VD_Account=@VD_Account,edit_date=GETDATE(),edit_num=@edit_num,edit_ip=@edit_ip";

                            if (!string.IsNullOrEmpty(item.ChangeReason))
                            {
                                T_SQL_D += ",Change_reason=@Change_reason,Change_num=@Change_num,Change_date=GETDATE(),Change_ip=@Change_ip";
                                parameters_d.AddRange(new List<SqlParameter> 
                                {
                                    new SqlParameter("@Change_reason", item.ChangeReason),
                                    new SqlParameter("@Change_num", model.User),
                                    new SqlParameter("@Change_ip", clientIp)
                                });
                            }

                            T_SQL_D += " where VD_ID=@VD_ID and VP_ID=@VP_ID";
                        }
                        else
                        {
                            T_SQL_D += @"Insert into InvPrepay_D (VP_ID,VP_type,Form_ID,VD_BC,VD_Fee_Summary,VD_Fee,VD_VAT,VD_Account_code,
                            VD_Account,add_date,add_num,add_ip) Values (@VP_ID,@VP_type,@Form_ID,@VD_BC,@VD_Fee_Summary,@VD_Fee,@VD_VAT,
                            @VD_Account_code,@VD_Account,GETDATE(),@add_num,@add_ip)";
                            parameters_d.AddRange(new List<SqlParameter> 
                            {
                                new SqlParameter("@VP_type", str_type),
                                new SqlParameter("@add_num", model.User),
                                new SqlParameter("@add_ip", clientIp)
                            });
                        }
                        parameters_d.AddRange(new List<SqlParameter> 
                        {
                            new SqlParameter("@Form_ID", string.IsNullOrEmpty(item.FormID) ? DBNull.Value : item.FormID),
                            new SqlParameter("@VD_BC",item.VD_BC),
                            new SqlParameter("@VD_Fee_Summary", item.FormCaption),
                            new SqlParameter("@VD_Fee", item.FormMoney),
                            new SqlParameter("@VD_VAT", item.VD_VAT),
                            new SqlParameter("@VD_Account_code", string.IsNullOrEmpty(item.VD_Account_code) ? DBNull.Value : item.VD_Account_code),
                            new SqlParameter("@VD_Account", string.IsNullOrEmpty(item.VD_Account) ? DBNull.Value : item.VD_Account),
                            new SqlParameter("@edit_num", model.User),
                            new SqlParameter("@edit_ip", clientIp),
                            new SqlParameter("@VD_ID", item.VD_ID),
                            new SqlParameter("@VP_ID", model.VP_ID)
                        });
                        #endregion
                        int result_d = _adoData.ExecuteNonQuery(T_SQL_D, parameters_d);
                        if (result_d == 0)
                        {
                            resultClass.ResultCode = "400";
                            resultClass.ResultMsg = "明細檔修改失敗";
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
                return StatusCode(500, resultClass);
            }
            
        }

        /// <summary>
        /// 刪除單筆明細
        /// </summary>
        [HttpDelete("InvPrepay_D_Del")]
        public ActionResult<ResultClass<string>> InvPrepay_D_Del(string VD_ID)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"Delete InvPrepay_D where VD_ID=@VD_ID";
                var parameters = new List<SqlParameter>
                {
                    new SqlParameter("@VD_ID",VD_ID)
                };
                #endregion
                int result = _adoData.ExecuteNonQuery(T_SQL, parameters);
                if (result == 0) 
                {
                    resultClass.ResultCode = "400";
                    resultClass.ResultMsg = "明細刪除失敗";
                    return BadRequest(resultClass);
                } 
                else
                {
                    resultClass.ResultCode = "000";
                    resultClass.ResultMsg = "明細刪除成功";
                    return Ok(resultClass);
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
        /// 請款預支單抽單 InvPrepay_Canel
        /// </summary>
        [HttpGet("InvPrepay_Canel")]
        public ActionResult<ResultClass<string>> InvPrepay_Canel(string VP_ID, string User)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();
            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"update InvPrepay_M set VP_Cancel='Y',cancel_date=GETDATE(),cancel_num=@cancel_num,candel_ip=@candel_ip where VP_ID=@VP_ID";
                var parameters = new List<SqlParameter> 
                {
                    new SqlParameter("@cancel_num", User),
                    new SqlParameter("@candel_ip", clientIp),
                    new SqlParameter("@VP_ID", VP_ID)
                };
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
                return StatusCode(500, resultClass);
            }
        }

        /// <summary>
        /// 財會單位判定 ChkUserRole
        /// </summary>
        [HttpGet("ChkUserRole")]
        public ActionResult<ResultClass<string>> ChkUserRole(string User)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"select * from User_M where U_num=@User and Role_num in ('1004','1005','1020','1006','1007')";
                var parameters = new List<SqlParameter> 
                {
                    new SqlParameter("@User", User)
                };
                #endregion
                var dtresult = _adoData.ExecuteQuery(T_SQL, parameters);
                resultClass.ResultCode = "000";
                if (dtresult.Rows.Count == 0)
                {
                    resultClass.objResult = "N";
                }
                else
                {
                    resultClass.objResult = "Y";
                }
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
        /// 取得請款單列印資料 GetPrintData
        /// </summary>
        [HttpGet("GetPrintData")]
        public ActionResult<ResultClass<string>> GetPrintData(string VP_ID)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL_M = @"select LI.item_d_name as BC_Name,UM.U_name,VP_ID,VP.VP_AppDate,VP.VP_Pay_Type,bank_code,bank_name,branch_name,
                    bank_account,payee_name,Format(VP_Total_Money,'N0') as VP_Total_Money,VP.VP_type,
                    FORMAT(VP.add_date,'yyyy/MM/dd HH:mm') as add_date,
                    case when exists (select 1 from InvPrepay_D VD where VD.VP_ID = VP.VP_ID and LEFT(VD.Form_ID, '2') = 'PO' ) then 'PO' else null end as VP_type_PO
                    from InvPrepay_M VP
                    left join User_M UM on UM.U_num = VP.VP_U_num
                    left join Item_list LI on LI.item_D_code = VP.VP_BC
                    where VP.VP_ID=@VP_ID";
                var parameters_m = new List<SqlParameter> 
                {
                    new SqlParameter("@VP_ID", VP_ID)
                };
                #endregion
                var resultModel = _adoData.ExecuteQuery(T_SQL_M, parameters_m).AsEnumerable().Select(row => new Invp_Print
                {
                    BC_Name = row.Field<string>("BC_Name"),
                    U_name = row.Field<string>("U_name"),
                    VP_ID = row.Field<string>("VP_ID"),
                    VP_AppDate = row.Field<string>("VP_AppDate"),
                    VP_Pay_Type = row.Field<string>("VP_Pay_Type"),
                    bank_code = row.Field<string>("bank_code"),
                    bank_name = row.Field<string>("bank_name"),
                    branch_name = row.Field<string>("branch_name"),
                    bank_account = row.Field<string>("bank_account"),
                    payee_name = row.Field<string>("payee_name"),
                    VP_Total_Money = row.Field<string>("VP_Total_Money"),
                    VP_type = row.Field<string>("VP_type"),
                    VP_type_PO = row.Field<string>("VP_type_PO"),
                    add_date = row.Field<string>("add_date")
                }).FirstOrDefault();

                #region Invp_Print_Deatil
                var T_SQL_DT = @"select Form_ID,VD_Fee_Summary,Format(VD_Fee,'N0') as VD_Fee from InvPrepay_D where VP_ID=@VP_ID";
                var parameters_dt = new List<SqlParameter> 
                {
                    new SqlParameter("@VP_ID", VP_ID)
                };
                resultModel.VP_Deatil_List = _adoData.ExecuteQuery(T_SQL_DT, parameters_dt).AsEnumerable().Select(row => new Invp_Print_Deatil
                {
                    Form_ID = row.Field<string>("Form_ID"),
                    VD_Fee_Summary = row.Field<string>("VD_Fee_Summary"),
                    VD_Fee = row.Field<string>("VD_Fee")
                }).ToList();
                #endregion

                #region Invp_Print_Flow
                var T_SQL_FW = @"select AD.FD_Sign_Countersign,AD.FD_Step,UM.U_name,
                    FORMAT(AD.FD_Step_date,'yyyy/MM/dd HH:mm') as FD_Step_date,AD.FD_Step_desc 
                    from AuditFlow_D AD
                    inner join User_M UM on UM.U_num = AD.FD_Step_num
                    where FM_Source_ID=@VP_ID and FD_Sign_Countersign='S' order by AD.FD_Step  ";
                var parameters_fw = new List<SqlParameter> 
                {
                    new SqlParameter("@VP_ID", VP_ID)
                };
                resultModel.VP_Flow_List = _adoData.ExecuteQuery(T_SQL_FW, parameters_fw).AsEnumerable().Select(row => new Invp_Print_Flow
                {
                    FD_Sign_Countersign = row.Field<string>("FD_Sign_Countersign"),
                    FD_Step = row.Field<string>("FD_Step"),
                    U_name = row.Field<string>("U_name"),
                    FD_Step_date = row.Field<string>("FD_Step_date"),
                    FD_Step_desc = row.Field<string>("FD_Step_desc")
                }).ToList();
                #endregion

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

        /// <summary>
        /// 提供所有文中系統分公司清單 item_M_code='winton_company'
        /// </summary>
        [HttpGet("GetWintonBCList")]
        public ActionResult<ResultClass<string>> GetWintonBCList(string User)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"select Li.item_D_code, Li.item_D_name, UWG.Winton_Group as WBC 
                    from Item_list Li left join User_Winton_Group UWG on UWG.U_num = @User 
                    where item_M_code='winton_company' and item_D_type='Y' and show_tag ='0'";
                var parameters = new List<SqlParameter>
                {
                    new SqlParameter("@User", User)
                };
                #endregion
                DataTable dtResult = _adoData.ExecuteQuery(T_SQL,parameters);
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
                return StatusCode(500, resultClass);
            }
        }
    }
}
