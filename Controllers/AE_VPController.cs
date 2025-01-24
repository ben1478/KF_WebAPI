using KF_WebAPI.BaseClass;
using KF_WebAPI.BaseClass.AE;
using KF_WebAPI.FunctionHandler;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Newtonsoft.Json;
using System;
using System.Data;
using System.Diagnostics.Eventing.Reader;
using System.Text.RegularExpressions;

namespace KF_WebAPI.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class AE_VPController : ControllerBase
    {
        [HttpGet("Fina_form_LQuery")]
        public ActionResult<ResultClass<string>> Fina_form_LQuery(string User, string Type)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = "";
                switch (Type)
                {
                    //TODO 要加入判定應該審核完成
                    //inner join AuditFlow_M AM on AM.FM_Source_ID = PM.PM_ID and AM.FM_Step='9'
                    //inner join AuditFlow_M AM on AM.FM_Source_ID = VD.VD_ID and AM.FM_Step='9'
                    case "PO":
                        T_SQL += @"select 'PO' as FormType,PM.PM_ID as FormID,PD.PD_Pro_name as FormCaption,PD.PD_Amt as FormMoney 
                                   from Procurement_M PM
                                   inner join Procurement_D PD on PM.PM_ID = PD.PM_ID
                                   where VP_ID is null and PM_U_num=@User";
                        break;
                    case "PS":
                        T_SQL += @"select VM.VP_type as FormType,VM.VP_ID as FormID,VD.VD_Fee_Summary as FormCaption,VD.VD_Fee as FormMoney    
                                   from InvPrepay_M VM
                                   inner join InvPrepay_D VD on VM.VP_ID=VD.VP_ID and not exists (select 1 from InvPrepay_D where Form_ID=VM.VP_ID)
                                   where VM.VP_type = 'PP' and VP_U_num =@User";
                        break;
                    default:
                        break;
                }
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
                var parameters_id = new List<SqlParameter>();
                var T_SQL_ID = @"exec GetFormID @formtype,@tablename";
                var str_type = "PA";
                if (model.VP_type != null && model.VP_type.Length > 0 )
                {
                    if (model.VP_type.Contains("PP"))
                        str_type = "PP";
                    if (model.VP_type.Contains("PS"))
                        str_type = "PS";
                }
                parameters_id.Add(new SqlParameter("@formtype", str_type));
                parameters_id.Add(new SqlParameter("@tablename", "InvPrepay_M"));
                var result_id = _adoData.ExecuteQuery(T_SQL_ID, parameters_id);
                var VP_ID = result_id.Rows[0]["resultID"].ToString();

                var parameters_m = new List<SqlParameter>();
                var T_SQL_M = @"Insert into InvPrepay_M (VP_ID,VP_type,VP_BC,VP_AppDate,VP_U_num,VP_Total_Money,bank_code,bank_name,
                    branch_name,bank_account,payee_name,VP_cknum,add_date,add_num,add_ip,edit_date,edit_num,edit_ip) 
                    Values (@VP_ID,@VP_type,@VP_BC,FORMAT(GETDATE(),'yyyy/MM/dd'),@VP_U_num,@VP_Total_Money,@bank_code,@bank_name,@branch_name,
                    @bank_account,@payee_name,@VP_cknum,GETDATE(),@add_num,@add_ip,GETDATE(),@edit_num,@edit_ip)";
                parameters_m.Add(new SqlParameter("@VP_ID", VP_ID));
                parameters_m.Add(new SqlParameter("@VP_type", str_type));
                parameters_m.Add(new SqlParameter("@VP_BC", model.VP_BC));
                parameters_m.Add(new SqlParameter("@VP_U_num", model.User));
                parameters_m.Add(new SqlParameter("@VP_Total_Money", model.VP_Total_Money));
                if (!string.IsNullOrEmpty(model.bank_code))
                {
                    parameters_m.Add(new SqlParameter("@bank_code", model.bank_code));
                }
                else
                {
                    parameters_m.Add(new SqlParameter("@bank_code",DBNull.Value));
                }
                if (!string.IsNullOrEmpty(model.bank_name))
                {
                    parameters_m.Add(new SqlParameter("@bank_name", model.bank_name));
                }
                else
                {
                    parameters_m.Add(new SqlParameter("@bank_name", DBNull.Value));
                }
                if (!string.IsNullOrEmpty(model.branch_name))
                {
                    parameters_m.Add(new SqlParameter("@branch_name", model.branch_name));
                }
                else
                {
                    parameters_m.Add(new SqlParameter("@branch_name", DBNull.Value));
                }
                if (!string.IsNullOrEmpty(model.bank_account))
                {
                    parameters_m.Add(new SqlParameter("@bank_account", model.bank_account));
                }
                else
                {
                    parameters_m.Add(new SqlParameter("@bank_account", DBNull.Value));
                }
                if (!string.IsNullOrEmpty(model.payee_name))
                {
                    parameters_m.Add(new SqlParameter("@payee_name", model.payee_name));
                }
                else
                {
                    parameters_m.Add(new SqlParameter("@payee_name", DBNull.Value));
                }
                parameters_m.Add(new SqlParameter("@VP_cknum", FuncHandler.GetCheckNum()));
                parameters_m.Add(new SqlParameter("@add_num", model.User));
                parameters_m.Add(new SqlParameter("@add_ip", clientIp));
                parameters_m.Add(new SqlParameter("@edit_num", model.User));
                parameters_m.Add(new SqlParameter("@edit_ip", model.User));
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
                        var parameters_d = new List<SqlParameter>();
                        var T_SQL_D = @"Insert into InvPrepay_D (VP_ID,VP_type,Form_ID,VD_Fee_Summary,VD_Fee,add_date,add_num,add_ip,
                            edit_date,edit_num,edit_ip,Change_reason,Change_num,Change_date,Change_ip) 
                            Values (@VP_ID,@VP_type,@Form_ID,@VD_Fee_Summary,@VD_Fee,GETDATE(),@add_num,@add_ip,
                            GETDATE(),@edit_num,@edit_ip,@Change_reason,@Change_num,@Change_date,@Change_ip)";
                        parameters_d.Add(new SqlParameter("@VP_ID", VP_ID));
                        parameters_d.Add(new SqlParameter("@VP_type", str_type));
                        if (!string.IsNullOrEmpty(item.FormID))
                        {
                            parameters_d.Add(new SqlParameter("@Form_ID", item.FormID));
                        }
                        else
                        {
                            parameters_d.Add(new SqlParameter("@Form_ID", DBNull.Value));
                        }
                        parameters_d.Add(new SqlParameter("@VD_Fee_Summary", item.FormCaption));
                        parameters_d.Add(new SqlParameter("@VD_Fee", item.FormMoney));
                        parameters_d.Add(new SqlParameter("@add_num", model.User));
                        parameters_d.Add(new SqlParameter("@add_ip", clientIp));
                        parameters_d.Add(new SqlParameter("@edit_num", model.User));
                        parameters_d.Add(new SqlParameter("@edit_ip", model.User));
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

                        //勾稽請採購單
                        var distinctFormIDs = model.Ins_List.Where(item => !string.IsNullOrEmpty(item.FormID))
                                        .GroupBy(item => item.FormID).Select(Group => Group.Key).ToList();
                        if(distinctFormIDs !=null && distinctFormIDs.Count() > 0)
                        {
                            if(str_type =="PA" || str_type == "PP")
                            {
                                foreach (var item in distinctFormIDs)
                                {
                                    var parameters_p = new List<SqlParameter>();
                                    var T_SQL_P = @"Update Procurement_M set VP_ID=@VP_ID where PM_ID=@Form_ID";
                                    parameters_p.Add(new SqlParameter("@VP_ID", VP_ID));
                                    parameters_p.Add(new SqlParameter("@Form_ID", item));
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
                var parameters = new List<SqlParameter>();
                var T_SQL = @"select VP.VP_ID,AM.FM_Step,LI.item_D_name,LI.item_D_name AS FM_Step_SignType,VP.VP_cknum,VP.VP_Cancel
                    ,(SELECT COUNT(*) FROM ASP_UpLoad WHERE cknum = VP.VP_cknum and del_tag='0') AS VP_cknum_count 
                    from InvPrepay_M VP
                    INNER JOIN AuditFlow_M AM ON AM.FM_Source_ID = VP.VP_ID and AM.AF_ID = VP.VP_type
                    LEFT JOIN Item_list LI ON LI.item_D_code = AM.FM_Step_SignType AND LI.item_M_code = 'Flow_sign_type'
                    where VP.add_num = @User ";
                parameters.Add(new SqlParameter("@User", User));
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
                var parameters = new List<SqlParameter>();
                var T_SQL = @"update InvPrepay_M set VP_Cancel='Y',cancel_date=GETDATE(),cancel_num=@cancel_num,candel_ip=@candel_ip where VP_ID=@VP_ID";
                parameters.Add(new SqlParameter("@cancel_num", User));
                parameters.Add(new SqlParameter("@candel_ip", clientIp));
                parameters.Add(new SqlParameter("@VP_ID", VP_ID));
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
    }
}
