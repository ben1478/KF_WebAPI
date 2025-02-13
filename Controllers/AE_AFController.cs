using KF_WebAPI.BaseClass;
using KF_WebAPI.BaseClass.AE;
using KF_WebAPI.FunctionHandler;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Newtonsoft.Json;
using OfficeOpenXml.ExternalReferences;
using System;
using System.Data;
using System.Text.RegularExpressions;

namespace KF_WebAPI.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class AE_AFController : ControllerBase
    {
        /// <summary>
        /// 審核流程設定列表查詢 AuditFlow_LQurey
        /// </summary>
        [HttpPost("AuditFlow_LQurey")]
        public ActionResult<ResultClass<string>> AuditFlow_LQurey()
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"select AF_ID,AF_Name,AF_Caption,AF_Step,AF_Step_Caption from AuditFlow ";
                #endregion
                DataTable dtResult = _adoData.ExecuteSQuery(T_SQL);
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
        /// 審核流程設定單筆查詢 AuditFlow_SQuery
        /// </summary>
        /// <param name="FM_ID">PO001</param>
        [HttpGet("AuditFlow_SQuery")]
        public ActionResult<ResultClass<string>> AuditFlow_SQuery(string AF_ID)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"select * from AuditFlow where AF_ID=@AF_ID";
                parameters.Add(new SqlParameter("@AF_ID", AF_ID));
                #endregion
                var dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
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
        /// 審核流程設定新增 AuditFlow_Ins
        /// </summary>
        [HttpPost("AuditFlow_Ins")]
        public ActionResult<ResultClass<string>> AuditFlow_Ins(AuditFlow_Ins model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();
            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"Insert into AuditFlow(AF_ID,AF_Name,AF_Caption,AF_Step,AF_Step_Caption,AF_Step_Person,add_date,add_num,add_ip,edit_date,edit_num,edit_ip) 
                    values (@AF_ID,@AF_Name,@AF_Caption,@AF_Step,@AF_Step_Caption,@AF_Step_Person,GETDATE(),@add_num,@add_ip,GETDATE(),@edit_num,@edit_ip)";
                parameters.Add(new SqlParameter("@AF_ID", model.AF_ID));
                parameters.Add(new SqlParameter("@AF_Name", model.AF_Name));
                parameters.Add(new SqlParameter("@AF_Caption", model.AF_Caption));
                parameters.Add(new SqlParameter("@AF_Step", model.AF_Step));
                parameters.Add(new SqlParameter("@AF_Step_Caption", model.AF_Step_Caption));
                parameters.Add(new SqlParameter("@AF_Step_Person", model.AF_Step_Person));
                parameters.Add(new SqlParameter("@add_num", model.add_num));
                parameters.Add(new SqlParameter("@add_ip", clientIp));
                parameters.Add(new SqlParameter("@edit_num", model.edit_num));
                parameters.Add(new SqlParameter("@edit_ip", clientIp));
                #endregion
                int result = _adoData.ExecuteNonQuery(T_SQL, parameters);
                if (result == 0)
                {
                    resultClass.ResultCode = "400";
                    resultClass.ResultMsg = "新增失敗";
                    return BadRequest(resultClass);
                }
                else
                {
                    resultClass.ResultCode = "000";
                    resultClass.ResultMsg = "新增成功";
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
        /// 審核流程設定修改 AuditFlow_Upd
        /// </summary>
        [HttpPost("AuditFlow_Upd")]
        public ActionResult<ResultClass<string>> AuditFlow_Upd(AuditFlow_Ins model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();
            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"Update AuditFlow set AF_Name=@AF_Name,AF_Caption=@AF_Caption,AF_Step=@AF_Step,AF_Step_Caption=@AF_Step_Caption,AF_Step_Person=@AF_Step_Person
                    ,edit_date=GETDATE(),edit_num=@edit_num where AF_ID=@AF_ID";
                parameters.Add(new SqlParameter("@AF_ID", model.AF_ID));
                parameters.Add(new SqlParameter("@AF_Name", model.AF_Name));
                parameters.Add(new SqlParameter("@AF_Caption", model.AF_Caption));
                parameters.Add(new SqlParameter("@AF_Step", model.AF_Step));
                parameters.Add(new SqlParameter("@AF_Step_Caption", model.AF_Step_Caption));
                parameters.Add(new SqlParameter("@AF_Step_Person", model.AF_Step_Person));
                parameters.Add(new SqlParameter("@edit_num", model.edit_num));
                parameters.Add(new SqlParameter("@edit_ip", clientIp));
                #endregion
                int result = _adoData.ExecuteNonQuery(T_SQL, parameters);
                if (result == 0)
                {
                    resultClass.ResultCode = "400";
                    resultClass.ResultMsg = "修改失敗";
                    return BadRequest(resultClass);
                }
                else
                {
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
        /// 審核主檔列表查詢 AuditFlow_M_LQurey
        /// </summary>
        [HttpPost("AuditFlow_M_LQurey")]
        public ActionResult<ResultClass<string>> AuditFlow_M_LQurey(AuditFlow_M_Req model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"select AM.AF_ID,FM_Source_ID,AM.FM_Step ,LI.item_D_name as FM_Step_SignType 
                    from AuditFlow_M AM
                    left join Item_list LI on LI.item_D_code=AM.FM_Step_SignType and LI.item_M_code='Flow_sign_type'
                    where 1=1";
                if (!string.IsNullOrEmpty(model.AF_ID))
                {
                    T_SQL += " and AF_ID=@AF_ID";
                    parameters.Add(new SqlParameter("@AF_ID", model.AF_ID));
                }
                if (!string.IsNullOrEmpty(model.FM_Source_ID))
                {
                    T_SQL += " and FM_Source_ID=@FM_Source_ID";
                    parameters.Add(new SqlParameter("@FM_Source_ID", model.FM_Source_ID));
                }
                #endregion
                DataTable dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
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
        /// 單筆審核流程狀態查詢 RevFlow_Query
        /// </summary>
        [HttpGet("RevFlow_Query")]
        public ActionResult<ResultClass<string>> RevFlow_Query(string FormID)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"select distinct AM.FM_Step,AD.FD_Step,AD.FD_Step_title,AD.FD_Step_SignType from AuditFlow_M AM
                    inner join AuditFlow_D AD on AM.AF_ID=AD.AF_ID and AM.FM_Source_ID = AD.FM_Source_ID 
                    where FD_Sign_Countersign = 'S' and AM.FM_Source_ID =@FormID";
                parameters.Add(new SqlParameter("@FormID", FormID));
                #endregion
                var dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
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
        /// 審核案件列表查詢 RevFlow_LQuery
        /// </summary>
        [HttpPost("RevFlow_LQuery")]
        public ActionResult<ResultClass<string>> RevFlow_LQuery(RevFlow_Req model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"SELECT AM.AF_ID,AM.FM_Source_ID,AM.FM_Step,AD.FD_Step,AD.FD_Step_title,AD.FD_Step_num,LI.item_D_name AS FM_Step_SignType,
                    AD.FD_Step_SignType,COALESCE(PM.PM_cknum, IM.VP_cknum) AS cknum
                    ,(SELECT COUNT(*) FROM ASP_UpLoad WHERE COALESCE(PM.PM_cknum, IM.VP_cknum) = cknum and del_tag='0') as cknum_count
                    FROM AuditFlow_M AM
                    INNER JOIN AuditFlow_D AD ON AM.AF_ID = AD.AF_ID AND AM.FM_Source_ID = AD.FM_Source_ID
                    LEFT JOIN Item_list LI ON LI.item_D_code = AM.FM_Step_SignType AND LI.item_M_code = 'Flow_sign_type'
                    LEFT JOIN Procurement_M PM ON AM.FM_Source_ID = PM.PM_ID 
                    LEFT JOIN InvPrepay_M IM ON AM.FM_Source_ID = IM.VP_ID 
                    WHERE (PM.PM_Cancel='N' OR IM.VP_Cancel='N') and AM.add_date BETWEEN @RF_Date_S AND @RF_Date_E
                    order by AM.add_date desc";
                if (!string.IsNullOrEmpty(model.U_BC))
                {
                    T_SQL += " and AM.FM_BC = @U_BC";
                    parameters.Add(new SqlParameter("@U_BC", model.U_BC));
                }
                parameters.Add(new SqlParameter("@RF_Date_S", FuncHandler.ConvertROCToGregorian(model.RF_Date_S)));
                parameters.Add(new SqlParameter("@RF_Date_E", FuncHandler.ConvertROCToGregorian(model.RF_Date_E)));
                #endregion
                var dtResult= _adoData.ExecuteQuery(T_SQL, parameters);
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
        /// 審核案件單筆查詢 RevFlow_SQuery
        /// </summary>
        [HttpGet("RevFlow_SQuery")]
        public ActionResult<ResultClass<string>> RevFlow_SQuery(string FormID)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"select FM.AF_ID,FM.FM_Step,FM.FM_Source_ID,FD.FD_Sign_Countersign,FD.FD_Step_num,
                    UM.U_name,ISNULL(FD.FD_Step_desc,'') as FD_Step_desc  
                    from AuditFlow_M FM
                    inner join AuditFlow_D FD on FM.FM_Source_ID=FD.FM_Source_ID and FM.FM_Step=FD.FD_Step
                    left join User_M UM on UM.U_num = FD.FD_Step_num
                    where FM.FM_Source_ID=@FormID";
                parameters.Add(new SqlParameter("@FormID", FormID));
                #endregion
                var dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
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

        /// <summary>
        /// 寫入會簽人員
        /// </summary>
        [HttpPost("Counter_Ins")]
        public ActionResult<ResultClass<string>> Counter_Ins(Counter_Ins model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

            try
            {
                ADOData _adoData = new ADOData();

                #region SQL
                if (model.arr_Unm != null && model.arr_Unm.Length > 0)
                {
                    var T_SQL = @"Insert into AuditFlow_D (AF_ID, FM_Source_ID,FD_Step,FD_Step_SignType,FD_Sign_Countersign,FD_Step_num,add_date,add_num
                        ,add_ip,edit_date,edit_num,edit_ip) 
                        Values (@AF_ID, @FM_Source_ID,@FD_Step,'FSIGN001','C', @FD_Step_num,GETDATE(),@add_num,@add_ip,GETDATE(),@edit_num,@edit_ip)";

                    foreach (var item in model.arr_Unm)
                    {
                        var parameters = new List<SqlParameter>();
                        parameters.Add(new SqlParameter("@AF_ID", model.AF_ID));
                        parameters.Add(new SqlParameter("@FM_Source_ID", model.FM_Source_ID));
                        parameters.Add(new SqlParameter("@FD_Step", model.FM_Step));
                        parameters.Add(new SqlParameter("@FD_Step_num", item));
                        parameters.Add(new SqlParameter("@add_num", model.User));
                        parameters.Add(new SqlParameter("@add_ip", clientIp));
                        parameters.Add(new SqlParameter("@edit_num", model.User));
                        parameters.Add(new SqlParameter("@edit_ip", clientIp));
                        _adoData.ExecuteQuery(T_SQL, parameters);
                    }

                }
                #endregion
                //TODO 發訊息通知會簽人員
                resultClass.ResultCode = "000";
                resultClass.ResultMsg = "新增成功";
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
        /// 審核明細檔異動
        /// </summary>
        [HttpPost("AuditFlow_D_Upd")]
        public ActionResult<ResultClass<string>> AuditFlow_D_Upd(AuditFlow_D_Upd model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"Update AuditFlow_D set FD_Step_desc=@FD_Step_desc,FD_Step_SignType=@FD_Step_SignType,FD_Step_date=GETDATE(),
                    edit_date=GETDATE(),edit_num=@User,edit_ip=@IP
                    where FM_Source_ID=@FM_Source_ID and FD_Step_num=@User and FD_Sign_Countersign=@FD_Sign_Countersign and FD_Step=@FD_Step";
                if(!string.IsNullOrEmpty(model.FD_Step_desc))
                {
                    parameters.Add(new SqlParameter("@FD_Step_desc", model.FD_Step_desc));
                }
                else
                {
                    parameters.Add(new SqlParameter("@FD_Step_desc", DBNull.Value));
                }
                parameters.Add(new SqlParameter("@FD_Step_SignType", model.FD_Step_SignType));
                parameters.Add(new SqlParameter("@FM_Source_ID", model.FM_Source_ID));
                parameters.Add(new SqlParameter("@User", model.User));
                parameters.Add(new SqlParameter("@IP", clientIp));
                parameters.Add(new SqlParameter("@FD_Sign_Countersign", model.FD_Sign_Countersign));
                parameters.Add(new SqlParameter("@FD_Step", model.FM_Step));
                #endregion
                int Result = _adoData.ExecuteNonQuery(T_SQL,parameters);
                if (Result==0)
                {
                    resultClass.ResultCode = "400";
                    resultClass.ResultMsg = "異動失敗";
                    return BadRequest(resultClass);
                }
                else
                {
                    var parameters_sp = new List<SqlParameter>();
                    var T_SQL_SP = @"exec UpdAuditFlowM @Form_ID,@FM_Step,@Signtype,@AF_Back_Reason";
                    parameters_sp.Add(new SqlParameter("@Form_ID", model.FM_Source_ID));
                    parameters_sp.Add(new SqlParameter("@FM_Step", model.FM_Step));
                    parameters_sp.Add(new SqlParameter("@Signtype", model.FD_Step_SignType));
                    if(model.FD_Step_SignType== "FSIGN003")
                    {
                        parameters_sp.Add(new SqlParameter("@AF_Back_Reason", model.FD_Step_desc));
                    }
                    else
                    {
                        parameters_sp.Add(new SqlParameter("@AF_Back_Reason",DBNull.Value));
                    }
                    _adoData.ExecuteQuery(T_SQL_SP, parameters_sp);
                    //TODO 訊息通知
                    resultClass.ResultCode = "000";
                    resultClass.ResultMsg = "異動成功";

                    //TODO 通知API抽單 ??
                    if (model.FD_Step_SignType == "FSIGN003")
                    {
                      

                    }
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
        /// 查詢案件所有的審核人員(含會簽) AuditFlow_D_LQuery
        /// </summary>
        /// <param name="Form_ID">PO20250204002</param>
        [HttpGet("AuditFlow_D_LQuery")]
        public ActionResult<ResultClass<string>> AuditFlow_D_LQuery(string Form_ID)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"select FD_ID,FD_Step,UM.U_name,FD_Step_num,FD_Step_SignType 
                    from AuditFlow_D FD 
                    left join User_M UM on UM.U_num = FD.FD_Step_num
                    where FM_Source_ID=@Form_ID order by FD_Step";
                parameters.Add(new SqlParameter("@Form_ID", Form_ID));
                #endregion
                var dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
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


        /// <summary>
        /// 審核人員變更原因 AuditFlowD_UpdReason
        /// </summary>
        [HttpPost("AuditFlowD_UpdReason")]
        public ActionResult<ResultClass<string>> AuditFlowD_UpdReason(AuditFlowReason model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"Update AuditFlow_D set FD_Step_num=@FD_Step_num,FD_Step_Reason=@Reason,
                    edit_date=GETDATE(),edit_num=@User,edit_ip=@IP where FD_ID=@FD_ID";
                parameters.Add(new SqlParameter("@FD_Step_num", model.FD_Step_num));
                parameters.Add(new SqlParameter("@Reason", model.Reason));
                parameters.Add(new SqlParameter("@User", model.User));
                parameters.Add(new SqlParameter("@IP", clientIp));
                parameters.Add(new SqlParameter("@FD_ID", model.FD_ID));
                #endregion
                int result=_adoData.ExecuteNonQuery(T_SQL, parameters);
                if (result == 0)
                {
                    resultClass.ResultCode = "400";
                    resultClass.ResultMsg = "異動失敗";
                    return BadRequest(resultClass);
                }
                else
                {
                    resultClass.ResultCode = "000";
                    resultClass.ResultMsg = "異動成功";
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
