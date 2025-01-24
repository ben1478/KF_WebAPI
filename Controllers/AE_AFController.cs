using KF_WebAPI.BaseClass;
using KF_WebAPI.BaseClass.AE;
using KF_WebAPI.FunctionHandler;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Newtonsoft.Json;
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
        /// 審核主檔列表查詢 AuditFlowM_LQurey
        /// </summary>
        [HttpPost("AuditFlowM_LQurey")]
        public ActionResult<ResultClass<string>> AuditFlowM_LQurey(AuditFlow_M_Req model)
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
        /// 單筆審核狀態查詢 RevFlow_SQuery
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
                var T_SQL = @"select AM.FM_Step,AD.FD_Step,AD.FD_Step_title,AD.FD_Step_SignType from AuditFlow_M AM
                    inner join AuditFlow_D AD on AM.AF_ID=AD.AF_ID and AM.FM_Source_ID = AD.FM_Source_ID where AM.FM_Source_ID =@FormID";
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
                var T_SQL = @"SELECT AM.AF_ID,AM.FM_Source_ID,AM.FM_Step,AD.FD_Step,AD.FD_Step_title,LI.item_D_name AS FM_Step_SignType,
                    COALESCE(PM.PM_cknum, IM.VP_cknum) AS cknum
                    ,(SELECT COUNT(*) FROM ASP_UpLoad WHERE COALESCE(PM.PM_cknum, IM.VP_cknum) = cknum and del_tag='0') as cknum_count
                    FROM AuditFlow_M AM
                    INNER JOIN AuditFlow_D AD ON AM.AF_ID = AD.AF_ID AND AM.FM_Source_ID = AD.FM_Source_ID
                    LEFT JOIN Item_list LI ON LI.item_D_code = AM.FM_Step_SignType AND LI.item_M_code = 'Flow_sign_type'
                    LEFT JOIN Procurement_M PM ON AM.FM_Source_ID = PM.PM_ID 
                    LEFT JOIN InvPrepay_M IM ON AM.FM_Source_ID = IM.VP_ID 
                    WHERE AM.add_date BETWEEN @RF_Date_S AND @RF_Date_E";
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
        /// 審核異動 AuditFlow_Upd
        /// </summary>
        //[HttpPost("AuditFlow_Upd")]
        //public ActionResult<ResultClass<string>> AuditFlow_Upd(Flow_Req model)
        //{
        //    ResultClass<string> resultClass = new ResultClass<string>();

        //    var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

        //    try
        //    {
        //        var step = model.FM_Step_Now;
        //        string columnSign = $"FD_Step{step}_SignType";
        //        string columnDate = $"FD_Step{step}_date";
        //        string columnNote = $"FD_Step{step}_note";

        //        ADOData _adoData = new ADOData();
        //        #region SQL
        //        var parameters = new List<SqlParameter>();
        //        var T_SQL = $@"update AuditFlow_D set FM_Step_SignType=@FM_Step_SignType,FM_Step_Now=@FM_Step_Now,{columnSign}=@columnSign
        //            ,{columnDate}=GETDATE(),{columnNote}=@columnNote,edit_date=GETDATE(),edit_num=@edit_num,edit_ip=@edit_ip 
        //            where FM_ID=@FM_ID and FD_Source_ID=@FD_Source_ID";

        //        if(model.FD_step_sign == "FSIGN002") //同意
        //        {
        //            if (model.FM_Step != model.FM_Step_Now)
        //            {
        //                parameters.Add(new SqlParameter("@FM_Step_Now", Convert.ToInt32(model.FM_Step_Now) + 1));
        //                parameters.Add(new SqlParameter("@FM_Step_SignType", "FSIGN001"));
        //            }
        //            else
        //            {
        //                parameters.Add(new SqlParameter("@FM_Step_Now", Convert.ToInt32(model.FM_Step_Now)));
        //                parameters.Add(new SqlParameter("@FM_Step_SignType", "FSIGN002"));
        //            }
        //        }
        //        else if (model.FD_step_sign == "FSIGN003") //不同意
        //        {
        //            parameters.Add(new SqlParameter("@FM_Step_Now", Convert.ToInt32(model.FM_Step_Now)));
        //            parameters.Add(new SqlParameter("@FM_Step_SignType", "FSIGN003"));
        //        }
        //        else
        //        {
        //            parameters.Add(new SqlParameter("@FM_Step_Now", Convert.ToInt32(model.FM_Step_Now)));
        //            parameters.Add(new SqlParameter("@FM_Step_SignType", "FSIGN001"));
        //        }
        //        parameters.Add(new SqlParameter("@columnSign", model.FD_step_sign));


        //        if (!string.IsNullOrEmpty(model.FD_step_note))
        //        {
        //            parameters.Add(new SqlParameter("@columnNote", model.FD_step_note));
        //        }
        //        else
        //        {
        //            parameters.Add(new SqlParameter("@columnNote", DBNull.Value));
        //        }
        //        parameters.Add(new SqlParameter("@edit_num", model.PM_U_num));
        //        parameters.Add(new SqlParameter("@edit_ip", clientIp));
        //        parameters.Add(new SqlParameter("@FM_ID", model.FM_ID));
        //        parameters.Add(new SqlParameter("@FD_Source_ID", model.FD_Source_ID));
        //        #endregion
        //        int result = _adoData.ExecuteNonQuery(T_SQL, parameters);
        //        if (result == 0)
        //        {
        //            resultClass.ResultCode = "400";
        //            resultClass.ResultMsg = "修改失敗";
        //            return BadRequest(resultClass);
        //        }
        //        else
        //        {
        //            if (model.FD_step_sign == "FSIGN002" && model.FM_Step != model.FM_Step_Now) //同意且有下一關才進行訊息通知
        //            {
        //                var parameters_ad = new List<SqlParameter>();
        //                var T_SQL_AD = @"select * from AuditFlow_D where FM_ID=@FM_ID and FD_Source_ID=@FD_Source_ID";
        //                parameters_ad.Add(new SqlParameter("@FM_ID", model.FM_ID));
        //                parameters_ad.Add(new SqlParameter("@FD_Source_ID", model.FD_Source_ID));

        //                var dtResult_ad=_adoData.ExecuteQuery(T_SQL_AD, parameters_ad);
        //                var fmStepNowValue = dtResult_ad.Rows[0]["FM_Step_Now"];
        //                string columnNum = $"FD_Step{fmStepNowValue}_num";
        //                var User_Num = dtResult_ad.Rows[0][columnNum].ToString();

        //                switch (model.Msg_kind)
        //                {
        //                    case "MSGK0005":
        //                        FuncHandler.MsgIns("MSGK0005", model.PM_U_num, User_Num, "採購單簽核通知,請前往處理!!", clientIp);
        //                        break;
        //                    default:
        //                        break;
        //                }
        //            }

        //            resultClass.ResultCode = "000";
        //            resultClass.ResultMsg = "修改成功";
        //            return Ok(resultClass);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        resultClass.ResultCode = "500";
        //        resultClass.ResultMsg = $" response: {ex.Message}";
        //        return StatusCode(500, resultClass); // 返回 500 錯誤碼
        //    }
        //}



        /// <summary>
        /// 審核明細檔單筆查詢 AuditFlowD_SQuery
        /// </summary>
        [HttpGet("AuditFlowD_SQuery")]
        public ActionResult<ResultClass<string>> AuditFlowD_SQuery(string FM_ID,string FD_Source_ID)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"select FM_ID,FD_Source_ID
                    ,FD_Step1_Desc,FD_Step1_num,(select U_name from User_M where U_num=FD_Step1_num) as U_name_1,FD_Step1_Reason
                    ,FD_Step2_Desc,FD_Step2_num,(select U_name from User_M where U_num=FD_Step2_num) as U_name_2,FD_Step2_Reason
                    ,FD_Step3_Desc,FD_Step3_num,(select U_name from User_M where U_num=FD_Step3_num) as U_name_3,FD_Step3_Reason 
                    ,FD_Step4_Desc,FD_Step4_num,(select U_name from User_M where U_num=FD_Step4_num) as U_name_4,FD_Step4_Reason 
                    ,FD_Step5_Desc,FD_Step5_num,(select U_name from User_M where U_num=FD_Step5_num) as U_name_5,FD_Step5_Reason 
                    ,FD_Step6_Desc,FD_Step6_num,(select U_name from User_M where U_num=FD_Step6_num) as U_name_6,FD_Step6_Reason 
                    ,FD_Step7_Desc,FD_Step7_num,(select U_name from User_M where U_num=FD_Step7_num) as U_name_7,FD_Step7_Reason 
                    ,FD_Step8_Desc,FD_Step8_num,(select U_name from User_M where U_num=FD_Step8_num) as U_name_8,FD_Step8_Reason 
                    ,FD_Step9_Desc,FD_Step9_num,(select U_name from User_M where U_num=FD_Step9_num) as U_name_9,FD_Step9_Reason 
                    from AuditFlow_D 
                    where FM_ID=@FM_ID and FD_Source_ID=@FD_Source_ID";
                parameters.Add(new SqlParameter("@FM_ID", FM_ID));
                parameters.Add(new SqlParameter("@FD_Source_ID", FD_Source_ID));
                #endregion
                var dtResult=_adoData.ExecuteQuery(T_SQL,parameters);
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
        /// 審核明細檔單筆異動 AuditFlowD_Upd
        /// </summary>
        [HttpPost("AuditFlowD_Upd")]
        public ActionResult<ResultClass<string>> AuditFlowD_Upd(AuditFlowDetail_Ins model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"Update AuditFlow_D set FD_Step1_Desc=@FD_Step1_Desc,FD_Step1_num=@FD_Step1_num
                    ,FD_Step2_Desc=@FD_Step2_Desc,FD_Step2_num=@FD_Step2_num
                    ,FD_Step3_Desc=@FD_Step3_Desc,FD_Step3_num=@FD_Step3_num
                    ,FD_Step4_Desc=@FD_Step4_Desc,FD_Step4_num=@FD_Step4_num
                    ,FD_Step5_Desc=@FD_Step5_Desc,FD_Step5_num=@FD_Step5_num
                    ,FD_Step6_Desc=@FD_Step6_Desc,FD_Step6_num=@FD_Step6_num
                    ,FD_Step7_Desc=@FD_Step7_Desc,FD_Step7_num=@FD_Step7_num
                    ,FD_Step8_Desc=@FD_Step8_Desc,FD_Step8_num=@FD_Step8_num
                    ,FD_Step9_Desc=@FD_Step9_Desc,FD_Step9_num=@FD_Step9_num
                    ,edit_date=GETDATE(),edit_num=@edit_num,edit_ip=@edit_ip
                    where FM_ID=@FM_ID and FD_Source_ID=@FD_Source_ID";
                if(!string.IsNullOrEmpty(model.FD_Step1_Desc) && !string.IsNullOrEmpty(model.FD_Step1_num))
                {
                    parameters.Add(new SqlParameter("@FD_Step1_Desc", model.FD_Step1_Desc));
                    parameters.Add(new SqlParameter("@FD_Step1_num", model.FD_Step1_num));
                }
                else
                {
                    parameters.Add(new SqlParameter("@FD_Step1_Desc", DBNull.Value));
                    parameters.Add(new SqlParameter("@FD_Step1_num", DBNull.Value));
                }
                if (!string.IsNullOrEmpty(model.FD_Step2_Desc) && !string.IsNullOrEmpty(model.FD_Step2_num))
                {
                    parameters.Add(new SqlParameter("@FD_Step2_Desc", model.FD_Step2_Desc));
                    parameters.Add(new SqlParameter("@FD_Step2_num", model.FD_Step2_num));
                }
                else
                {
                    parameters.Add(new SqlParameter("@FD_Step2_Desc", DBNull.Value));
                    parameters.Add(new SqlParameter("@FD_Step2_num", DBNull.Value));
                }
                if (!string.IsNullOrEmpty(model.FD_Step3_Desc) && !string.IsNullOrEmpty(model.FD_Step3_num))
                {
                    parameters.Add(new SqlParameter("@FD_Step3_Desc", model.FD_Step3_Desc));
                    parameters.Add(new SqlParameter("@FD_Step3_num", model.FD_Step3_num));
                }
                else
                {
                    parameters.Add(new SqlParameter("@FD_Step3_Desc", DBNull.Value));
                    parameters.Add(new SqlParameter("@FD_Step3_num", DBNull.Value));
                }
                if (!string.IsNullOrEmpty(model.FD_Step4_Desc) && !string.IsNullOrEmpty(model.FD_Step4_num))
                {
                    parameters.Add(new SqlParameter("@FD_Step4_Desc", model.FD_Step4_Desc));
                    parameters.Add(new SqlParameter("@FD_Step4_num", model.FD_Step4_num));
                }
                else
                {
                    parameters.Add(new SqlParameter("@FD_Step4_Desc", DBNull.Value));
                    parameters.Add(new SqlParameter("@FD_Step4_num", DBNull.Value));
                }
                if (!string.IsNullOrEmpty(model.FD_Step5_Desc) && !string.IsNullOrEmpty(model.FD_Step5_num))
                {
                    parameters.Add(new SqlParameter("@FD_Step5_Desc", model.FD_Step5_Desc));
                    parameters.Add(new SqlParameter("@FD_Step5_num", model.FD_Step5_num));
                }
                else
                {
                    parameters.Add(new SqlParameter("@FD_Step5_Desc", DBNull.Value));
                    parameters.Add(new SqlParameter("@FD_Step5_num", DBNull.Value));
                }
                if (!string.IsNullOrEmpty(model.FD_Step6_Desc) && !string.IsNullOrEmpty(model.FD_Step6_num))
                {
                    parameters.Add(new SqlParameter("@FD_Step6_Desc", model.FD_Step6_Desc));
                    parameters.Add(new SqlParameter("@FD_Step6_num", model.FD_Step6_num));
                }
                else
                {
                    parameters.Add(new SqlParameter("@FD_Step6_Desc", DBNull.Value));
                    parameters.Add(new SqlParameter("@FD_Step6_num", DBNull.Value));
                }
                if (!string.IsNullOrEmpty(model.FD_Step7_Desc) && !string.IsNullOrEmpty(model.FD_Step7_num))
                {
                    parameters.Add(new SqlParameter("@FD_Step7_Desc", model.FD_Step7_Desc));
                    parameters.Add(new SqlParameter("@FD_Step7_num", model.FD_Step7_num));
                }
                else
                {
                    parameters.Add(new SqlParameter("@FD_Step7_Desc", DBNull.Value));
                    parameters.Add(new SqlParameter("@FD_Step7_num", DBNull.Value));
                }
                if (!string.IsNullOrEmpty(model.FD_Step8_Desc) && !string.IsNullOrEmpty(model.FD_Step8_num))
                {
                    parameters.Add(new SqlParameter("@FD_Step8_Desc", model.FD_Step8_Desc));
                    parameters.Add(new SqlParameter("@FD_Step8_num", model.FD_Step8_num));
                }
                else
                {
                    parameters.Add(new SqlParameter("@FD_Step8_Desc", DBNull.Value));
                    parameters.Add(new SqlParameter("@FD_Step8_num", DBNull.Value));
                }
                if (!string.IsNullOrEmpty(model.FD_Step9_Desc) && !string.IsNullOrEmpty(model.FD_Step9_num))
                {
                    parameters.Add(new SqlParameter("@FD_Step9_Desc", model.FD_Step9_Desc));
                    parameters.Add(new SqlParameter("@FD_Step9_num", model.FD_Step9_num));
                }
                else
                {
                    parameters.Add(new SqlParameter("@FD_Step9_Desc", DBNull.Value));
                    parameters.Add(new SqlParameter("@FD_Step9_num", DBNull.Value));
                }
                parameters.Add(new SqlParameter("@edit_num", model.edit_num));
                parameters.Add(new SqlParameter("@edit_ip", clientIp));
                parameters.Add(new SqlParameter("@FM_ID", model.FM_ID));
                parameters.Add(new SqlParameter("@FD_Source_ID", model.FD_Source_ID));
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
        /// 審核人員變更原因 AuditFlowD_UpdReason
        /// </summary>
        [HttpPost("AuditFlowD_UpdReason")]
        public ActionResult<ResultClass<string>> AuditFlowD_UpdReason(AuditFlowReason model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();
            string number = Regex.Match(model.Type, @"\d+").Value;

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = $@"Update AuditFlow_D set FD_Step{number}_Reason=@Reason,edit_date=GETDATE(),edit_num=@edit_num,edit_ip=@edit_ip 
                    where FM_ID=@FM_ID and FD_Source_ID=@FD_Source_ID";
                parameters.Add(new SqlParameter("@Reason", model.Reason));
                parameters.Add(new SqlParameter("@edit_num", model.User));
                parameters.Add(new SqlParameter("@edit_ip",clientIp));
                parameters.Add(new SqlParameter("@FM_ID",model.FM_ID));
                parameters.Add(new SqlParameter("@FD_Source_ID", model.FD_Source_ID));
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
    }
}
