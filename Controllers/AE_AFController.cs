using KF_WebAPI.BaseClass;
using KF_WebAPI.BaseClass.AE;
using KF_WebAPI.FunctionHandler;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Newtonsoft.Json;

namespace KF_WebAPI.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class AE_AFController : ControllerBase
    {
        /// <summary>
        /// 審核列表查詢
        /// </summary>
        [HttpPost("AuditFlow_LQuery")]
        public ActionResult<ResultClass<string>> AuditFlow_LQuery(AuditFlow_Req model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"select * from AuditFlow_D where FD_Source_ID=@FD_Source_ID and FM_ID=@FM_ID";
                parameters.Add(new SqlParameter("@FD_Source_ID", model.FD_Source_ID));
                parameters.Add(new SqlParameter("@FM_ID", model.FM_ID));
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
        /// 審核異動
        /// </summary>
        [HttpPost("AuditFlow_Upd")]
        public ActionResult<ResultClass<string>> AuditFlow_Upd(Flow_Req model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

            try
            {
                var step = model.FM_Step_Now;
                string columnSign = $"FD_Step{step}_SignType";
                string columnDate = $"FD_Step{step}_date";
                string columnNote = $"FD_Step{step}_note";

                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = $@"update AuditFlow_D set FM_Step_SignType=@FM_Step_SignType,FM_Step_Now=@FM_Step_Now,{columnSign}=@columnSign
                    ,{columnDate}=GETDATE(),{columnNote}=@columnNote,edit_date=GETDATE(),edit_num=@edit_num,edit_ip=@edit_ip 
                    where FM_ID=@FM_ID and FD_Source_ID=@FD_Source_ID";

                if(model.FD_step_sign == "FSIGN002") //同意
                {
                    if (model.FM_Step != model.FM_Step_Now)
                    {
                        parameters.Add(new SqlParameter("@FM_Step_Now", Convert.ToInt32(model.FM_Step_Now) + 1));
                        parameters.Add(new SqlParameter("@FM_Step_SignType", "FSIGN001"));
                    }
                    else
                    {
                        parameters.Add(new SqlParameter("@FM_Step_Now", Convert.ToInt32(model.FM_Step_Now)));
                        parameters.Add(new SqlParameter("@FM_Step_SignType", "FSIGN002"));
                    }
                }
                else if (model.FD_step_sign == "FSIGN003") //不同意
                {
                    parameters.Add(new SqlParameter("@FM_Step_Now", Convert.ToInt32(model.FM_Step_Now)));
                    parameters.Add(new SqlParameter("@FM_Step_SignType", "FSIGN003"));
                }
                else
                {
                    parameters.Add(new SqlParameter("@FM_Step_Now", Convert.ToInt32(model.FM_Step_Now)));
                    parameters.Add(new SqlParameter("@FM_Step_SignType", "FSIGN001"));
                }
                parameters.Add(new SqlParameter("@columnSign", model.FD_step_sign));


                if (!string.IsNullOrEmpty(model.FD_step_note))
                {
                    parameters.Add(new SqlParameter("@columnNote", model.FD_step_note));
                }
                else
                {
                    parameters.Add(new SqlParameter("@columnNote", DBNull.Value));
                }
                parameters.Add(new SqlParameter("@edit_num", model.PM_U_num));
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
    }
}
