using KF_WebAPI.BaseClass;
using KF_WebAPI.BaseClass.AE;
using KF_WebAPI.FunctionHandler;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Newtonsoft.Json;
using System.Data;

namespace KF_WebAPI.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class AE_MAFController : ControllerBase
    {
        /// <summary>
        /// 廠商列表查詢 Manufacturer_LQuery
        /// </summary>
        [HttpGet("Manufacturer_LQuery")]
        public ActionResult<ResultClass<string>> Manufacturer_LQuery()
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = "select * from Manufacturer";
                #endregion
                DataTable dtResult = _adoData.ExecuteSQuery(T_SQL);

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
        /// 廠商資料單筆查詢 Manufacturer_SQuery
        /// </summary>
        [HttpGet("Manufacturer_SQuery")]
        public ActionResult<ResultClass<string>> Manufacturer_SQuery(string MF_ID)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"select * from Manufacturer where MF_ID=@MF_ID";
                parameters.Add(new SqlParameter("@MF_ID", MF_ID));
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
        /// 新增廠商資料 Manufacturer_Ins
        /// </summary>
        [HttpPost("Manufacturer_Ins")]
        public ActionResult<ResultClass<string>> Manufacturer_Ins(Manufacturer_Ins model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();
            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"Insert into Manufacturer(MF_ID,MF_cknum,Company_name,Company_number,Company_addr,Company_busin,Company_tel,Company_fax,Invoice_Iss,Overseas,add_date,add_num,add_ip) 
                    Values (@MF_ID,@MF_cknum,@Company_name,@Company_number,@Company_addr,@Company_busin,@Company_tel,@Company_fax,@Invoice_Iss,@Overseas,GETDATE(),@add_num,@add_ip)";
                parameters.Add(new SqlParameter("@MF_ID", model.MF_ID));
                parameters.Add(new SqlParameter("@MF_cknum", FuncHandler.GetCheckNum()));
                parameters.Add(new SqlParameter("@Company_name", model.Company_name));
                parameters.Add(new SqlParameter("@Company_number", model.Company_number));
                parameters.Add(new SqlParameter("@Company_addr", model.Company_addr));
                parameters.Add(new SqlParameter("@Company_busin", model.Company_busin));
                parameters.Add(new SqlParameter("@Company_tel", model.Company_tel));
                parameters.Add(new SqlParameter("@Company_fax", model.Company_fax));
                parameters.Add(new SqlParameter("@Invoice_Iss", model.Invoice_Iss));
                parameters.Add(new SqlParameter("@Overseas", model.Overseas));
                parameters.Add(new SqlParameter("@add_num", model.add_num));
                parameters.Add(new SqlParameter("@add_ip", clientIp));
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
                return StatusCode(500, resultClass);
            }
        }

        /// <summary>
        /// 修改廠商資料 Manufacturer_Upd
        /// </summary>
        [HttpPost("Manufacturer_Upd")]
        public ActionResult<ResultClass<string>> Manufacturer_Upd(Manufacturer_Ins model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();
            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"Update Manufacturer set Company_name=@Company_name,Company_number=@Company_number,Company_addr=@Company_addr,
                    Company_busin=@Company_busin,Company_tel=@Company_tel,Company_fax=@Company_fax, 
                    Invoice_Iss=@Invoice_Iss,Overseas=@Overseas,edit_date=GETDATE(),edit_num=@edit_num,edit_ip=@edit_ip where MF_ID=@MF_ID";
                parameters.Add(new SqlParameter("@Company_name", model.Company_name ?? ""));
                parameters.Add(new SqlParameter("@Company_number", model.Company_number ?? ""));
                parameters.Add(new SqlParameter("@Company_addr", model.Company_addr ?? ""));
                parameters.Add(new SqlParameter("@Company_busin", model.Company_busin ?? ""));
                parameters.Add(new SqlParameter("@Company_tel", model.Company_tel ?? ""));
                parameters.Add(new SqlParameter("@Company_fax", model.Company_fax ?? ""));
                parameters.Add(new SqlParameter("@Invoice_Iss", model.Invoice_Iss ?? ""));
                parameters.Add(new SqlParameter("@Overseas", model.Overseas ?? ""));
                parameters.Add(new SqlParameter("@edit_num", model.edit_num));
                parameters.Add(new SqlParameter("@edit_ip", clientIp));
                parameters.Add(new SqlParameter("@MF_ID", model.MF_ID));
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
                return StatusCode(500, resultClass);
            }
        }
    }
}
