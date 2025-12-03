using KF_WebAPI.BaseClass;
using KF_WebAPI.BaseClass.AE;
using KF_WebAPI.FunctionHandler;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Newtonsoft.Json;
using System.Data;
using System.Text;

namespace KF_WebAPI.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class AE_MAFController : ControllerBase
    {
        /// <summary>
        /// 廠商列表查詢 Manufacturer_LQuery
        /// </summary>
        [HttpPost("Manufacturer_LQuery")]
        public ActionResult<ResultClass<string>> Manufacturer_LQuery(Manufacturer_Req model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            try
            {
                ADOData _adoData = new ADOData();
                var sqlBuilder = new StringBuilder("SELECT * FROM Manufacturer WHERE 1 = 1");
                var parameters = new List<SqlParameter>();

                if (!string.IsNullOrEmpty(model.MF_ID))
                {
                    sqlBuilder.Append(" AND MF_ID = @MF_ID");
                    parameters.Add(new SqlParameter("@MF_ID", model.MF_ID));
                }
                if (!string.IsNullOrEmpty(model.Company_name))
                {
                    sqlBuilder.Append(" AND Company_name LIKE @Company_name");
                    parameters.Add(new SqlParameter("@Company_name", "%" + model.Company_name + "%"));
                }
                if (!string.IsNullOrEmpty(model.Company_addr))
                {
                    sqlBuilder.Append(" AND Company_addr LIKE @Company_addr");
                    parameters.Add(new SqlParameter("@Company_addr", "%" + model.Company_addr + "%"));
                }
                DataTable dtResult = _adoData.ExecuteQuery(sqlBuilder.ToString(), parameters);

                resultClass.ResultCode = "000";
                resultClass.objResult = JsonConvert.SerializeObject(dtResult);
                return Ok(resultClass);
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                resultClass.ResultMsg = $"Response error: {ex.Message}";
                return StatusCode(500, resultClass);
            }
        }

        /// <summary>
        /// 廠商資料單筆查詢 Manufacturer_SQuery
        /// </summary>
        [HttpGet("Manufacturer_SQuery")]
        public ActionResult<ResultClass<string>> Manufacturer_SQuery(string ID)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"select * from Manufacturer where MF_Number=@ID";
                var parameters = new List<SqlParameter> 
                {
                    new SqlParameter("@ID", ID)
                };
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
               
                var T_SQL = @"Insert into Manufacturer(MF_ID,MF_cknum,Company_name,Company_number,Company_addr,Company_busin,Company_tel,Company_fax,Invoice_Iss,Overseas,add_date,add_num,add_ip) 
                    Values (@MF_ID,@MF_cknum,@Company_name,@Company_number,@Company_addr,@Company_busin,@Company_tel,@Company_fax,@Invoice_Iss,@Overseas,GETDATE(),@add_num,@add_ip)";
                var parameters = new List<SqlParameter>
                {
                    new SqlParameter("@MF_ID", string.IsNullOrEmpty(model.MF_ID) ? DBNull.Value : model.MF_ID),
                    new SqlParameter("@MF_cknum", FuncHandler.GetCheckNum()),
                    new SqlParameter("@Company_name", string.IsNullOrEmpty(model.Company_name) ? DBNull.Value : model.Company_name),
                    new SqlParameter("@Company_number", string.IsNullOrEmpty(model.Company_number) ? DBNull.Value : model.Company_number),
                    new SqlParameter("@Company_addr", string.IsNullOrEmpty(model.Company_addr) ? DBNull.Value : model.Company_addr),
                    new SqlParameter("@Company_busin", string.IsNullOrEmpty(model.Company_busin) ? DBNull.Value : model.Company_busin),
                    new SqlParameter("@Company_tel", string.IsNullOrEmpty(model.Company_tel) ? DBNull.Value : model.Company_tel),
                    new SqlParameter("@Company_fax", string.IsNullOrEmpty(model.Company_fax) ? DBNull.Value : model.Company_fax),
                    new SqlParameter("@Invoice_Iss", string.IsNullOrEmpty(model.Invoice_Iss) ? DBNull.Value : model.Invoice_Iss),
                    new SqlParameter("@Overseas", string.IsNullOrEmpty(model.Overseas) ? DBNull.Value : model.Overseas),
                    new SqlParameter("@add_num", model.add_num),
                    new SqlParameter("@add_ip", clientIp)
                };
                #endregion
                int result = _adoData.ExecuteNonQuery(T_SQL, parameters);

                // 2025.12.3 取消新增文中廠商資料
                //if (!string.IsNullOrEmpty(model.MF_ID))
                //{
                //    var _aWintonController = new A_WintonController();
                //    Task.Run(() => _aWintonController.SendManufacturer("1", model.MF_ID));
                //}

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
                var T_SQL = @"Update Manufacturer set MF_ID=@MF_ID, Company_name=@Company_name,Company_number=@Company_number,Company_addr=@Company_addr,
                    Company_busin=@Company_busin,Company_tel=@Company_tel,Company_fax=@Company_fax, 
                    Invoice_Iss=@Invoice_Iss,Overseas=@Overseas,edit_date=GETDATE(),edit_num=@edit_num,edit_ip=@edit_ip where MF_Number=@MF_Number";
                var parameters = new List<SqlParameter>
                {
                    new SqlParameter("@MF_ID", string.IsNullOrEmpty(model.MF_ID) ? DBNull.Value : model.MF_ID),
                    new SqlParameter("@Company_name", string.IsNullOrEmpty(model.Company_name) ? DBNull.Value : model.Company_name),
                    new SqlParameter("@Company_number", string.IsNullOrEmpty(model.Company_number) ? DBNull.Value : model.Company_number),
                    new SqlParameter("@Company_addr", string.IsNullOrEmpty(model.Company_addr) ? DBNull.Value : model.Company_addr),
                    new SqlParameter("@Company_busin", string.IsNullOrEmpty(model.Company_busin) ? DBNull.Value : model.Company_busin),
                    new SqlParameter("@Company_tel", string.IsNullOrEmpty(model.Company_tel) ? DBNull.Value : model.Company_tel),
                    new SqlParameter("@Company_fax", string.IsNullOrEmpty(model.Company_fax) ? DBNull.Value : model.Company_fax),
                    new SqlParameter("@Invoice_Iss", string.IsNullOrEmpty(model.Invoice_Iss) ? DBNull.Value : model.Invoice_Iss),
                    new SqlParameter("@Overseas", string.IsNullOrEmpty(model.Overseas) ? DBNull.Value : model.Overseas),
                    new SqlParameter("@edit_num", model.edit_num),
                    new SqlParameter("@edit_ip", clientIp),
                    new SqlParameter("@MF_Number", model.MF_Number)
                };
                #endregion
                int result = _adoData.ExecuteNonQuery(T_SQL, parameters);

                // 2025.12.3 取消新增文中廠商資料
                //if (!string.IsNullOrEmpty(model.MF_ID))
                //{
                //    var _aWintonController = new A_WintonController();
                //    Task.Run(() => _aWintonController.SendManufacturer("1", model.MF_ID));
                //}

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
