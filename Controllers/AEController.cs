using KF_WebAPI.BaseClass;
using KF_WebAPI.BaseClass.AE;
using KF_WebAPI.FunctionHandler;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.IdentityModel.Tokens;
using Newtonsoft.Json;
using System;
using System.Data;
using Microsoft.Data.SqlClient;
using System.IdentityModel.Tokens.Jwt;
using System.Net;
using System.Security.Claims;
using System.Text;

namespace KF_WebAPI.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class AEController : ControllerBase
    {
        private readonly string _storagePath = @"C:\UploadedFiles";

        [HttpPost("SendSMS")]
        public ActionResult<ResultClass<string>> SendSMS(string smbody, string dstaddr)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            try
            {
                StringBuilder reqUrl = new StringBuilder();
                reqUrl.Append("http://smsapi.mitake.com.tw/api/mtk/SmSend?CharsetURL=UTF-8");
                StringBuilder m_params = new StringBuilder();
                m_params.Append("username=52611690SMS");
                m_params.Append("&password=WgZi3m33KfJFnAPvBDwF");
                m_params.Append("&dstaddr=" + dstaddr);
                m_params.Append("&smbody=" + smbody);
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(new Uri(reqUrl.ToString()));
                request.Method = "POST";
                request.ContentType = "application/x-www-form-urlencoded";
                byte[] bs = Encoding.UTF8.GetBytes(m_params.ToString());
                request.ContentLength = bs.Length;
                request.GetRequestStream().Write(bs, 0, bs.Length);
             
                using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                {
                    using (StreamReader sr = new StreamReader(response.GetResponseStream()))
                    {
                        resultClass.objResult = sr.ReadToEnd();
                    }
                }
                resultClass.ResultCode = "000";
               
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "999";
                resultClass.ResultMsg = $" response: {ex.Message}";
            }
            return Ok(resultClass);
        }

        [HttpPost("CheckAPI")]
        public ActionResult<ResultClass<string>> CheckAPI()
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            try
            {

                resultClass.objResult = "";
                resultClass.ResultCode = "000";

            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "999";
                resultClass.ResultMsg = $" response: {ex.Message}";
            }
            return Ok(resultClass);
        }

        [HttpPost("Login")]
        public ActionResult<ResultClass<string>> Login(string user, string password)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            ADOData _adoData = new ADOData();
            var T_SQL = "SELECT * FROM User_M WHERE U_num = @UserName AND U_psw = @Password";
            var parameters = new List<SqlParameter>
            {
                new SqlParameter("@UserName", user),
                new SqlParameter("@Password", password)
            };

            try
            {
                DataTable dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
                if (dtResult.Rows.Count > 0)
                {
                    var Role_num = dtResult.Rows[0]["Role_num"].ToString();
                    var User_U_BC = dtResult.Rows[0]["U_BC"].ToString();

                    // 設置 Session
                    HttpContext.Session.SetString("UserID", user);
                    HttpContext.Session.SetString("Role_num", string.Join(',', Role_num));
                    HttpContext.Session.SetString("User_U_BC", User_U_BC);
                    var roleNum = HttpContext.Session.GetString("Role_num");
                    resultClass.objResult = User_U_BC;
                    return Ok(resultClass);
                }
                else
                {
                    resultClass.ResultCode = "401";
                    resultClass.ResultMsg = "用戶名或密碼不正確";
                    return Unauthorized(resultClass);
                }

            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass); // 返回 500 錯誤碼
            }
        }

        [HttpPost("GetMenuList")]
        public ActionResult<ResultClass<string>> GetMenuList(string U_num)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"SELECT distinct Menu_list.*  
                    FROM Menu_list  
                    LEFT JOIN (select * from Menu_set where U_num in (select Role_num from User_M where U_num=@U_num) AND del_tag='0') Menu_set ON Menu_list.Menu_id = Menu_set.Menu_id  
                    WHERE Menu_set.U_num in  (select Role_num from User_M where U_num=@U_num) and Menu_list.item_id = 0 
                    and (Menu_set.per_read = 1 or Menu_set.per_add = 1 or Menu_set.per_edit = 1 or Menu_set.per_del = 1) 
                    and Menu_list.del_tag = '0'
                    ORDER BY  Menu_list.top_id, Menu_list.sub_id, Menu_list.item_id";
                parameters.Add(new SqlParameter("@U_num", U_num));
                #endregion
                DataTable dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
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

        [HttpPost("ASP_File_Query")]
        public ActionResult<ResultClass<string>> ASP_File_Query(string cknum)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var User_Num = HttpContext.Session.GetString("UserID");
            try
            {
                ADOData _adoData = new ADOData();
                var T_SQL = @"select upload_name_show,FORMAT(add_date, 'yyyy/MM/dd', 'en-US') + ' ' + CASE WHEN DATEPART(HOUR, add_date) < 12 
                    THEN '上午' ELSE '下午' END + ' ' + FORMAT(add_date, 'hh:mm:ss', 'en-US') AS add_date,upload_id,Case When del_tag='1' 
                    Then FORMAT(del_date, 'yyyy/MM/dd', 'en-US') + ' ' + CASE WHEN DATEPART(HOUR, del_date) < 12 THEN '上午' 
                    ELSE '下午' END + ' ' + FORMAT(del_date, 'hh:mm:ss', 'en-US') else '0' end as del_tag 
                    from ASP_UpLoad where cknum=@cknum order by upload_id";
                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@cknum", cknum));

                DataTable dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
                if (dtResult.Rows.Count > 0)
                {
                    resultClass.ResultCode = "000";
                    resultClass.objResult = JsonConvert.SerializeObject(dtResult);
                }
                else
                {
                    resultClass.ResultCode = "201";
                    resultClass.ResultMsg = "查無資料";
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

        [HttpPost("ASP_File_Upload")]
        public ActionResult<ResultClass<string>> ASP_File_Upload(IFormFile file, string cknum)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var User_Num = HttpContext.Session.GetString("UserID");
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

            try
            {
                #region 宣告路徑
                // 獲取當前日期
                var now = DateTime.Now;
                var year = now.Year;
                var month = now.ToString("MM");
                var day = now.ToString("dd");

                // 設定儲存路徑
                var folderPath = Path.Combine(_storagePath, year.ToString() + month, year.ToString() + month + day);
                if (!Directory.Exists(folderPath))
                {
                    Directory.CreateDirectory(folderPath);
                }
                #endregion
                string number = FuncHandler.GetCheckNum();

                var fileExtension = Path.GetExtension(file.FileName);
                var filePath = Path.Combine(folderPath, $"{number}{fileExtension}");

                // 儲存檔案到指定的路徑
                using (var stream = new FileStream(filePath, FileMode.Create))
                {
                    file.CopyTo(stream);
                }

                var T_SQL = "Insert into ASP_UpLoad (cknum,upload_name_show,upload_name_code,upload_folder";
                T_SQL = T_SQL + ",upload_code,del_tag,add_date,add_num,add_ip,send_api_name_code) Values";
                T_SQL = T_SQL + " (@cknum,@upload_name_show,@upload_name_code,@upload_folder,@upload_code,@del_tag,@add_date,@add_num";
                T_SQL = T_SQL + " ,@add_ip,@send_api_name_code)";

                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@cknum", cknum));
                parameters.Add(new SqlParameter("@upload_name_show", file.FileName));
                parameters.Add(new SqlParameter("@upload_name_code", $"{number}{fileExtension}"));
                parameters.Add(new SqlParameter("@upload_folder", "AE_Web_UpLoad"));
                parameters.Add(new SqlParameter("@upload_code", number));
                parameters.Add(new SqlParameter("@del_tag", "0"));
                parameters.Add(new SqlParameter("@add_date", DateTime.Now));
                parameters.Add(new SqlParameter("@add_num", User_Num));
                parameters.Add(new SqlParameter("@add_ip", clientIp));
                parameters.Add(new SqlParameter("@send_api_name_code", ""));

                ADOData _adoData = new ADOData();
                int result = _adoData.ExecuteNonQuery(T_SQL, parameters);
                if (result == 0)
                {
                    resultClass.ResultCode = "201";
                    resultClass.ResultMsg = "上傳失敗";
                    return Ok(resultClass);
                }
                else
                {
                    resultClass.ResultCode = "000";
                    resultClass.ResultMsg = "上傳成功";
                    return Ok(resultClass);
                }
            }
            catch (Exception)
            {
                resultClass.ResultCode = "500";
                return StatusCode(500, resultClass); // 返回 500 錯誤碼
            }
            
        }

        [HttpGet("ASP_File_Download")]
        public IActionResult ASP_File_Download(string upload_id)
        {
            try
            {
                var T_SQL = "select upload_name_code,upload_name_show from ASP_UpLoad where upload_id=@upload_id";
                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@upload_id", upload_id));
                ADOData _adoData = new ADOData();
                DataTable dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
                if (dtResult.Rows.Count > 0)
                {
                    DataRow row = dtResult.Rows[0];
                    var upload_name_code = (row["upload_name_code"]).ToString();
                    var upload_name_show = (row["upload_name_show"]).ToString();

                    string _filePath = Path.Combine(_storagePath, upload_name_code.Substring(0, 6), upload_name_code.Substring(0, 8), upload_name_code);
                    if (!System.IO.File.Exists(_filePath))
                    {
                        return NotFound(); // 檔案不存在時返回 404
                    }
                    var fileBytes = System.IO.File.ReadAllBytes(_filePath);
                    var fileName = upload_name_show;

                    Response.Headers.Add("Content-Disposition", $"attachment; filename=\"{fileName}\"");

                    return File(fileBytes, "application/octet-stream", fileName);
                }
                else
                {
                    return NotFound(); // 檔案不存在時返回 404
                }
            }
            catch (Exception)
            {
                return StatusCode(500);
            }
        }

        [HttpPost("ASP_File_Del")]
        public ActionResult<ResultClass<string>> ASP_File_Del(string upload_id)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            var User_Num = HttpContext.Session.GetString("UserID");
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

            try
            {
                var T_SQL = "Update ASP_UpLoad set del_tag=@del_tag,del_date=@del_date,del_ip=@del_ip,del_num=@del_num where upload_id=@upload_id";
                var parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@del_tag", "1"));
                parameters.Add(new SqlParameter("@del_date", DateTime.Now));
                parameters.Add(new SqlParameter("@del_ip", clientIp));
                parameters.Add(new SqlParameter("@del_num", User_Num));
                parameters.Add(new SqlParameter("@upload_id", upload_id));

                ADOData _adoData = new ADOData();
                int result = _adoData.ExecuteNonQuery(T_SQL, parameters);
                if (result == 0)
                {
                    resultClass.ResultCode = "201";
                    resultClass.ResultMsg = "刪除失敗";
                    return Ok(resultClass);
                }
                else
                {
                    resultClass.ResultCode = "000";
                    resultClass.ResultMsg = "刪除成功";
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
        /// 提供所有分公司清單 item_M_code='branch_company'
        /// </summary>
        [HttpGet("GetBCList")]
        public ActionResult<ResultClass<string>> GetBCList()
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = "select item_D_code,item_D_name from Item_list where item_M_code='branch_company' and item_D_type='Y' and show_tag ='0'";
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
                return StatusCode(500, resultClass);
            }
        }

        /// <summary>
        /// 提供所有學歷 item_M_code = 'school_level'
        /// </summary>
        /// [HttpGet("Get_School_level")]
        [HttpGet("GetSchoolLevelList")]
        public ActionResult<ResultClass<string>> GetSchoolLevelList()
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = "select item_D_code,item_D_name from Item_list where item_M_code = 'school_level' AND item_D_type='Y' AND show_tag='0' AND del_tag='0' order by item_sort";
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
                return StatusCode(500, resultClass);
            }
        }

        /// <summary>
        ///提供所有職稱
        /// </summary>
        [HttpGet("GetTitleList")]
        public ActionResult<ResultClass<string>> GetTitleList()
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = "select item_D_code,item_D_name from Item_list where item_M_code='professional_title' and item_D_type='Y' and show_tag ='0'";
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
                return StatusCode(500, resultClass);
            }
        }

        /// <summary>
        /// 提供所有角色 select R_num,R_name from Role_M
        /// </summary>
        [HttpGet("GetRoleProfessionalList")]
        public ActionResult<ResultClass<string>> GetRoleProfessionalList()
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                #region SQL
                ADOData _adoData = new ADOData();
                var T_SQL = "select R_num,R_name from Role_M where del_tag = 0 order by r_id";
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
                return StatusCode(500, resultClass);
            }

        }

        /// <summary>
        /// 提供放款公司列表 GetCompanyList/_fn.asp
        /// </summary>
        [HttpGet("GetCompanyList")]
        public ActionResult<ResultClass<string>> GetCompanyList()
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = "select item_D_code,item_D_name from Item_list where item_M_code='fund_company' and item_D_type='Y' and del_tag='0' and show_tag='0'";
                #endregion
                var dtResult = _adoData.ExecuteSQuery(T_SQL);
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
        /// 提供撥款查詢年月 GetSendcaseYYYMM/Performance_Plot.asp
        /// </summary>
        [HttpGet("GetSendcaseYYYMM")]
        public ActionResult<ResultClass<string>> GetSendcaseYYYMM()
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = "select distinct convert(varchar(4),(convert(varchar(4),get_amount_date,126)-1911))+'-'+convert(varchar(2),month(get_amount_date)) yyyMM";
                T_SQL = T_SQL + " ,convert(varchar(7),get_amount_date, 126) yyyymm from House_sendcase";
                T_SQL = T_SQL + " where year(get_amount_date) > year(DATEADD(year,-2,SYSDATETIME()))";
                T_SQL = T_SQL + " order by convert(varchar(7),get_amount_date, 126) desc";
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
                return StatusCode(500, resultClass);
            }
        }

        /// <summary>
        /// 提供採購單申報類型列表
        /// </summary>
        [HttpGet("GetPMPayType")]
        public ActionResult<ResultClass<string>> GetPMPayType()
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = "select item_D_code,item_D_name from Item_list where item_M_code='Procurement_Pay' and item_D_type='Y' and del_tag='0' and show_tag='0'";
                #endregion
                var dtResult = _adoData.ExecuteSQuery(T_SQL);
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
