using KF_WebAPI.BaseClass;
using KF_WebAPI.BaseClass.AE;
using KF_WebAPI.BaseClass.Max104;
using KF_WebAPI.DataLogic;
using KF_WebAPI.FunctionHandler;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using Microsoft.IdentityModel.Tokens;
using Newtonsoft.Json;
using System;
using System.Data;
using System.IdentityModel.Tokens.Jwt;
using System.Net;
using System.Reflection;
using System.Security.Claims;
using System.Text;

namespace KF_WebAPI.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class AEController : ControllerBase
    {
       // private readonly string _storagePath = @"C:\UploadedFiles";
        private readonly string _storagePath = @"D:\AE_Web_UpLoad";
        private AEData _AEData = new();
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
                return StatusCode(500, resultClass); 
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
                var T_SQL = @"SELECT distinct Menu_list.*  
                    FROM Menu_list  
                    LEFT JOIN (select * from Menu_set where U_num in (select Role_num from User_M where U_num=@U_num) AND del_tag='0') Menu_set ON Menu_list.Menu_id = Menu_set.Menu_id  
                    WHERE Menu_set.U_num in  (select Role_num from User_M where U_num=@U_num) and Menu_list.item_id = 0 
                    and (Menu_set.per_read = 1 or Menu_set.per_add = 1 or Menu_set.per_edit = 1 or Menu_set.per_del = 1) 
                    and Menu_list.del_tag = '0'
                    ORDER BY  Menu_list.top_id, Menu_list.sub_id, Menu_list.item_id";
                var parameters = new List<SqlParameter> 
                {
                    new SqlParameter("@U_num", U_num)
                };
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
                return StatusCode(500, resultClass); 
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
                var _Fun = new FuncHandler();
                #region SQL
                var T_SQL = @"select upload_name_show,FORMAT(add_date, 'yyyy/MM/dd', 'en-US') + ' ' + CASE WHEN DATEPART(HOUR, add_date) < 12 
                    THEN '上午' ELSE '下午' END + ' ' + FORMAT(add_date, 'hh:mm:ss', 'en-US') AS add_date,upload_id,Case When del_tag='1' 
                    Then FORMAT(del_date, 'yyyy/MM/dd', 'en-US') + ' ' + CASE WHEN DATEPART(HOUR, del_date) < 12 THEN '上午' 
                    ELSE '下午' END + ' ' + FORMAT(del_date, 'hh:mm:ss', 'en-US') else '0' end as del_tag 
                    from ASP_UpLoad where cknum=@cknum order by del_date,upload_id desc";
                var parameters = new List<SqlParameter>
                {
                    new SqlParameter("@cknum", cknum)
                };
                #endregion
                var result = _adoData.ExecuteQuery(T_SQL, parameters).AsEnumerable().Select(row => new
                {
                    upload_name_show = _Fun.DeCodeBNWords(row.Field<string>("upload_name_show")),
                    add_date = row.Field<string>("add_date"),
                    upload_id = row.Field<decimal>("upload_id"),
                    del_tag = row.Field<string>("del_tag"),
                }).ToList();
                if (result.Count() > 0)
                {
                    resultClass.ResultCode = "000";
                    resultClass.objResult = JsonConvert.SerializeObject(result);
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
        public ActionResult<ResultClass<string>> ASP_File_Upload(IFormFile file, string cknum, string WebUser)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
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

                var T_SQL = @"Insert into ASP_UpLoad (cknum,upload_name_show,upload_name_code,upload_folder,
                    upload_code,del_tag,add_date,add_num,add_ip,send_api_name_code) Values
                    ( @cknum,@upload_name_show,@upload_name_code,@upload_folder,@upload_code,@del_tag,@add_date,@add_num,
                    @add_ip,@send_api_name_code )";

                var parameters = new List<SqlParameter> 
                {
                    new SqlParameter("@cknum", cknum),
                    new SqlParameter("@upload_name_show", file.FileName),
                    new SqlParameter("@upload_name_code", $"{number}{fileExtension}"),
                    new SqlParameter("@upload_folder", "AE_Web_UpLoad"),
                    new SqlParameter("@upload_code", number),
                    new SqlParameter("@del_tag", "0"),
                    new SqlParameter("@add_date", DateTime.Now),
                    new SqlParameter("@add_num", WebUser),
                    new SqlParameter("@add_ip", clientIp),
                    new SqlParameter("@send_api_name_code", "")
                };

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
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass); 
            }

        }

        [HttpGet("ASP_File_Download")]
        public IActionResult ASP_File_Download(string upload_id)
        {
            try
            {
                var _Fun = new FuncHandler();
                var T_SQL = "select upload_name_code,upload_name_show from ASP_UpLoad where upload_id=@upload_id";
                var parameters = new List<SqlParameter> 
                {
                    new SqlParameter("@upload_id", upload_id)
                };
                ADOData _adoData = new ADOData();
                DataTable dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
                if (dtResult.Rows.Count > 0)
                {
                    DataRow row = dtResult.Rows[0];
                    var upload_name_code = (row["upload_name_code"]).ToString();
                    var upload_name_show = _Fun.DeCodeBNWords((row["upload_name_show"]).ToString());

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
        public ActionResult<ResultClass<string>> ASP_File_Del(string upload_id, string WebUser)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

            try
            {
                ADOData _adoData = new ADOData();
                var T_SQL = "Update ASP_UpLoad set del_tag=@del_tag,del_date=@del_date,del_ip=@del_ip,del_num=@del_num where upload_id=@upload_id";
                var parameters = new List<SqlParameter> 
                {
                    new SqlParameter("@del_tag", "1"),
                    new SqlParameter("@del_date", DateTime.Now),
                    new SqlParameter("@del_ip", clientIp),
                    new SqlParameter("@del_num", WebUser),
                    new SqlParameter("@upload_id", upload_id)
                };

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
                var _Fun = new FuncHandler();
                var T_SQL = "select R_num,R_name from Role_M where del_tag = 0 order by r_id";
                #endregion
                var result = _adoData.ExecuteSQuery(T_SQL).AsEnumerable().Select(row => new
                {
                    R_num = row.Field<string>("R_num"),
                    R_name = _Fun.DeCodeBNWords(row.Field<string>("R_name"))
                }).ToList();
                if (result.Count > 0)
                {
                    resultClass.ResultCode = "000";
                    resultClass.objResult = JsonConvert.SerializeObject(result);
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
                var T_SQL = @"select distinct convert(varchar(4),(convert(varchar(4),get_amount_date,126)-1911))+'-'+convert(varchar(2),month(get_amount_date)) yyyMM
                    ,convert(varchar(7),get_amount_date, 126) yyyymm from House_sendcase
                    where year(get_amount_date) > year(DATEADD(year,-2,SYSDATETIME()))
                    order by convert(varchar(7),get_amount_date, 126) desc";
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
        /// 提供所有訊息類型及筆數 GetMesListCount
        /// </summary>
        [HttpGet("GetMesListCount")]
        public ActionResult<ResultClass<string>> GetMesListCount(string User, DateTime StartDate, DateTime EndDate)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL_M = @"select item_D_code,item_D_name from Item_list where item_M_code = 'Msg_kind' AND item_D_type='Y'  
                    AND del_tag='0'  order by item_sort";
                #endregion
                var dtResult_m = _adoData.ExecuteSQuery(T_SQL_M);

                #region 組動態語法
                var caseStatements = new List<string>();
                var sumStatements = new List<string>();
                var labels = new List<string>();
                var modelList = new List<Message>();

                foreach (DataRow row in dtResult_m.Rows)
                {
                    var model = new Message();
                    model.Name = row["item_D_name"].ToString();
                    model.MsgID = row["item_D_code"].ToString();
                    model.Count = 0;

                    string caseStatement = $@" (CASE WHEN Msg.Msg_kind='{model.MsgID}' THEN 1 ELSE 0 END) AS 'Msg_kind_{model.MsgID}'";
                    caseStatements.Add(caseStatement);
                    sumStatements.Add($"ISNULL(sum(Msg_kind_{model.MsgID}),0) AS '{model.Name}'");

                    labels.Add(model.Name);
                    modelList.Add(model);
                }
                string caseColumns = string.Join(",\n       ", caseStatements);
                string sumColumns = string.Join(",\n       ", sumStatements);

                string T_SQL = $@"SELECT  {sumColumns} FROM ( SELECT *,{caseColumns} FROM Msg WHERE del_tag = '0' AND Msg_read_type = 'N' 
                               AND msg_to_num = @num and Msg_show_date between @StartDate and @EndDate) AS SubQuery";
                var parameters = new List<SqlParameter> 
                {
                    new SqlParameter("@num", User),
                    new SqlParameter("@StartDate", StartDate),
                    new SqlParameter("@EndDate", EndDate)
                };
                #endregion
                var dtResult = _adoData.ExecuteQuery(T_SQL, parameters);

                foreach (DataRow row in dtResult.Rows)
                {
                    foreach (var model in modelList)
                    {
                        string columnName = model.Name;
                        if (dtResult.Columns.Contains(columnName))
                        {
                            model.Count = Convert.ToInt32(row[columnName]);
                        }
                    }
                }
                resultClass.ResultCode = "000";
                resultClass.ResultMsg = "查詢成功";
                resultClass.objResult = JsonConvert.SerializeObject(modelList);
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
        /// 訊息列表查詢 MesList_Query
        /// </summary>
        [HttpPost("MesList_Query")]
        public ActionResult<ResultClass<string>> MesList_Query(Message_Req model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                

                var queryBuilder = new StringBuilder();

                queryBuilder.AppendLine(@"select Msg_id,Case When Msg_source ='sys' Then '系統通知' Else Msg_source End as Msg_source,
                                          li.item_D_name as Msg_kind_Name,
                                          Um.U_name as Msg_to_name,FORMAT(Msg_show_date, 'yyyy/MM/dd tt hh:mm:ss') AS Msg_show_date,
                                          Msg_title,Msg_note,Um_1.U_name as add_num_name,Msg_read_type 
                                          from Msg
                                          Left join Item_list li on li.item_M_code='Msg_kind' and li.item_D_code=Msg.Msg_kind
                                          Left join User_M Um on Um.U_num=Msg.Msg_to_num
                                          Left join User_M Um_1 on Um_1.U_num=Msg.add_num
                                          where Msg_show_date between @StartDate and @EndDate");
                var parameters = new List<SqlParameter> 
                {
                    new SqlParameter("@StartDate", model.StartDate),
                    new SqlParameter("@EndDate", model.EndDate)
                };
                if (model.UserType == "kind_get")
                {
                    queryBuilder.AppendLine(" and Msg.Msg_to_num = @User");
                    parameters.Add(new SqlParameter("@User", model.User));
                }
                if (model.UserType == "kind_add")
                {
                    queryBuilder.AppendLine(" and Msg.add_num = @User");
                    parameters.Add(new SqlParameter("@User", model.User));
                }
                if (model.ReadType != "X")
                {
                    queryBuilder.AppendLine(" and Msg.Msg_read_type = @ReadType");
                    parameters.Add(new SqlParameter("@ReadType", model.ReadType));
                }
                if (!string.IsNullOrEmpty(model.Mes_Kind))
                {
                    queryBuilder.AppendLine(" and Msg.Msg_kind = @Msg_kind");
                    parameters.Add(new SqlParameter("@Msg_kind", model.Mes_Kind));
                }
                queryBuilder.AppendLine(" order by Msg_show_date desc");
                var sqlQuery = queryBuilder.ToString();
                #endregion
                var drResult = _adoData.ExecuteQuery(sqlQuery, parameters);

                resultClass.ResultCode = "000";
                resultClass.ResultMsg = "查詢成功";
                resultClass.objResult = JsonConvert.SerializeObject(drResult);

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
        /// 所有未讀訊息改為已讀 MessageALL_Read
        /// </summary>
        [HttpPost("MessageALL_Read")]
        public ActionResult<ResultClass<string>> MessageALL_Read(Message_Req model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();
            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var queryBuilder = new StringBuilder();
                queryBuilder.AppendLine( @"Update Msg set Msg_read_type='Y',edit_date=GETDATE(),edit_num=@User,edit_ip=@UserIp 
                                           where Msg_read_type='N' and Msg_show_date between @StartDate and @EndDate");
                var parameters = new List<SqlParameter> 
                {
                    new SqlParameter("@User", model.User),
                    new SqlParameter("@UserIp", clientIp),
                    new SqlParameter("@StartDate", model.StartDate),
                    new SqlParameter("@EndDate", model.EndDate)
                };

                if (model.UserType == "kind_get")
                {
                    queryBuilder.AppendLine(" and Msg.Msg_to_num = @User");
                  
                }
                if (model.UserType == "kind_add")
                {
                    queryBuilder.AppendLine(" and Msg.add_num = @User");
                }
                if (model.ReadType != "X")
                {
                    queryBuilder.AppendLine(" and Msg.Msg_read_type = @ReadType");
                    parameters.Add(new SqlParameter("@ReadType", model.ReadType));
                }
                if (!string.IsNullOrEmpty(model.Mes_Kind))
                {
                    queryBuilder.AppendLine(" and Msg.Msg_kind = @Msg_kind");
                    parameters.Add(new SqlParameter("@Msg_kind", model.Mes_Kind));
                }
                var sqlQuery = queryBuilder.ToString();
                #endregion
                int result = _adoData.ExecuteNonQuery(sqlQuery, parameters);
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
                return StatusCode(500, resultClass);
            }
        }

        /// <summary>
        /// 單筆變更已讀 MesRead_Upd
        /// </summary>
        [HttpGet("MesRead_Upd")]
        public ActionResult<ResultClass<string>> MesRead_Upd(string Msg_id,string User)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();
            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"Update Msg set Msg_read_type='Y',edit_date=GETDATE(),edit_num=@User,edit_ip=@UserIp where Msg_id=@Msg_id";
                var parameters = new List<SqlParameter> 
                {
                    new SqlParameter("@Msg_id", Msg_id),
                    new SqlParameter("@User", User),
                    new SqlParameter("@UserIp", clientIp)
                };
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
                return StatusCode(500, resultClass);
            }
        }

        /// <summary>
        /// 特殊權限判定
        /// </summary>
        [HttpGet("SpecialCkeck")]
        public ActionResult<ResultClass<string>> SpecialCkeck(string User,string Sp_ID)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                var T_SQL = @"select * from Special_list ST inner join Special_set SS on SS.sp_id = ST.sp_id 
                    where ST.del_tag='0' and SS.del_tag='0' and SS.U_num = @User and SS.sp_id = @Sp_ID";
                var parameters = new List<SqlParameter>
                {
                    new SqlParameter("@User",User),
                    new SqlParameter("@Sp_ID",Sp_ID)
                };
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
        /// 取得代碼資料
        /// </summary>
        [HttpGet("GetItemList")]
        public ActionResult<ResultClass<string>> GetItemList(string item_M_code)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                var _Fun = new FuncHandler();
                var T_SQL = @"select item_D_code,item_D_name from Item_list 
                              where item_M_code=@item_M_code and item_D_type='Y' and del_tag='0' order by item_sort";
                var parameters = new List<SqlParameter>
                {
                    new SqlParameter("@item_M_code",item_M_code)
                };
                var result = _adoData.ExecuteQuery(T_SQL,parameters).AsEnumerable().Select(row => new
                {
                    item_D_code = row.Field<string>("item_D_code"),
                    item_D_name = _Fun.DeCodeBNWords(row.Field<string>("item_D_name"))
                }).ToList();
                resultClass.ResultCode = "000";
                resultClass.objResult = JsonConvert.SerializeObject(result);
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
        /// 取得GUID
        /// </summary>
        [HttpGet("GetWebToken")]
        public ActionResult<ResultClass<string>> GetWebToken(string chknum,string Type,string user)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

            try
            {
                ADOData _adoData = new ADOData();
                string guid = Guid.NewGuid().ToString().ToUpper();
                #region SQL
                var T_SQL = @"Insert into AE_WebToken(GUID,chk_num,TokenType,isConfirm,Effect_time,add_num,add_date,add_ip,ErrCount,isVerify) 
                              Values (@GUID,@chk_num,@TokenType,'N',DATEADD(MINUTE, 30, GETDATE()),@add_num,GETDATE(),@add_ip,0,'N')";
                var parameters = new List<SqlParameter> 
                {
                    new SqlParameter("@GUID",guid),
                    new SqlParameter("@chk_num",chknum),
                    new SqlParameter("@TokenType",Type),
                    new SqlParameter("@add_num",user),
                    new SqlParameter("@add_ip",clientIp)
                };
                #endregion
                int Result = _adoData.ExecuteNonQuery(T_SQL, parameters);
                if (Result == 0)
                {
                    resultClass.ResultCode = "400";
                    resultClass.ResultMsg = "Token失敗,請洽資訊人員";
                }
                else
                {
                    resultClass.ResultCode = "000";
                    resultClass.objResult = guid;
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
        /// 驗證GUID
        /// </summary>
        [HttpGet("CheckWebToken")]
        public ActionResult<ResultClass<string>> CheckWebToken(string GUID)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL_Token 確認GUID是否有效且抓取對應者
                var T_SQL = @" select 　chk_num [User],U_BC,isVerify from AE_WebToken A Left Join User_M M on A.chk_num=M.U_num where GUID=@GUID and GETDATE() < Effect_time and isConfirm='N'";
                var parameters = new List<SqlParameter>
                {
                    new SqlParameter("@GUID", GUID)
                };
                #endregion
                var dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
                if (dtResult.Rows.Count > 0)
                {
                   
                    dtResult.AsEnumerable().Select(row => new
                    {
                        User = row.Field<string>("User"),
                        U_BC = row.Field<string>("U_BC"),
                        isVerify = row.Field<string>("isVerify")
                    });
                    resultClass.ResultCode = "000";
                    resultClass.objResult = JsonConvert.SerializeObject(dtResult);
                    /*不需要驗證的才可以直接更新isConfirm='Y'*/
                    if (dtResult.Rows[0]["isVerify"].ToString() == "N")
                    {
                        #region SQL_UPDATE
                        var T_SQL_U = @"Update AE_WebToken set isConfirm='Y',Confirm_date=SYSDATETIME() where GUID=@GUID";
                        var parameters_u = new List<SqlParameter>
                        {
                            new SqlParameter("@GUID", GUID)
                        };
                        _adoData.ExecuteNonQuery(T_SQL_U, parameters_u);
                        #endregion
                    }

                    return Ok(resultClass);
                }
                else
                {
                    resultClass.ResultCode = "401";
                    resultClass.ResultMsg = $"Token失效請重新抓取網址";
                    return StatusCode(401, resultClass);
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
        /// 更新WebToken isConfirm By  GUID
        /// </summary>
        [HttpGet("UpdWebToken")]
        public ActionResult<ResultClass<string>> UpdWebToken(string GUID)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL_UPDATE
                var T_SQL_U = @"Update AE_WebToken set isConfirm='Y',Confirm_date=SYSDATETIME() where GUID=@GUID";
                var parameters_u = new List<SqlParameter>
                        {
                            new SqlParameter("@GUID", GUID)
                        };
                _adoData.ExecuteNonQuery(T_SQL_U, parameters_u);
                #endregion
                resultClass.ResultCode = "000";
                resultClass.objResult = "更新成功";
                return Ok(resultClass);

            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass);
            }
        }



        [Route("GetFileBySeq")]
        [HttpPost]
        public ActionResult<string> GetFileBySeq(string KeyID, string Type, string Index)
        {
            ResultClass<AE_Files> resultClass = new();
            try
            {
                resultClass = _AEData.GetFile(KeyID, Type, Index);
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "999";
                resultClass.ResultMsg = ex.Message;
            }
            return Ok(resultClass);
        }

        [Route("GetFilesByKeyID")]
        [HttpPost]
        public ActionResult<string> GetFilesByKeyID(string KeyID, string Type)
        {
            ResultClass<AE_Files[]> resultClass = new();
            try
            {
                resultClass = _AEData.GetFilesByKeyID(KeyID, Type);
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "999";
                resultClass.ResultMsg = ex.Message;
            }
            return Ok(resultClass);
        }

        [Route("InsertFile")]
        [HttpPost]
        public ActionResult<ResultClass<int>> InsertFile([FromBody] AE_Files[] AE_Files, string KeyID, string Type, string u_num)
        {
            ResultClass<int> resultClass = new();
            try
            {
                int m_Execut  = _AEData.InsertFile(AE_Files, KeyID, Type, u_num);
                resultClass.ResultCode = "000";
                resultClass.ResultMsg = "";
                resultClass.objResult = m_Execut;
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "999";
                resultClass.ResultMsg = ex.Message;
            }
            return Ok(resultClass);
        }


        [Route("DeleteFile")]
        [HttpPost]
        public ActionResult<ResultClass<int>> DeleteFile(string KeyID, string Type, string file_index)
        {
            ResultClass<int> resultClass = new();
            try
            {
                int m_Execut  = _AEData.DeleteFile( KeyID, Type, file_index);
                resultClass.ResultCode = "000";
                resultClass.ResultMsg = "";
                resultClass.objResult = m_Execut;
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "999";
                resultClass.ResultMsg = ex.Message;
            }
            return Ok(resultClass);
        }

        [Route("GetWebsiteURL")]
        [HttpPost]
        public ActionResult<ResultClass<string>> GetWebsiteURL()
        {
            ResultClass<string> resultClass = new();
            try
            {
                string m_URL = _AEData.GetWebsiteURL();
                resultClass.ResultCode = "000";
                resultClass.ResultMsg = "";
                resultClass.objResult = m_URL;
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "999";
                resultClass.ResultMsg = ex.Message;
            }
            return Ok(resultClass);
        }

        


    }

}
