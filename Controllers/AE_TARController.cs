using KF_WebAPI.BaseClass;
using KF_WebAPI.BaseClass.AE;
using KF_WebAPI.BaseClass.Max104;
using KF_WebAPI.FunctionHandler;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Http.Extensions;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Microsoft.SqlServer.Server;
using Newtonsoft.Json;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlTypes;
using System.Reflection;
using System.Reflection.Metadata;
using System.Security.Cryptography;
using System.Text.RegularExpressions;

namespace KF_WebAPI.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class AE_TARController : ControllerBase
    {
        /// <summary>
        /// 職務責任額列表
        /// </summary>
        [HttpGet("Pro_Target_LQuery")]
        public ActionResult<ResultClass<string>> Pro_Target_LQuery(string? Title_name, string? PR_DATE)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"select PR_ID,PR_title,Li.item_D_name as PR_title_Name,PR_target,PR_Date,
                              RIGHT('000' + CAST(YEAR(CAST(PR_Date + '-01' AS DATE)) - 1911 AS VARCHAR), 3) + '-' +
                              RIGHT('00' + CAST(MONTH(CAST(PR_Date + '-01' AS DATE)) AS VARCHAR), 2) AS PR_Date_Minguo
                              from Professional_target Pr
                              left join Item_list Li on Li.item_M_code='professional_title' and Li.item_D_code=Pr.PR_title
                              where 1 = 1";
                if (!string.IsNullOrEmpty(Title_name))
                {
                    T_SQL += " and Li.item_D_name = @Title_name";
                    parameters.Add(new SqlParameter("@Title_name", Title_name));
                }
                if (!string.IsNullOrEmpty(PR_DATE))
                {
                    T_SQL += " and PR_DATE = @PR_DATE";
                    parameters.Add(new SqlParameter("@PR_DATE", PR_DATE));
                }
                T_SQL += " order by PR_Date desc,Li.item_sort";
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
                return StatusCode(500, resultClass);
            }
        }

        /// <summary>
        /// 修改職稱責任額
        /// </summary>
        [HttpPost("Pro_Target_Upd")]
        public ActionResult<ResultClass<string>> Pro_Target_Upd(Pro_Target model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();
            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"Update Professional_target set PR_target=@PR_target,PR_Date=@PR_Date,edit_date=getdate(),
                              edit_num=@edit_num,edit_ip=@edit_ip where PR_ID=@PR_ID";
                var parameters = new List<SqlParameter>()
                {
                    new SqlParameter("@PR_target", model.PR_target),
                    new SqlParameter("@PR_Date",model.PR_Date),
                    new SqlParameter("@edit_num",model.user),
                    new SqlParameter("@edit_ip",clientIp),
                    new SqlParameter("@PR_ID",model.PR_ID)
                };
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
                    //同步修改其他同職稱業務責任額
                    var T_SQL_U = @"Update Person_target set PE_target=@PE_target,edit_date=getdate(),edit_num=@edit_num 
                                    ,edit_ip=@edit_ip where PE_title=@PE_title and PE_Date=@PE_Date";
                    var parameters_u = new List<SqlParameter>()
                    {
                        new SqlParameter("@PE_target",model.PR_target),
                        new SqlParameter("@edit_num",model.user),
                        new SqlParameter("@edit_ip",clientIp),
                        new SqlParameter("@PE_title",model.PR_title),
                        new SqlParameter("@PE_Date",model.PR_Date)
                    };
                    int result_u = _adoData.ExecuteNonQuery(T_SQL_U, parameters_u);
                    if (result_u == 0) 
                    {
                        resultClass.ResultCode = "400";
                        resultClass.ResultMsg = "同步業務責任額失敗";
                        return Ok(resultClass);
                    }
                    else
                    {
                        resultClass.ResultCode = "000";
                        resultClass.ResultMsg = "修改成功";
                        return Ok(resultClass);
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
        /// 取得業務職稱
        /// </summary>
        [HttpGet("GetproFessionalTitle")]
        public ActionResult<ResultClass<string>> GetproFessionalTitle()
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"select item_D_code,item_D_name,Case item_D_code When 'PFT030' Then 650 When 'PFT050' Then 550
                              When 'PFT060' Then 450 When 'PFT300' Then 350 Else 0 END as targetInt
                              from Item_list where item_M_code='professional_title' 
                              and item_D_code IN 
                              (select U_PFT from User_M where Role_num in ('1009') 
                              and U_leave_date is null and U_num <> 'K9999')
                              order by item_sort";
                #endregion
                var dtResult=_adoData.ExecuteSQuery(T_SQL);
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
        /// 新增職稱責任額
        /// </summary>
        [HttpPost("Pro_Target_Ins")]
        public ActionResult<ResultClass<string>> Pro_Target_Ins(List<Pro_Target_Ins> list)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL

                #region 先檢查有無資料
                var T_SQL_C = @"select * from Professional_target where PR_Date=@PR_Date";
                var parameters_c = new List<SqlParameter>()
                {
                    new SqlParameter("@PR_Date",list[0].PR_Date)
                };
                var dtReult_c = _adoData.ExecuteQuery(T_SQL_C, parameters_c);
                if (dtReult_c.Rows.Count > 0)
                {
                    resultClass.ResultCode = "400";
                    resultClass.ResultMsg = "已有資料";
                    return BadRequest(resultClass);
                }
                #endregion

                var T_SQL = @"Insert into Professional_target(PR_title,PR_target,PR_Date,add_date,add_num,add_ip)
                              Values (@PR_title,@PR_target,@PR_Date,GETDATE(),@add_num,@add_ip)";
                foreach (Pro_Target_Ins item in list)
                {
                    var parameters = new List<SqlParameter>
                    {
                        new SqlParameter("@PR_title", item.PR_title),
                        new SqlParameter("@PR_target", item.PR_target),
                        new SqlParameter("@PR_Date", item.PR_Date),
                        new SqlParameter("@add_num", item.user),
                        new SqlParameter("@add_ip", clientIp)
                    };
                    int result = _adoData.ExecuteNonQuery(T_SQL, parameters);
                    if(result == 0)
                    {
                        resultClass.ResultCode = "400";
                        resultClass.ResultMsg = "新增失敗";
                        return BadRequest(resultClass);
                    }
                }
                #endregion

                //呼叫SP=>業務個人責任額新增
                var T_SQL_SP = "exec UpdatePersonTargets @PR_Date";
                var parameters_sp = new List<SqlParameter>()
                {
                    new SqlParameter("@PR_Date",list[0].PR_Date)
                };
                _adoData.ExecuteQuery(T_SQL_SP, parameters_sp);

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
        /// 複製上期責任額
        /// </summary>
        [HttpGet("Pro_Target_Clone")]
        public ActionResult<ResultClass<string>> Pro_Target_Clone(string PR_DATE,string User)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();
            try
            {
                ADOData _adoData = new ADOData();

                #region 檢查有無上期資料
                var parts = PR_DATE.Split('-');
                int year = int.Parse(parts[0]);
                int month = int.Parse(parts[1]);

                DateTime current = new DateTime(year, month, 1);
                DateTime previous = current.AddMonths(-1);

                var old_PR_DATE = $"{previous.Year}-{previous.Month.ToString("D2")}";

                var T_SQL_C = @"select * from Professional_target where PR_DATE=@PR_DATE";
                var parameters_c = new List<SqlParameter>()
                {
                    new SqlParameter("@PR_DATE",old_PR_DATE)
                };
                #endregion

                var dtResult_c = _adoData.ExecuteQuery(T_SQL_C, parameters_c);
                if (dtResult_c.Rows.Count > 0)
                {
                    var TargetList = dtResult_c.AsEnumerable().Select(row => new Pro_Target_Ins
                    {
                        PR_title = row.Field<string>("PR_title"),
                        PR_target = row.Field<int>("PR_target"),
                        PR_Date = row.Field<string>("PR_Date"),
                    }).ToList();

                    #region 檢查是否已有資料
                    var T_SQL_R = @"select * from Professional_target where PR_Date=@PR_Date";
                    var parameters_r = new List<SqlParameter>()
                    {
                        new SqlParameter("@PR_Date",PR_DATE)
                    };
                    var dtReult_r = _adoData.ExecuteQuery(T_SQL_R, parameters_r);
                    if (dtReult_r.Rows.Count > 0)
                    {
                        resultClass.ResultCode = "400";
                        resultClass.ResultMsg = "已有資料";
                        return BadRequest(resultClass);
                    }
                    #endregion

                    var T_SQL = @"Insert into Professional_target(PR_title,PR_target,PR_Date,add_date,add_num,add_ip)
                              Values (@PR_title,@PR_target,@PR_Date,GETDATE(),@add_num,@add_ip)";

                    foreach (Pro_Target_Ins item in TargetList)
                    {
                        var parameters = new List<SqlParameter>
                        {
                            new SqlParameter("@PR_title", item.PR_title),
                            new SqlParameter("@PR_target", item.PR_target),
                            new SqlParameter("@PR_Date", PR_DATE),
                            new SqlParameter("@add_num", User),
                            new SqlParameter("@add_ip", clientIp)
                        };
                        int result = _adoData.ExecuteNonQuery(T_SQL, parameters);
                        if (result == 0)
                        {
                            resultClass.ResultCode = "400";
                            resultClass.ResultMsg = "新增失敗";
                            return BadRequest(resultClass);
                        }
                    }

                    //呼叫SP=>業務個人責任額新增
                    var T_SQL_SP = "exec UpdatePersonTargets @PR_Date";
                    var parameters_sp = new List<SqlParameter>() 
                    {
                        new SqlParameter("@PR_Date",PR_DATE)
                    };
                    _adoData.ExecuteQuery(T_SQL_SP, parameters_sp);

                    resultClass.ResultCode = "000";
                    resultClass.ResultMsg = "新增成功";
                    return Ok(resultClass);
                }
                else
                {
                    resultClass.ResultCode = "400";
                    resultClass.ResultMsg = "無上期資料";
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
        /// 業務責任額列表
        /// </summary>
        [HttpGet("Per_Target_LQuery")]
        public ActionResult<ResultClass<string>> Per_Target_LQuery(string? c_name,string? PE_DATE)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"select PE_ID,Li.item_D_name as titleName,Case When ISNULL(Lis.item_D_name,'') <> '' THEN Lis.item_D_name ELSE Um.U_name END as U_name,
                              PE_num,PE_target,PE_Date,
                              RIGHT('000' + CAST(YEAR(CAST(PE_Date + '-01' AS DATE)) - 1911 AS VARCHAR), 3) + '-' +
                              RIGHT('00' + CAST(MONTH(CAST(PE_Date + '-01' AS DATE)) AS VARCHAR), 2) AS PE_Date_Minguo
                              from Person_target Pe
                              left join Item_list Li on Li.item_M_code='professional_title' and Li.item_D_code=Pe.PE_title
                              left join User_M Um on Um.U_num=Pe.PE_num
                              left join Item_list Lis on Lis.item_M_code = 'SpecName' AND Lis.item_D_type = 'Y' and Lis.item_D_txt_A = Um.U_num
                              where PE_title IN ('PFT060','PFT030','PFT050','PFT300') ";
                if (!string.IsNullOrEmpty(c_name))
                {
                    T_SQL += " and U_name like @c_name";
                    parameters.Add(new SqlParameter("@c_name", "%" + c_name + "%"));
                }
                if (!string.IsNullOrEmpty(PE_DATE))
                {
                    T_SQL += " and PE_DATE = @PE_DATE";
                    parameters.Add(new SqlParameter("@PE_DATE", PE_DATE));
                }
                T_SQL += " order by PE_Date desc,Li.item_sort";
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
        /// 修改業務業績額
        /// </summary>
        [HttpPost("Per_Target_Upd")]
        public ActionResult<ResultClass<string>> Per_Target_Upd(Per_Target model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"Update Person_target set PE_target=@PE_target,PE_title=@PE_title,edit_date=getdate(),
                             edit_num=@edit_num,edit_ip=@edit_ip where PE_ID=@PE_ID";
                var parameters = new List<SqlParameter>()
                {
                    new SqlParameter("@PE_target",model.PE_target),
                    new SqlParameter("@PE_title",model.PE_title),
                    new SqlParameter("@edit_num",model.user),
                    new SqlParameter("@edit_ip",clientIp),
                    new SqlParameter("@PE_ID",model.PE_ID)
                };
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
        /// 新增業務責任額
        /// </summary>
        [HttpPost("Per_Target_Ins")]
        public ActionResult<ResultClass<string>> Per_Target_Ins(List<Per_Target_Ins> list)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

            try
            {
                ADOData _adoData = new ADOData();

                var T_SQL = @"Insert into Person_target(PE_title,PE_num,PE_target,PE_Date,add_date,add_num,add_ip) Values 
                              ((select U_PFT from User_M where U_num=@PE_num),@PE_num,@PE_target,@PE_Date,GETDATE(),@add_num,@add_ip)";
                foreach (Per_Target_Ins item in list)
                {
                    var parameters = new List<SqlParameter>
                    {
                        new SqlParameter("@PE_num", item.PE_num),
                        new SqlParameter("@PE_target", item.PE_target),
                        new SqlParameter("@PE_Date", item.PE_Date),
                        new SqlParameter("@add_num", item.user),
                        new SqlParameter("@add_ip", clientIp)
                    };
                    int result = _adoData.ExecuteNonQuery(T_SQL, parameters);
                    if (result == 0)
                    {
                        resultClass.ResultCode = "400";
                        resultClass.ResultMsg = "新增失敗";
                        return BadRequest(resultClass);
                    }
                }

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
        /// 取得現有資料的年月份
        /// </summary>
        [HttpGet("GetTargetYM")]
        public ActionResult<ResultClass<string>> GetTargetYM()
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"select distinct PR_Date,
                              RIGHT('000' + CAST(YEAR(CAST(PR_Date + '-01' AS DATE)) - 1911 AS VARCHAR), 3) + '-' +
                              RIGHT('00' + CAST(MONTH(CAST(PR_Date + '-01' AS DATE)) AS VARCHAR), 2) AS PR_Date_Minguo
                              from Professional_target order by PR_Date desc";
                #endregion
                var dtResult = _adoData.ExecuteSQuery(T_SQL);
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

        #region 業績平均表
        /// <summary>
        /// 取得現有資料的年
        /// </summary>
        [HttpGet("GetTargetYYYY")]
        public ActionResult<ResultClass<string>> GetTargetYYYY()
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"select distinct LEFT(PR_Date,4) as yyyy,
                              LEFT(PR_Date,4)-1911 as yyy_Minguo
                              from Professional_target order by LEFT(PR_Date,4) desc";
                #endregion
                var dtResult = _adoData.ExecuteSQuery(T_SQL);
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
        /// 取得各區業績平均表
        /// </summary>
        [HttpGet("GetTargetAchieve")]
        public ActionResult<ResultClass<string>> GetTargetAchieve(string YYYY,string U_BC)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var Fun = new FuncHandler();
            try
            {
                var result = Fun.GetTargetAchieveList(YYYY, U_BC);
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
        /// 取得各區業績平均表報表
        /// </summary>
        [HttpGet("GetTargetAchieve_Excel")]
        public ActionResult<ResultClass<string>> GetTargetAchieve_Excel(string YYYY)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var Fun = new FuncHandler();
            var result = Fun.GetTargetAchieveList(YYYY, null);

            return Ok(resultClass);
        }
        #endregion
    }
}
