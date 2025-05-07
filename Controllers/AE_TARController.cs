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
        public ActionResult<ResultClass<string>> Pro_Target_LQuery(string? Title_name, string? PR_DATE_S)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"select PR_ID,PR_title,Li.item_D_name as PR_title_Name,PR_target,PR_Date_S,PR_Date_E
                              from Professional_target Pr
                              left join Item_list Li on Li.item_M_code='professional_title' and Li.item_D_code=Pr.PR_title
                              where 1 = 1";
                if (!string.IsNullOrEmpty(Title_name))
                {
                    T_SQL += " and Li.item_D_name = @Title_name";
                    parameters.Add(new SqlParameter("@Title_name", Title_name));
                }
                if (!string.IsNullOrEmpty(PR_DATE_S))
                {
                    T_SQL += " and PR_DATE_S = @PR_DATE_S";
                    parameters.Add(new SqlParameter("@PR_DATE_S", PR_DATE_S));
                }
                T_SQL += " order by PR_Date_S desc,Li.item_sort";
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
                var T_SQL = @"Update Professional_target set PR_target=@PR_target,PR_Date_S=@PR_Date_S,PR_Date_E=@PR_Date_E,edit_date=getdate(),
                              edit_num=@edit_num,edit_ip=@edit_ip where PR_ID=@PR_ID";
                var parameters = new List<SqlParameter>()
                {
                    new SqlParameter("@PR_target", model.amount),
                    new SqlParameter("@PR_Date_S",model.startMonth),
                    new SqlParameter("@PR_Date_E",model.endMonth),
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
                    var T_SQL_U = @"Update Person_target set PE_target=@PE_target,PE_Date_E=@PE_Date_E,edit_date=getdate(),edit_num=@edit_num 
                                    ,edit_ip=@edit_ip where PE_title=@PE_title and PE_Date_S=@PE_Date_S";
                    var parameters_u = new List<SqlParameter>()
                    {
                        new SqlParameter("@PE_target",model.amount),
                        new SqlParameter("@PE_Date_E",model.endMonth),
                        new SqlParameter("@edit_num",model.user),
                        new SqlParameter("@edit_ip",clientIp),
                        new SqlParameter("@PE_title",model.title),
                        new SqlParameter("@PE_Date_S",model.startMonth)
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
                var T_SQL = @"select item_D_code,item_D_name from Item_list where item_M_code='professional_title' 
                              and item_D_code IN 
                              (select U_PFT from User_M where Role_num in ('1008','1009') 
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
                var T_SQL_C = @"select * from Professional_target where CONVERT(date, @Date + '/01')
                                between CONVERT(date, PR_Date_S + '/01') AND CONVERT(date, PR_Date_E + '/01')";
                var parameters_c = new List<SqlParameter>()
                {
                    new SqlParameter("@Date",list[0].startMonth)
                };
                var dtReult_c = _adoData.ExecuteQuery(T_SQL_C, parameters_c);
                if (dtReult_c.Rows.Count > 0)
                {
                    resultClass.ResultCode = "400";
                    resultClass.ResultMsg = "時間內已有資料";
                    return BadRequest(resultClass);
                }
                #endregion

                var T_SQL = @"Insert into Professional_target(PR_title,PR_target,PR_Date_S,PR_Date_E,add_date,add_num,add_ip)
                              Values (@PR_title,@PR_target,@PR_Date_S,@PR_Date_E,GETDATE(),@add_num,@add_ip)";
                foreach (Pro_Target_Ins item in list)
                {
                    var parameters = new List<SqlParameter>
                    {
                        new SqlParameter("@PR_title", item.title),
                        new SqlParameter("@PR_target", item.amount),
                        new SqlParameter("@PR_Date_S", item.startMonth),
                        new SqlParameter("@PR_Date_E", item.endMonth),
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
                var T_SQL_SP = "exec UpdatePersonTargets @PR_Date_S";
                var parameters_sp = new List<SqlParameter>()
                {
                    new SqlParameter("@PR_Date_S",list[0].startMonth)
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
        public ActionResult<ResultClass<string>> Pro_Target_Clone(string PR_DATE_S,string PR_DATE_E,string User)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();
            try
            {
                ADOData _adoData = new ADOData();

                #region 檢查有無上期資料
                var parts = PR_DATE_S.Split('/');
                int year = int.Parse(parts[0]);
                int month = int.Parse(parts[1]);

                DateTime current = new DateTime(year, month, 1);
                DateTime previous = current.AddMonths(-1);

                var old_PR_DATE_E = $"{previous.Year}/{previous.Month.ToString("D2")}";

                var T_SQL_C = @"select * from Professional_target where PR_DATE_E=@PR_DATE_E";
                var parameters_c = new List<SqlParameter>()
                {
                    new SqlParameter("@PR_DATE_E",old_PR_DATE_E)
                };
                #endregion

                var dtResult_c = _adoData.ExecuteQuery(T_SQL_C, parameters_c);
                if (dtResult_c.Rows.Count > 0)
                {
                    var TargetList = dtResult_c.AsEnumerable().Select(row => new Pro_Target_Ins
                    {
                        title = row.Field<string>("PR_title"),
                        amount = row.Field<int>("PR_target"),
                        startMonth = row.Field<string>("PR_Date_S"),
                        endMonth = row.Field<string>("PR_Date_E"),
                    }).ToList();

                    var T_SQL = @"Insert into Professional_target(PR_title,PR_target,PR_Date_S,PR_Date_E,add_date,add_num,add_ip)
                              Values (@PR_title,@PR_target,@PR_Date_S,@PR_Date_E,GETDATE(),@add_num,@add_ip)";

                    foreach (Pro_Target_Ins item in TargetList)
                    {
                        var parameters = new List<SqlParameter>
                        {
                            new SqlParameter("@PR_title", item.title),
                            new SqlParameter("@PR_target", item.amount),
                            new SqlParameter("@PR_Date_S", PR_DATE_S),
                            new SqlParameter("@PR_Date_E", PR_DATE_E),
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
                    var T_SQL_SP = "exec UpdatePersonTargets @PR_Date_S";
                    var parameters_sp = new List<SqlParameter>() 
                    {
                        new SqlParameter("@PR_Date_S",PR_DATE_S)
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
        public ActionResult<ResultClass<string>> Per_Target_LQuery(string? c_name,string? PE_DATE_S)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"select PE_ID,Li.item_D_name as titleName,Um.U_name,PE_num,PE_target,PE_Date_S,PE_Date_E
                              from Person_target Pe
                              left join Item_list Li on Li.item_M_code='professional_title' and Li.item_D_code=Pe.PE_title
                              left join User_M Um on Um.U_num=Pe.PE_num
                              where 1=1";
                if (!string.IsNullOrEmpty(c_name))
                {
                    T_SQL += " and Um.U_name = @c_name";
                    parameters.Add(new SqlParameter("@c_name", c_name));
                }
                if (!string.IsNullOrEmpty(PE_DATE_S))
                {
                    T_SQL += " and PE_DATE_S = @PE_DATE_S";
                    parameters.Add(new SqlParameter("@PE_DATE_S", PE_DATE_S));
                }
                T_SQL += " order by PE_Date_S desc,Li.item_sort";
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
    }
}
