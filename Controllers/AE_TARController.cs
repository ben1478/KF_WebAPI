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
using System.Data;
using System.Data.SqlTypes;
using System.Reflection;
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
        public ActionResult<ResultClass<string>> Pro_Target_LQuery()
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"select PR_ID,Li.item_D_name as PR_title,PR_target,PR_Date_S,PR_Date_E
                              from Professional_target Pr
                              left join Item_list Li on Li.item_M_code='professional_title' and Li.item_D_code=Pr.PR_title";
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
        /// 業務責任額列表
        /// </summary>
        [HttpGet("Per_Target_LQuery")]
        public ActionResult<ResultClass<string>> Per_Target_LQuery()
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"select PE_ID,Li.item_D_name as PE_title,PE_num,PE_target,PE_Date_S,PE_Date_E
                              from Person_target Pe
                              left join Item_list Li on Li.item_M_code='professional_title' and Li.item_D_code=Pe.PE_title";
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
                              and item_D_code IN (select U_PFT from User_M where Role_num in ('1008','1009') and U_leave_date is null)";
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

            try
            {
                //TODO 新增
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"";
                #endregion
                //TODO 順便呼叫SP完成業務個人責任額新增
                return Ok();
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
