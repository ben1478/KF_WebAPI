using KF_WebAPI.BaseClass;
using KF_WebAPI.BaseClass.AE;
using KF_WebAPI.FunctionHandler;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Newtonsoft.Json;
using System.Data;
using System.Reflection;
using System.Text;

namespace KF_WebAPI.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class AE_PROController : ControllerBase
    {
        /// <summary>
        /// 提供廠商資料列表 GetManufaList
        /// </summary>
        [HttpGet("GetManufaList")]
        public ActionResult<ResultClass<string>> GetManufaList(string Name)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"select Company_name as name from Manufacturer where Company_name like @Name";
                var parameters = new List<SqlParameter> 
                {
                    new SqlParameter("@Name", "%" + Name + "%")
                };
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
        /// 採購單資料新增 Procurement_M_Ins
        /// </summary>
        [HttpPost("Procurement_M_Ins")]
        public ActionResult<ResultClass<string>> Procurement_M_Ins(Procurement_Ins model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();
            try
            {
                ADOData _adoData = new ADOData();
                #region SQL_Procurement_M
                //取PM_ID
                var T_SQL_ID = @"exec GetFormID @formtype,@tablename,@New_ID output";
                var parameters_id = new List<SqlParameter> 
                {
                    new SqlParameter("@formtype", "PO"),
                    new SqlParameter("@tablename", "Procurement_M")
                };
                SqlParameter newIdParameter = new SqlParameter("@New_ID", SqlDbType.VarChar, 20)
                {
                    Direction = ParameterDirection.Output
                };
                parameters_id.Add(newIdParameter);

                var result_id = _adoData.ExecuteQuery(T_SQL_ID, parameters_id);
                model.PM_ID = newIdParameter.Value.ToString();
                
                var T_SQL_M = @"Insert into Procurement_M (PM_ID,PM_type,PM_BC,PM_Pay_Type,PM_AppDate,PM_U_num,PM_Caption,PM_Amt,PM_Busin_Tax,PM_Tax_Amt
                    ,PM_Other,add_date,add_num,add_ip,PM_cknum)
                    Values (@PM_ID,@PM_type,@PM_BC,@PM_Pay_Type,@PM_AppDate,@PM_U_num,@PM_Caption,@PM_Amt,@PM_Busin_Tax,@PM_Tax_Amt
                    ,@PM_Other,GETDATE(),@add_num,@add_ip,@PM_cknum)";
                var parameters_m = new List<SqlParameter> 
                {
                    new SqlParameter("@PM_ID", model.PM_ID),
                    new SqlParameter("@PM_type", "PO"),
                    new SqlParameter("@PM_BC", model.PM_BC),
                    new SqlParameter("@PM_Pay_Type", model.PM_Pay_Type),
                    new SqlParameter("@PM_AppDate", DateTime.Now.ToString("yyyy/MM/dd")),
                    new SqlParameter("@PM_U_num", model.PM_U_num),
                    new SqlParameter("@PM_Caption", model.PM_Caption),
                    new SqlParameter("@PM_Amt", model.PM_Amt),
                    new SqlParameter("@PM_Busin_Tax", model.PM_Busin_Tax),
                    new SqlParameter("@PM_Tax_Amt", model.PM_Tax_Amt),
                    new SqlParameter("@PM_Other", string.IsNullOrEmpty(model.PM_Other) ? DBNull.Value : model.PM_Other),
                    new SqlParameter("@add_num", model.PM_U_num),
                    new SqlParameter("@add_ip", clientIp),
                    new SqlParameter("@PM_cknum", FuncHandler.GetCheckNum())
                };
                #endregion
                int result_m = _adoData.ExecuteNonQuery(T_SQL_M, parameters_m);
                if (result_m == 0)
                {
                    resultClass.ResultCode = "400";
                    resultClass.ResultMsg = "主檔新增失敗";
                    return BadRequest(resultClass);
                }
                else
                {
                    foreach (var item in model.PD_Ins_List) 
                    {
                        #region Procurement_D
                        var T_SQL_D = @"Insert into Procurement_D (PM_ID,PD_Pro_name,PD_Unit,PD_Count,PD_Date,PD_Univalent,PD_Amt,PD_Company_name,PD_Est_Cost
                            ,add_date,add_num,add_ip,edit_date,edit_num,edit_ip) 
                            values (@PM_ID,@PD_Pro_name,@PD_Unit,@PD_Count,@PD_Date,@PD_Univalent,@PD_Amt,@PD_Company_name,@PD_Est_Cost,GETDATE()
                            ,@add_num,@add_ip,GETDATE(),@edit_num,@edit_ip)";
                        var parameters_d = new List<SqlParameter> 
                        {
                            new SqlParameter("@PM_ID", model.PM_ID),
                            new SqlParameter("@PD_Pro_name", item.PD_Pro_name),
                            new SqlParameter("@PD_Unit", item.PD_Unit),
                            new SqlParameter("@PD_Count", item.PD_Count),
                            new SqlParameter("@PD_Date", FuncHandler.ConvertROCToGregorian(item.PD_Date)),
                            new SqlParameter("@PD_Univalent", item.PD_Univalent),
                            new SqlParameter("@PD_Amt", item.PD_Amt),
                            new SqlParameter("@PD_Company_name", item.PD_Company_name),
                            new SqlParameter("@PD_Est_Cost", item.PD_Est_Cost),
                            new SqlParameter("@add_num", model.PM_U_num),
                            new SqlParameter("@add_ip", clientIp),
                            new SqlParameter("@edit_num", model.PM_U_num),
                            new SqlParameter("@edit_ip", clientIp)
                        };
                        #endregion
                        int result_d = _adoData.ExecuteNonQuery(T_SQL_D, parameters_d);
                        if (result_d == 0)
                        {
                            resultClass.ResultCode = "400";
                            resultClass.ResultMsg = "明細檔新增失敗";
                            return BadRequest(resultClass);
                        }
                    }

                    //寫入審核流程
                    bool resultFlow = FuncHandler.AuditFlow(model.PM_U_num, model.PM_BC, "PO", model.PM_ID, clientIp);
                    if (resultFlow)
                    {
                        resultClass.ResultCode = "000";
                        resultClass.ResultMsg = "新增成功";
                        return Ok(resultClass);
                    }
                    else
                    {
                        resultClass.ResultCode = "400";
                        resultClass.ResultMsg = "審核新增失敗,請洽資訊人員";
                        return BadRequest(resultClass);
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
        /// 採購單列表查詢 Procurement_M_LQuery
        /// </summary>
        [HttpGet("Procurement_M_LQuery")]
        public ActionResult<ResultClass<string>> Procurement_M_LQuery(string User)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"select PM.PM_ID,AM.FM_Step,LI.item_D_name,LI.item_D_name AS FM_Step_SignType,PM.PM_cknum,PM.PM_Cancel
                    ,(SELECT COUNT(*) FROM ASP_UpLoad WHERE cknum = PM.PM_cknum and del_tag='0') AS PM_cknum_count 
                    from Procurement_M PM
                    INNER JOIN AuditFlow_M AM ON AM.FM_Source_ID = PM.PM_ID and AM.AF_ID = PM.PM_type
                    LEFT JOIN Item_list LI ON LI.item_D_code = AM.FM_Step_SignType AND LI.item_M_code = 'Flow_sign_type'
                    where PM.add_num = @User order by PM.add_date desc";
                var parameters = new List<SqlParameter> 
                {
                    new SqlParameter("@User", User)
                };
                #endregion
                var dtResult=_adoData.ExecuteQuery(T_SQL, parameters);

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
        /// 採購單單筆查詢 Procurement_M_SQuery
        /// </summary>
        [HttpGet("Procurement_M_SQuery")]
        public ActionResult<ResultClass<string>> Procurement_M_SQuery(string PM_ID)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"select PM_BC,PM_Pay_Type,PM_Caption,PM_Amt,PM_Busin_Tax,PM_Tax_Amt,PM_Other　from Procurement_M where PM_ID=@PM_ID ";
                var parameters = new List<SqlParameter>
                {
                    new SqlParameter("@PM_ID", PM_ID)
                };
                #endregion
                var dtResult = _adoData.ExecuteQuery(T_SQL, parameters);

                if (dtResult.Rows.Count > 0)
                {
                    var model = dtResult.AsEnumerable().Select(row => new Procurement_Res {
                        PM_BC = row.Field<string>("PM_BC"),
                        PM_Pay_Type = row.Field<string>("PM_Pay_Type"),
                        PM_Caption = row.Field<string>("PM_Caption"),
                        PM_Amt = row.Field<decimal>("PM_Amt"),
                        PM_Busin_Tax = row.Field<decimal>("PM_Busin_Tax"),
                        PM_Tax_Amt = row.Field<decimal>("PM_Tax_Amt"),
                        PM_Other = row.Field<string>("PM_Other")
                    }).FirstOrDefault();

                    #region SQL Procurement_D
                    var T_SQL_D = @"select PD_ID,PD_Pro_name,PD_Unit,PD_Count,PD_Date,PD_Univalent,PD_Amt,PD_Company_name,PD_Est_Cost
                        from Procurement_D where PM_ID=@PM_ID ";
                    var parameters_d = new List<SqlParameter> 
                    {
                        new SqlParameter("@PM_ID", PM_ID)
                    };
                    #endregion
                    model.Procurement_D = _adoData.ExecuteQuery(T_SQL_D, parameters_d).AsEnumerable().Select(row => new Procurement_D_Res {
                        PD_ID = row.Field<int>("PD_ID"),
                        PD_Pro_name = row.Field<string>("PD_Pro_name"),
                        PD_Unit = row.Field<string>("PD_Unit"),
                        PD_Count = row.Field<string>("PD_Count"),
                        PD_Date = FuncHandler.ConvertGregorianToROC(row.Field<string>("PD_Date")),
                        PD_Univalent = row.Field<decimal>("PD_Univalent"),
                        PD_Amt = row.Field<decimal>("PD_Amt"),
                        PD_Company_name = row.Field<string>("PD_Company_name"),
                        PD_Est_Cost = row.Field<decimal>("PD_Est_Cost")
                    }).ToList();


                    resultClass.ResultCode = "000";
                    resultClass.objResult = JsonConvert.SerializeObject(model);
                    return Ok(resultClass);
                }
                else
                {
                    resultClass.ResultCode = "400";
                    resultClass.objResult = JsonConvert.SerializeObject(resultClass);
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
        /// 採購單抽單 Procurement_Canel
        /// </summary>
        [HttpGet("Procurement_Canel")]
        public ActionResult<ResultClass<string>> Procurement_Canel(string PM_ID, string User)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"update Procurement_M set PM_Cancel='Y',cancel_date=GETDATE(),cancel_num=@cancel_num,candel_ip=@candel_ip where PM_ID=@PM_ID";
                var parameters = new List<SqlParameter> 
                {
                    new SqlParameter("@cancel_num", User),
                    new SqlParameter("@candel_ip", clientIp),
                    new SqlParameter("@PM_ID", PM_ID)
                };
                #endregion
                int result = _adoData.ExecuteNonQuery(T_SQL, parameters);
                if (result == 0)
                {
                    resultClass.ResultCode = "400";
                    resultClass.ResultMsg = "取消失敗";
                    return BadRequest(resultClass);
                }
                else
                {
                    resultClass.ResultCode = "000";
                    resultClass.ResultMsg = "取消成功";
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
        /// 採購單單筆修改 Procurement_Upd
        /// </summary>
        [HttpPost("Procurement_M_Upd")]
        public ActionResult<ResultClass<string>> Procurement_M_Upd(Procurement_Ins model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();
            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"update Procurement_M set PM_Pay_Type=@PM_Pay_Type,PM_Caption=@PM_Caption,PM_Amt=@PM_Amt,PM_Busin_Tax=@PM_Busin_Tax,
                    PM_Tax_Amt=@PM_Tax_Amt,PM_Other=@PM_Other,edit_num=@User,edit_date=GETDATE(),edit_ip=@IP where PM_ID=@PM_ID ";
                var parameters = new List<SqlParameter> 
                {
                    new SqlParameter("@PM_ID", model.PM_ID),
                    new SqlParameter("@PM_Pay_Type", model.PM_Pay_Type),
                    new SqlParameter("@PM_Caption", model.PM_Caption),
                    new SqlParameter("@PM_Amt", model.PM_Amt),
                    new SqlParameter("@PM_Busin_Tax", model.PM_Busin_Tax),
                    new SqlParameter("@PM_Tax_Amt", model.PM_Tax_Amt),
                    new SqlParameter("@PM_Other", string.IsNullOrEmpty(model.PM_Other) ? DBNull.Value : model.PM_Other),
                    new SqlParameter("@User", model.PM_U_num),
                    new SqlParameter("@IP", clientIp)
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
                    var T_SQL_DE = @"Delete Procurement_D where PM_ID=@PM_ID";
                    var parameters_de = new List<SqlParameter> 
                    {
                        new SqlParameter("@PM_ID", model.PM_ID)
                    };
                    _adoData.ExecuteQuery(T_SQL_DE, parameters_de);

                    foreach (var item in model.PD_Ins_List)
                    {
                        #region Procurement_D
                        var T_SQL_D = @"Insert into Procurement_D (PM_ID,PD_Pro_name,PD_Unit,PD_Count,PD_Date,PD_Univalent,PD_Amt,PD_Company_name,PD_Est_Cost
                            ,add_date,add_num,add_ip,edit_date,edit_num,edit_ip) 
                            values (@PM_ID,@PD_Pro_name,@PD_Unit,@PD_Count,@PD_Date,@PD_Univalent,@PD_Amt,@PD_Company_name,@PD_Est_Cost,GETDATE()
                            ,@add_num,@add_ip,GETDATE(),@edit_num,@edit_ip)";
                        var parameters_d = new List<SqlParameter> 
                        {
                            new SqlParameter("@PM_ID", model.PM_ID),
                            new SqlParameter("@PD_Pro_name", item.PD_Pro_name),
                            new SqlParameter("@PD_Unit", item.PD_Unit),
                            new SqlParameter("@PD_Count", item.PD_Count),
                            new SqlParameter("@PD_Date", FuncHandler.ConvertROCToGregorian(item.PD_Date)),
                            new SqlParameter("@PD_Univalent", item.PD_Univalent),
                            new SqlParameter("@PD_Amt", item.PD_Amt),
                            new SqlParameter("@PD_Company_name", item.PD_Company_name),
                            new SqlParameter("@PD_Est_Cost", item.PD_Est_Cost),
                            new SqlParameter("@add_num", model.PM_U_num),
                            new SqlParameter("@add_ip", clientIp),
                            new SqlParameter("@edit_num", model.PM_U_num),
                            new SqlParameter("@edit_ip", clientIp)
                        };
                        #endregion
                        int result_d = _adoData.ExecuteNonQuery(T_SQL_D, parameters_d);
                        if (result_d == 0)
                        {
                            resultClass.ResultCode = "400";
                            resultClass.ResultMsg = "明細檔新增失敗";
                            return BadRequest(resultClass);
                        }
                    }
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

        /// <summary>
        /// 取得請採購單列印資料 GetPrintData
        /// </summary>
        [HttpGet("GetPrintData")]
        public ActionResult<ResultClass<string>> GetPrintData(string PM_ID)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL_M = @"select LI.item_d_name as BC_Name,UM.U_name,PM.PM_AppDate,PM.PM_ID,PM.PM_Caption,Format(PM.PM_Amt,'N0') as PM_Amt,
                    Format(PM.PM_Busin_Tax,'N0') as PM_Busin_Tax,Format(PM.PM_Tax_Amt,'N0') as PM_Tax_Amt,PM.PM_cknum,
                    Case PM.PM_Pay_Type When 'GC' THEN '領現' When 'MT' THEN '匯款' END as PM_Pay_Type,PM.PM_Other,
                    FORMAT(PM.add_date,'yyyy/MM/dd HH:mm') as add_date
                    from Procurement_M PM
                    left join User_M UM on UM.U_num = PM.PM_U_num
                    left join Item_list LI on LI.item_D_code = PM.PM_BC
                    where PM.PM_ID=@PM_ID";
                var parameters_m = new List<SqlParameter> 
                {
                    new SqlParameter("@PM_ID", PM_ID)
                };
                #endregion
                var resultModel = _adoData.ExecuteQuery(T_SQL_M, parameters_m).AsEnumerable().Select(row => new Proc_Print
                {
                    BC_Name = row.Field<string>("BC_Name"),
                    U_name = row.Field<string>("U_name"),
                    PM_AppDate = row.Field<string>("PM_AppDate"),
                    PM_ID = row.Field<string>("PM_ID"),
                    PM_Caption = row.Field<string>("PM_Caption"),
                    PM_Amt = row.Field<string>("PM_Amt"),
                    PM_Busin_Tax = row.Field<string>("PM_Busin_Tax"),
                    PM_Tax_Amt = row.Field<string>("PM_Tax_Amt"),
                    PM_cknum = row.Field<string>("PM_cknum"),
                    PM_Pay_Type = row.Field<string>("PM_Pay_Type"),
                    PM_Other = row.Field<string>("PM_Other"),
                    add_date = row.Field<string>("add_date")
                }).FirstOrDefault();

                #region Proc_Print_Deatil
                var T_SQL_DT = @"select PD_Pro_name,PD_Unit,PD_Count,PD_Date,Format(PD_Univalent,'N0') as PD_Univalent,
                    Format(PD_Amt,'N0') as PD_Amt,PD_Company_name,Format(PD_Est_Cost,'N0') as PD_Est_Cost 
                    from Procurement_D where PM_ID=@PM_ID";
                var parameters_dt = new List<SqlParameter> 
                {
                    new SqlParameter("@PM_ID", PM_ID)
                };
                resultModel.PT_Deatil_List = _adoData.ExecuteQuery(T_SQL_DT, parameters_dt).AsEnumerable().Select(row => new Proc_Print_Deatil
                {
                    PD_Pro_name = row.Field<string>("PD_Pro_name"),
                    PD_Unit = row.Field<string>("PD_Unit"),
                    PD_Count = row.Field<string>("PD_Count"),
                    PD_Date = row.Field<string>("PD_Date"),
                    PD_Univalent = row.Field<string>("PD_Univalent"),
                    PD_Amt = row.Field<string>("PD_Amt"),
                    PD_Company_name = row.Field<string>("PD_Company_name"),
                    PD_Est_Cost = row.Field<string>("PD_Est_Cost")
                }).ToList();
                #endregion

                #region Proc_Print_File
                var T_SQL_FI = @"select upload_name_show from ASP_UpLoad where cknum=@cknum";
                var parameters_fi = new List<SqlParameter> 
                {
                    new SqlParameter("@cknum", resultModel.PM_cknum)
                };
                resultModel.PT_File_List = _adoData.ExecuteQuery(T_SQL_FI, parameters_fi).AsEnumerable().Select(row => new Proc_Print_File
                {
                    upload_name_show = row.Field<string>("upload_name_show"),
                }).ToList();
                #endregion

                #region Proc_Print_Flow
                var T_SQL_FW = @"select AD.FD_Sign_Countersign,AD.FD_Step,UM.U_name,
                    FORMAT(AD.FD_Step_date,'yyyy/MM/dd HH:mm') as FD_Step_date,AD.FD_Step_desc 
                    from AuditFlow_D AD
                    inner join User_M UM on UM.U_num = AD.FD_Step_num
                    where FM_Source_ID=@PM_ID order by AD.FD_Step";
                var parameters_fw = new List<SqlParameter> 
                {
                    new SqlParameter("@PM_ID", PM_ID)
                };
                resultModel.PT_Flow_List = _adoData.ExecuteQuery(T_SQL_FW, parameters_fw).AsEnumerable().Select(row => new Proc_Print_Flow
                {
                    FD_Sign_Countersign = row.Field<string>("FD_Sign_Countersign"),
                    FD_Step = row.Field<string>("FD_Step"),
                    U_name = row.Field<string>("U_name"),
                    FD_Step_date = row.Field<string>("FD_Step_date"),
                    FD_Step_desc = row.Field<string>("FD_Step_desc")
                }).ToList();
                #endregion

                resultClass.ResultCode = "000";
                resultClass.objResult = JsonConvert.SerializeObject(resultModel);

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
        /// 財務表單報表 PT_Rpt
        /// </summary>
        [HttpPost("PT_Rpt")]
        public ActionResult<ResultClass<string>> PT_Rpt (PT_Rpt_req model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter> 
                {
                    new SqlParameter("@Date_S", FuncHandler.ConvertROCToGregorian(model.Date_S)),
                    new SqlParameter("@Date_E", FuncHandler.ConvertROCToGregorian(model.Date_E))
                };
                var T_SQL = @"select VP_ID as Form_ID,VP_type as strType,LI.item_D_name as BC_Name,UM.U_name as U_Name
                    ,VP_Summary as str_Name,Format(VP_Total_Money,'N0') as str_Amt,VP.add_date as add_date
                    from InvPrepay_M VP
                    left join Item_list LI on LI.item_M_code='branch_company' and LI.item_D_code=VP_BC
                    left join User_M UM on UM.U_num=VP.add_num
                    where VP.add_date BETWEEN @Date_S AND @Date_E";
                if (!string.IsNullOrEmpty(model.Type))
                {
                    T_SQL += " and VP_type=@VP_type";
                    parameters.Add(new SqlParameter("@VP_type", model.Type));
                }
                if (string.IsNullOrEmpty(model.Type) || model.Type=="PO")
                {
                    T_SQL += @" union
                    select PM_ID as Form_ID,'PO' as strType,LI.item_D_name as BC_Name,UM.U_name as U_Name
                    ,PM_Caption as str_Name,Format(PM_Amt,'N0') as str_Amt,PM.add_date as add_date
                    from Procurement_M PM
                    left join Item_list LI on LI.item_M_code='branch_company' and LI.item_D_code=PM_BC
                    left join User_M UM on UM.U_num=PM.add_num
                    where PM.add_date BETWEEN @Date_S AND @Date_E";
                }
                #endregion
                var dtResult = _adoData.ExecuteQuery(T_SQL,parameters).AsEnumerable()
                    .OrderByDescending(row=> row.Field<DateTime>("add_date")).CopyToDataTable(); ;
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
        /// 採購單報表Excel下載 PT_Rpt_Excel
        /// </summary>
        [HttpPost("PT_Rpt_Excel")]
        public IActionResult PT_Rpt_Excel(PT_Rpt_req model)
        {
            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>
                {
                    new SqlParameter("@Date_S", FuncHandler.ConvertROCToGregorian(model.Date_S)),
                    new SqlParameter("@Date_E", FuncHandler.ConvertROCToGregorian(model.Date_E))
                };
                var queryBuilder = new StringBuilder();
                queryBuilder.AppendLine(@"select VP_ID as Form_ID,VP_type as strType,LI.item_D_name as BC_Name,UM.U_name as U_Name,
                                          VP_Summary as str_Name,Format(VP_Total_Money,'N0') as str_Amt,VP.add_date AS add_date
                                          from InvPrepay_M VP
                                          left join Item_list LI on LI.item_M_code='branch_company' and LI.item_D_code=VP_BC
                                          left join User_M UM on UM.U_num=VP.add_num
                                          where VP.add_date BETWEEN @Date_S AND @Date_E");
                if (!string.IsNullOrEmpty(model.Type))
                {
                    queryBuilder.AppendLine("  AND VP_type = @VP_type");
                    parameters.Add(new SqlParameter("@VP_type", model.Type));
                }
                if (string.IsNullOrEmpty(model.Type) || model.Type == "PO")
                {
                    queryBuilder.AppendLine("UNION");
                    queryBuilder.AppendLine(@"SELECT PM_ID AS Form_ID,'PO' AS strType,LI.item_D_name AS BC_Name,UM.U_name AS U_Name,
                                              PM_Caption AS str_Name,FORMAT(PM_Amt, 'N0') AS str_Amt,PM.add_date AS add_date
                                              FROM Procurement_M PM
                                              LEFT JOIN Item_list LI ON LI.item_M_code = 'branch_company' AND LI.item_D_code = PM_BC
                                              LEFT JOIN User_M UM ON UM.U_num = PM.add_num
                                              WHERE PM.add_date BETWEEN @Date_S AND @Date_E");
                }
                #endregion
                var dtResult = _adoData.ExecuteQuery(queryBuilder.ToString(), parameters);
                if (dtResult.Rows.Count > 0)
                {
                    var excelList = dtResult.AsEnumerable().Select(row => new PT_Excel
                    {
                        Form_ID = row.Field<string>("Form_ID"),
                        strType = GetStrType(row),
                        BC_Name = row.Field<string>("BC_Name"),
                        U_Name = row.Field<string>("U_Name"),
                        str_Name = row.Field<string>("str_Name"),
                        str_Amt = row.Field<string>("str_Amt")
                    }).ToList();

                    var Excel_Headers = new Dictionary<string, string>
                    {
                        { "Form_ID","單號" },
                        { "strType", "類型" },
                        { "BC_Name", "申請部門" },
                        { "U_Name", "申請人" },
                        { "str_Name", "費用類別" },
                        { "str_Amt","總價" }
                    };

                    var fileBytes = FuncHandler.ExportToExcel(excelList, Excel_Headers);
                    var fileName = "採購單資料" + DateTime.Now.ToString("yyyyMMddHHmm") + ".xlsx";
                    return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
                }
                else
                {
                    return NotFound();
                }
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex);
            }
        }

        private string GetStrType(DataRow row)
        {
            string typeStr = row.Field<string>("strType");
            return typeStr == "PO" ? "請採購" :
                   typeStr == "PA" ? "請款" :
                   typeStr == "PP" ? "預支" :
                   typeStr == "PS" ? "沖銷預支":"-";
        }

    }
}
