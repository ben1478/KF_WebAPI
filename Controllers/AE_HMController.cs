using KF_WebAPI.BaseClass.AE;
using KF_WebAPI.BaseClass;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using KF_WebAPI.FunctionHandler;
using Microsoft.Data.SqlClient;
using System.Diagnostics.Eventing.Reader;
using System.Data;
using KF_WebAPI.BaseClass.Winton;
using Newtonsoft.Json;
using System.Reflection;
using System.Collections.Generic;
using KF_WebAPI.DataLogic;
using System.Xml.Linq;

namespace KF_WebAPI.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class AE_HMController : Controller
    {
        private AEData _AEData = new();

        private AE_HM _HM = new AE_HM();

        /// <summary>
        /// 汽車進件API HouseApplyCar_Ins / House_apply_addCarDB.asp
        /// </summary>
        [HttpPost("HouseApplyCar_Ins")]
        public ActionResult<ResultClass<string>> HouseApplyCar_Ins(HouseApplyCar_req model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

            try
            {
                ADOData _adoData = new ADOData();
                SqlParameter pNull(string name, string value) => new SqlParameter(name, string.IsNullOrWhiteSpace(value) ? (object)DBNull.Value : value);
                #region 檢查重復
                if (model.HouseApplyChk == "N")
                {
                    var T_SQL_CK = @"select * from House_apply where del_tag = '0' and plan_num = @plan_num and datediff(day,add_date,getdate()) < 90
                                     and (CS_name <> '' and CS_name=@CS_name or (CS_MTEL1 <>'' and (CS_MTEL1=@CS_MTEL1 or CS_MTEL1=CS_MTEL2)) 
                                     or (CS_MTEL2<>'' and (CS_MTEL2=@CS_MTEL1 or CS_MTEL2=@CS_MTEL2)))";
                    var parameters_ck = new List<SqlParameter>
                    {
                        new SqlParameter("@plan_num", model.tbInfo.add_num),
                        new SqlParameter("@CS_name", model.CS_name),
                        pNull("@CS_MTEL1", model.CS_MTEL1),
                        pNull("@CS_MTEL2", model.CS_MTEL2)
                    };
                    var dtResult_ck = _adoData.ExecuteQuery(T_SQL_CK, parameters_ck);
                    if (dtResult_ck.Rows.Count != 0)
                    {
                        resultClass.ResultCode = "400";
                        resultClass.ResultMsg = "客戶重複";
                        return Ok(resultClass);
                    }
                }

                #endregion
                #region 確認介紹人
                if (!string.IsNullOrEmpty(model.CS_introducer_PID) && model.CS_introducer_PID != "無"){
                    var T_SQL_in = @"select Introducer_name from Introducer_Comm where Introducer_PID = @Introducer_PID";
                    var parameters_in = new List<SqlParameter>
                    {
                        new SqlParameter("@Introducer_PID", model.CS_introducer_PID)
                    };
                    model.CS_introducer = _adoData.ExecuteQuery(T_SQL_in, parameters_in).AsEnumerable().Select(row => new
                    {
                        CS_introducer = row.Field<string>("Introducer_name")
                    }).FirstOrDefault().CS_introducer;
                }
                #endregion
                #region 新增House_apply
                var T_SQL_apply = @"Insert into House_apply (HA_cknum,CS_name,CS_sex,CS_PID,CS_birth_year,CS_MTEL1,CS_MTEL2,CS_introducer,CS_introducer_PID,CS_note
                                    ,plan_num,plan_date,plan_type,plan_type_date,add_date,add_num,add_ip,CS_birthday,CS_register_address,CS_car,CS_carBrand
                                    ,CS_carNumber,CS_carModel,CS_carManufacture,CS_carDisplacement,CS_EngineNo,CS_company_name,CS_company_address,CS_company_tel
                                    ,CS_job_kind,CS_job_title,CS_job_years,CS_income_way,CS_rental,CS_income_everymonth,CS_license,CS_EMAIL) 
                                    Values (@HA_cknum,@CS_name,@CS_sex,@CS_PID,@CS_birth_year,@CS_MTEL1,@CS_MTEL2,@CS_introducer,@CS_introducer_PID,@CS_note
                                    ,@plan_num,@plan_date,'plan_T001',getdate(),getdate(),@add_num,@add_ip,@CS_birthday,@CS_register_address,@CS_car,@CS_carBrand
                                    ,@CS_carNumber,@CS_carModel,@CS_carManufacture,@CS_carDisplacement,@CS_EngineNo,@CS_company_name,@CS_company_address,@CS_company_tel
                                    ,@CS_job_kind,@CS_job_title,@CS_job_years,@CS_income_way,@CS_rental,@CS_income_everymonth,@CS_license,@CS_EMAIL);
                                    SELECT SCOPE_IDENTITY();";
                var parameters_apply = new List<SqlParameter>
                {
                    new SqlParameter("@HA_cknum", FuncHandler.GetCheckNum()),
                    new SqlParameter("@CS_name", model.CS_name),
                    pNull("@CS_sex", model.CS_sex),
                    new SqlParameter("@CS_PID", model.CS_PID),
                    pNull("@CS_birth_year", model.CS_birth_year),
                    pNull("@CS_MTEL1", model.CS_MTEL1),
                    pNull("@CS_MTEL2", model.CS_MTEL2),
                    pNull("@CS_introducer", model.CS_introducer),
                    pNull("@CS_introducer_PID", model.CS_introducer_PID),
                    pNull("@CS_note", model.CS_note),
                    pNull("@plan_num", model.tbInfo.add_num),
                    pNull("@plan_date", model.plan_date),
                    pNull("@add_num", model.tbInfo.add_num),
                    pNull("@add_ip", clientIp),
                    pNull("@CS_birthday", model.CS_birthday),
                    pNull("@CS_register_address", model.CS_register_address),
                    pNull("@CS_car", model.CS_car),
                    pNull("@CS_carBrand", model.CS_carBrand),
                    pNull("@CS_carNumber", model.CS_carNumber),
                    pNull("@CS_carModel", model.CS_carModel),
                    pNull("@CS_carManufacture", model.CS_carManufacture),
                    pNull("@CS_carDisplacement", model.CS_carDisplacement),
                    pNull("@CS_EngineNo", model.CS_EngineNo),
                    pNull("@CS_company_name", model.CS_company_name),
                    pNull("@CS_company_address", model.CS_company_address),
                    pNull("@CS_company_tel", model.CS_company_tel),
                    pNull("@CS_job_kind", model.CS_job_kind),
                    pNull("@CS_job_title", model.CS_job_title),
                    pNull("@CS_job_years", model.CS_job_years),
                    pNull("@CS_income_way", model.CS_income_way),
                    pNull("@CS_rental", model.CS_rental),
                    pNull("@CS_income_everymonth", model.CS_income_everymonth),
                    pNull("@CS_license", model.CS_license),
                    pNull("@CS_EMAIL", model.CS_EMAIL)
                };
                int newHAid = Convert.ToInt32(_adoData.ExecuteScalar(T_SQL_apply, parameters_apply));
                #endregion
                #region 新增House_pre
                var T_SOL_pre = @"Insert into House_pre (HP_cknum,HA_id,pre_apply_name,pre_apply_date,pre_address,pre_process_type,add_date,add_num,add_ip) 
                                  Values (@HP_cknum,@HA_id,@pre_apply_name,getdate(),@pre_address,'PRCT0005',getdate(),@add_num,@add_ip);
                                  SELECT SCOPE_IDENTITY();";
                var parameters_pre = new List<SqlParameter>
                {
                    new SqlParameter("@HP_cknum", FuncHandler.GetCheckNum()),
                    new SqlParameter("@HA_id", newHAid),
                    new SqlParameter("@pre_apply_name", model.CS_name),
                    pNull("@pre_address", model.pre_address),
                    pNull("@add_num", model.tbInfo.add_num),
                    pNull("@add_ip", clientIp)
                };
                int newHPID = Convert.ToInt32(_adoData.ExecuteScalar(T_SOL_pre, parameters_pre));
                #endregion
                #region 新增House_pre_appraise
                var T_SQL_app = @"Insert into House_pre_appraise (HA_id,HP_id,build_area_price,land_price,building_price,total_price,rise_tax,net_value
                                  ,appraise_company,add_date,add_num,add_ip) 
                                  Values (@HA_id,@HP_id,100,100,100,100,100,100,'APCOM099',getdate(),@add_num,@add_ip)";
                var parameters_app = new List<SqlParameter>
                {
                    new SqlParameter("@HA_id", newHAid),
                    new SqlParameter("@HP_id", newHPID),
                    pNull("@add_num", model.tbInfo.add_num),
                    pNull("@add_ip", clientIp)
                };
                _adoData.ExecuteQuery(T_SQL_app, parameters_app);
                #endregion
                #region 新增House_pre_project
                var T_SQL_ppj = @"Insert into House_pre_project (HA_id,appraise_company,project_apply_amount,project_title,project_type,project_type_date
                                  ,project_leader_num,project_leader_sign,project_leader_agree,sendcase_handle_num,add_date,add_num,add_ip) 
                                  Values (@HA_id,'APCOM099',@project_apply_amount,'PJ00048','PRO_T002',getdate(),'K0120','N','',@sendcase_handle_num
                                  ,getdate(),@add_num,@add_ip);
                                  SELECT SCOPE_IDENTITY();";
                var parameters_ppj = new List<SqlParameter>
                {
                    new SqlParameter("@HA_id", newHAid),
                    new SqlParameter("@project_apply_amount", model.project_apply_amount),
                    pNull("@sendcase_handle_num", model.tbInfo.add_num),
                    pNull("@add_num", model.tbInfo.add_num),
                    pNull("@add_ip", clientIp)
                };
                int HPRID = Convert.ToInt32(_adoData.ExecuteScalar(T_SQL_ppj, parameters_ppj));
                #endregion
                #region 新增House_pre_pawn
                var T_SQL_prp = @"Insert into House_pre_pawn (HA_id,HP_id,appraise_company,pawn_type,pawn_type_date,pawn_address,pawn_building_kind
                                  ,pawn_storey,pawn_land_size,pawn_building_size,pawn_parking_kind,pawn_house_years,pawn_keep_time,add_date,add_num,add_ip) 
                                  Values (@HA_id,@HP_id,'APCOM099','Y',getdate(),'','','','','','','','',getdate(),@add_num,@add_ip)";
                var parameters_prp = new List<SqlParameter>
                {
                    new SqlParameter("@HA_id", newHAid),
                    new SqlParameter("@HP_id", newHPID),
                    pNull("@add_num", model.tbInfo.add_num),
                    pNull("@add_ip", clientIp)
                };
                _adoData.ExecuteQuery(T_SQL_prp, parameters_prp);
                #endregion
                resultClass.ResultCode = "000";
                resultClass.ResultMsg = "變更成功";
                return Ok(resultClass);
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass);
            }
        }

        #region 客戶繳款回報
        /// <summary>
        /// 業務或業助抓取待沖銷分期資料(queryType 1.查看全部資訊  2.查看分區資訊  3.查看單一業務資訊)
        /// </summary>
        [HttpPost("Receivable_Pay_LQuery")]
        public ActionResult<ResultClass<string>> Receivable_Pay_LQuery(Receivable_Pay_req model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            FuncHandler _Fun = new FuncHandler();

            try
            {
                ADOData _adoData = new ADOData();
                var parameters = new List<SqlParameter>
                {
                    new SqlParameter("@user",model.tbInfo.add_num)
                };
                #region T_SQL
                var T_SQL = @"select H.*,D.*,M.amount_per_month,M.HS_id,amount_total,month_total
                              ,(select COUNT(*) from AE_Files where AE_Files.KeyID = 'P' + cast(D.RCD_id as varchar)) as FileCount from ";
                if (model.queryType == 1)
                {
                    T_SQL += @" ( select HA_id,CS_Name,CS_PID,plan_num FROM House_apply h) H";
                }
                else if (model.queryType == 2)
                {
                    T_SQL += @" ( select HA_id,CS_Name,CS_PID,plan_num FROM House_apply h
	                              JOIN User_M u1 ON u1.U_num = @user
	                              JOIN User_M u2 ON u1.U_BC = u2.U_BC
	                              WHERE h.plan_num = u2.U_num ) H";
                }
                else
                {
                    T_SQL += @" ( select HA_id,CS_Name,CS_PID,plan_num FROM House_apply
	                              JOIN User_M ON House_apply.plan_num = User_M.U_num WHERE User_M.U_num = @user ) H ";
                }
                T_SQL += @" left join House_sendcase S on H.HA_id=S.HA_id and get_amount_type = 'GTAT002'                                          
                            left join Receivable_M M on H.HA_id=M.HA_id 
                            left join ( select * from Receivable_D where cast(RCM_id as varchar)+'-'+ cast( (RC_count) as varchar) in 
                            (/*抓出最近一期沒繳款的資料*/
                            select cast(RCM_id as varchar)+'-'+ cast( min(RC_count) as varchar)RC_count 
                            from Receivable_D where check_pay_type ='N' and bad_debt_type = 'N'";
                if (!string.IsNullOrEmpty(model.RC_Date_S) && !string.IsNullOrEmpty(model.RC_Date_E))
                {
                    T_SQL += @" and RC_date >= @RC_Date_S and RC_date <= @RC_Date_E ";
                }
                T_SQL += @"group by RCM_id
                           union all 
                           select cast(RCM_id as varchar)+'-'+ cast( min(RC_count) +1 as varchar)RC_count 
                           from Receivable_D where check_pay_type ='N' and bad_debt_type = 'N'";
                if (!string.IsNullOrEmpty(model.RC_Date_S) && !string.IsNullOrEmpty(model.RC_Date_E))
                {
                    T_SQL += @" and RC_date >= @RC_Date_S and RC_date <= @RC_Date_E ";
                }
                T_SQL += @"group by RCM_id ) ) D on M.RCM_id=D.RCM_id
                           where RCM_note not like '%清償%' and D.check_pay_type ='N' and M.del_tag='0' and D.del_tag='0' 
                           and S.del_tag='0' and fund_company='FDCOM003' and not exists (select 1 from ClientPayback Ck where Ck.RCD_id = D.RCD_id and CP_Win_CK <> 'D')";

                if (!string.IsNullOrEmpty(model.RC_Date_S) && !string.IsNullOrEmpty(model.RC_Date_E))
                {
                    model.RC_Date_S = FuncHandler.ConvertROCToGregorian(model.RC_Date_S);
                    model.RC_Date_E = FuncHandler.ConvertROCToGregorian(model.RC_Date_E);
                    parameters.Add(new SqlParameter("@RC_Date_S", model.RC_Date_S));
                    parameters.Add(new SqlParameter("@RC_Date_E", model.RC_Date_E));
                }
                if (!string.IsNullOrEmpty(model.CS_Name))
                {
                    T_SQL += " and CS_name = @CS_Name";
                    parameters.Add(new SqlParameter("@CS_name", model.CS_Name));
                }
                if (!string.IsNullOrEmpty(model.CS_PID))
                {
                    T_SQL += " and CS_PID = @CS_PID";
                    parameters.Add(new SqlParameter("@CS_PID", model.CS_PID));
                }
                if (!string.IsNullOrEmpty(model.plan_num))
                {
                    T_SQL += " and plan_num = @plan_num";
                    parameters.Add(new SqlParameter("@plan_num", model.plan_num));
                }
                #endregion

                var result = _adoData.ExecuteQuery(T_SQL, parameters).AsEnumerable().Select(row => new Receivable_Pay_res
                {
                    HS_id = row.Field<decimal>("HS_id"),
                    RCD_id = row.Field<decimal>("RCD_id"),
                    CS_name = _Fun.DeCodeBNWords(row.Field<string>("CS_name")),
                    CS_PID = row.Field<string>("CS_PID"),
                    RC_count = row.Field<int>("RC_count"),
                    roc_RC_date = FuncHandler.ConvertGregorianToROC(row.Field<DateTime>("RC_date").ToString("yyyy/MM/dd")),
                    amount_per_month = row.Field<decimal>("amount_per_month"),
                    interest = row.Field<decimal>("interest"),
                    Rmoney = row.Field<decimal>("Rmoney"),
                    HFees = 20,
                    Ex_RemainingPrincipal = row.Field<decimal>("Ex_RemainingPrincipal"),
                    amount_total = row.Field<decimal>("amount_total"),
                    month_total = row.Field<int>("month_total"),
                    FileCount = row.Field<int>("FileCount"),
                    CP_bus_remark = row.Field<string?>("RC_note"),
                    CP_Pay_Amt = row.Field<decimal?>("RecPayAmt")
                }).ToList();

                resultClass.ResultCode = "000";
                resultClass.ResultMsg = "成功";
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
        /// 業務新增客戶繳款資料
        /// </summary>
        [HttpPost("Client_Pay_Ins")]
        public ActionResult<ResultClass<string>> Client_Pay_Ins([FromBody] List<Receivable_Pay_Ins> List)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();

            try
            {
                ADOData _adoData = new ADOData();

                foreach (var item in List)
                {
                    var T_SQL = "";
                    var parameters = new List<SqlParameter>();
                    //判定已繳款金額是否足額
                    if (item.CP_Pay_Amt >= item.amount_per_month)
                    {
                        T_SQL = @"Insert into ClientPayback (RCD_id,CP_pay_date,CP_account_last,CP_Pay_Amt,CP_bus_remark,add_date,add_num) 
                                  Values (@RCD_id,@CP_pay_date,@CP_account_last,@CP_Pay_Amt,@CP_bus_remark,getdate(),@add_num)";
                        parameters.Add(new SqlParameter("@RCD_id", item.RCD_id));
                        parameters.Add(new SqlParameter("@CP_pay_date", item.CP_pay_date));
                        parameters.Add(new SqlParameter("@CP_account_last", item.CP_account_last));
                        parameters.Add(new SqlParameter("@CP_Pay_Amt", item.CP_Pay_Amt));
                        parameters.Add(new SqlParameter("@CP_bus_remark", item.CP_bus_remark));
                        parameters.Add(new SqlParameter("@add_num", item.User));
                    }
                    else
                    {
                        //異動部分Receivable_D
                        T_SQL += @"Update Receivable_D set RecPayAmt=@CP_Pay_Amt,RC_note=@CP_bus_remark where RCD_id = @RCD_id";
                        parameters.Add(new SqlParameter("@CP_Pay_Amt", item.CP_Pay_Amt));
                        parameters.Add(new SqlParameter("@CP_bus_remark", item.CP_bus_remark));
                        parameters.Add(new SqlParameter("@RCD_id", item.RCD_id));
                    }
                    _adoData.ExecuteNonQuery(T_SQL, parameters);

                    #region 寫入LogTable
                    var logTable = new LogTable();
                    logTable.TableNA = "ClientPayback";
                    logTable.KeyVal = item.RCD_id.ToString();
                    logTable.ColumnNA = "CP_pay_date" + "," + "CP_Pay_Amt";
                    logTable.ColumnVal = item.CP_pay_date.ToString("yyyy-MM-dd") + "," + item.CP_Pay_Amt.ToString();
                    logTable.Remark = item.CP_bus_remark;
                    logTable.LogID = item.User;
                    _AEData.InsLogTable(logTable);
                    #endregion

                }

                resultClass.ResultCode = "000";
                resultClass.ResultMsg = "儲存成功";
                return Ok(resultClass);
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass);
            }
        }

        // <summary>
        // 財務抓取待開發票的客戶資料
        // </summary>
        [HttpPost("Client_Pay_LQuery")]
        public ActionResult<ResultClass<string>> Client_Pay_LQuery(Client_Pay_req model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var clientIp = HttpContext.Connection.RemoteIpAddress.ToString();
            FuncHandler _Fun = new FuncHandler();

            try
            {
                ADOData _adoData = new ADOData();
                var parameters = new List<SqlParameter>();
                #region SQL
                var T_SQL = @"select *,(select COUNT(*) from AE_Files where AE_Files.KeyID = 'P' + cast(Cp.RCD_id as varchar)) as FileCount,Um.U_name as U_Name
                              ,CASE WHEN P.project_title IN ('PJ00046', 'PJ00047') THEN '機車貸' WHEN P.project_title IN ('PJ00048') THEN '汽車貸' ELSE '房貸' END as CaseType
                              from ClientPayback Cp
                              inner join Receivable_D Rd ON Rd.RCD_id = Cp.RCD_id
                              inner join Receivable_M Rm ON Rm.RCM_id = Rd.RCM_id
                              inner join House_apply HA ON HA.HA_id = RM.HA_id
                              left join User_M Um ON Um.U_num = Cp.add_num
                              LEFT JOIN House_sendcase S ON Rm.HS_id=S.HS_id
                              LEFT JOIN House_pre_project P ON S.HP_project_id = P.HP_project_id
                              where RM.del_tag = 0  and cancel_type <> 'Y' and bad_debt_type = 'N'";
                if (!string.IsNullOrEmpty(model.CP_WIN_CK))
                {
                    T_SQL += @" and CP_Win_CK = @CP_WIN_CK ";
                    parameters.Add(new SqlParameter("@CP_WIN_CK", model.CP_WIN_CK));
                }

                if (!string.IsNullOrEmpty(model.str_CP_pay_date))
                {
                    var CP_pay_date = FuncHandler.ConvertROCToGregorian(model.str_CP_pay_date);
                    T_SQL += @" and CP_pay_date = @CP_pay_date ";
                    parameters.Add(new SqlParameter("@CP_pay_date", CP_pay_date));
                }
                #endregion
                var result = _adoData.ExecuteQuery(T_SQL, parameters).AsEnumerable().Select(row => new PaySelf_Win_Inv
                {
                    HS_id = row.Field<decimal>("HS_id"),
                    RCD_id = row.Field<decimal>("RCD_id"),
                    CS_name = _Fun.DeCodeBNWords(row.Field<string>("CS_name")),
                    CS_PID = row.Field<string>("CS_PID"),
                    RC_count = row.Field<int>("RC_count"),
                    roc_RC_date = FuncHandler.ConvertGregorianToROC(row.Field<DateTime>("RC_date").ToString("yyyy/MM/dd")),
                    amount_per_month = row.Field<decimal>("amount_per_month"),
                    interest = row.Field<decimal>("interest"),
                    Rmoney = row.Field<decimal>("Rmoney"),
                    HFees = 20,
                    Ex_RemainingPrincipal = row.Field<decimal>("Ex_RemainingPrincipal"),
                    amount_total = row.Field<decimal>("amount_total"),
                    month_total = row.Field<int>("month_total"),
                    RecPayDate = row.Field<DateTime>("CP_pay_date"),
                    CP_account_last = row.Field<string>("CP_account_last"),
                    CP_bus_remark = row.Field<string>("CP_bus_remark"),
                    CP_Pay_Amt = row.Field<decimal?>("CP_Pay_Amt"),
                    FileCount = row.Field<int>("FileCount"),
                    str_Pay_Date = FuncHandler.ConvertGregorianToROC(row.Field<DateTime>("CP_pay_date").ToString("yyyy/MM/dd")),
                    U_Name = row.Field<string>("U_Name"),
                    CaseType = row.Field<string>("CaseType")
                }).ToList(); ;
                resultClass.ResultCode = "000";
                resultClass.ResultMsg = "變更成功";
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

        // <summary>
        // 財務刪除待開發票的客戶資料
        // </summary>
        [HttpPost("Client_Pay_Del")]
        public ActionResult<ResultClass<string>> Client_Pay_Del([FromBody] int[] ids, string user)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();

                if (ids == null || ids.Length == 0)
                {
                    resultClass.ResultCode = "400";
                    resultClass.ResultMsg = "沒有要刪除的資料";
                    return Ok(resultClass);
                }

                var paramNames = ids.Select((id, i) => $"@id{i}").ToArray();

                var T_SQL = $@"
                     UPDATE ClientPayback
                     SET CP_Win_CK = 'D',
                         del_date = GETDATE(),
                         del_num = @user
                     WHERE RCD_id IN ({string.Join(",", paramNames)})
                 ";

                List<SqlParameter> parameters = new List<SqlParameter>();

                parameters.Add(new SqlParameter("@user", user));

                for (int i = 0; i < ids.Length; i++)
                {
                    parameters.Add(new SqlParameter(paramNames[i], ids[i]));
                }

                _adoData.ExecuteQuery(T_SQL, parameters);

                resultClass.ResultCode = "000";
                resultClass.ResultMsg = "刪除成功";
                return Ok(resultClass);
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                resultClass.ResultMsg = ex.Message;
                return StatusCode(500, resultClass);
            }
        }
        #endregion

        #region 客訴資料維護
        [HttpPost("Complaint_LQuery")]
        public ActionResult<ResultClass<string>> Complaint_LQuery(Complaint_M_req model)
        {
            ResultClass<string> resultClass = new();

            try
            {
                resultClass = _HM.Complaint_LQuery(model);
                resultClass.ResultCode = "000";
                return Ok(resultClass);
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass);
            }

        }

        [HttpPost("Complaint_Ins")]
        public ActionResult<ResultClass<string>> Complaint_Ins(Complaint_M model)
        {
            ResultClass<string> resultClass = new ();

            try
            {
                var reslt = _HM.Complaint_Ins(model);
                if(reslt > 0)
                {
                    resultClass.ResultCode = "000";
                    resultClass.ResultMsg = "儲存成功";
                }
                else
                {
                    resultClass.ResultCode = "400";
                    resultClass.ResultMsg = "儲存失敗";
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

        [HttpGet("Complaint_SQuery")]
        public ActionResult<ResultClass<string>> Complaint_SQuery(string Comp_Id)
        {
            ResultClass<string> resultClass = new();

            try
            {
                resultClass = _HM.Complaint_SQuery(Comp_Id);
                resultClass.ResultCode = "000";
                return Ok(resultClass);
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                resultClass.ResultMsg = $" response: {ex.Message}";
                return StatusCode(500, resultClass);
            }
        }
        //Complaint_Upd
        //Complaint_Excel
        //Complaint_Close_Excel
        #endregion
    }
}
