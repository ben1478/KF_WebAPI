using KF_WebAPI.BaseClass.AE;
using KF_WebAPI.BaseClass;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using KF_WebAPI.FunctionHandler;
using Microsoft.Data.SqlClient;
using System.Diagnostics.Eventing.Reader;
using System.Data;

namespace KF_WebAPI.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class AE_HMController : Controller
    {
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
    }
}
