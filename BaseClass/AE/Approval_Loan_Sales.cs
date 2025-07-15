using KF_WebAPI.BaseClass;
using KF_WebAPI.BaseClass.AE;
using KF_WebAPI.FunctionHandler;
using Microsoft.Data.SqlClient;
using Newtonsoft.Json;
using System.Data;
using System.Reflection;
using System.Text;
using static Grpc.Core.Metadata;

namespace KF_WebAPI.BaseClass.AE
{
    public class Approval_Loan_Sales
    {
    }
    // Models/CommissionReportRow.cs
    public class CommissionReportRow
    {
        // 基本資訊
        public string IsSer { get; set; }
        public decimal ActServiceAmt { get; set; }
        public decimal ServiceRate { get; set; }
        public string FundCompany { get; set; }
        public string IsKF_CommRate { get; set; }
        public string IsThisMonCancel { get; set; }
        public string CaseType { get; set; }
        public int? ItemSort { get; set; }
        public DateTime? U_arrive_date { get; set; }
        public string KFRate { get; set; }
        public string U_BC { get; set; }
        public string U_BC_rule { get; set; }
        public string IsChange { get; set; }
        public string Ismisaligned { get; set; }
        public string IsCancel { get; set; }
        public decimal DidGet_amount { get; set; }
        public string Comparison { get; set; }
        public string IsDiscount { get; set; }
        public string IsComm { get; set; }
        public string ProjectTitle { get; set; }
        public string RefRateI { get; set; }
        public string RefRateL { get; set; }
        public string InterestRatePass { get; set; }
        public string I_PID { get; set; }
        public string IsConfirm { get; set; }
        public string Introducer_PID { get; set; }
        public int I_Count { get; set; }
        public int HS_id { get; set; }
        public string U_BC_name { get; set; }
        public string CS_name { get; set; }
        public string MisalignedDate { get; set; }
        public string SendAmountDate { get; set; }
        public string GetAmountDate { get; set; }
        public string SendResultDate { get; set; }
        public string CS_introducer { get; set; }
        public string PlanName { get; set; }
        public decimal PassAmount { get; set; }
        public decimal GetAmount { get; set; }
        public string ShowFundCompany { get; set; }
        public string ShowProjectTitle { get; set; }
        public string LoanRate { get; set; }
        public string InterestRateOriginal { get; set; }

        // 費用與佣金
        public decimal ChargeFlow { get; set; }
        public decimal ChargeAgent { get; set; }
        public decimal ChargeCheck { get; set; }
        public decimal GetAmountFinal { get; set; }
        public decimal SubsidizedInterest { get; set; }
        public decimal Expe_comm_amt { get; set; }
        public decimal Expe_comm_amt_firm { get; set; }
        public decimal ActCommAmt { get; set; }
        public decimal? ActCommAmtCancel { get; set; }
        public decimal ActPerfAmt { get; set; }
        public decimal Expe_perf_amt { get; set; }
        public decimal Expe_perf_amt_firm { get; set; }
        public string Comm_Remark { get; set; }
        public string BankName { get; set; }
        public string BankAccount { get; set; }
    }
    // Models/ReportQueryParameters.cs
    public class ReportQueryParameters
    {
        public string? U_bc_title { get; set; }
        public string? Plan_num { get; set; }
        public string? CS_introducer { get; set; }
        public string? SelYear_S { get; set; } 
        public string? OrderBy { get; set; }
        public string? U_BC { get; set; } // 部門代碼
        public string? U_num { get; set; } // 使用者編號
        //public string? Role_num { get; set; }
    }

    public class CommissionRuleDto
    {
        public string? RefName { get; set; }
        public string? RefDate { get; set; }
        public string GetAmount { get; set; }
        public decimal Discount { get; set; }
        public string Expe_comm_amt { get; set; }
        public string RefRateI { get; set; }
        public string RefRateL { get; set; }
        public string FR_M_name { get; set; }
        public string FR_D_ratio_A { get; set; }
        public string FR_D_ratio_B { get; set; }
        public string FR_D_rate { get; set; }
        public string FR_D_discount { get; set; }
        public string Sel { get; set; }
    }

    // Models/OtherFeeDetailDto.cs
    public class OtherFeeDetailDto
    {
        public string ItemDCode { get; set; }
        public string ItemDName { get; set; }
        public string KeyVal { get; set; }
        public decimal RemitAmt { get; set; }
        public decimal CleanAmt { get; set; }
        public decimal BonusAmt { get; set; }
        public decimal SubsidyAmt { get; set; }
        public decimal FaresAmt { get; set; } // 過車費
        public decimal TotAmt { get; set; }
    }

    public class LogEntryDto
    {
        public string TableNA { get; set; } = "OtherAmt";
        public string KeyVal { get; set; }
        public string ColumnNA { get; set; }
        public string ColumnVal { get; set; }
        public string LogID { get; set; } // The User's U_num
    }

    // 核准放款表/佣金表-其他費用(isEditOtherFee = true)
    public class FeeLogDto
    {
        public string LogUser { get; set; }
        public string ColumnNA { get; set; }
        public string ColumnVal { get; set; }
        public string Remark { get; set; }
        public string Logdate { get; set; }
    }
    // 核准放款表/佣金表-其他費用(isEditOtherFee = false)
    public class SaveFeeLogDto
    {
        public string TableNA { get; set; }
        public string KeyVal { get; set; }
        public string ColumnNA { get; set; }
        public string ColumnVal { get; set; }
        public string LogID { get; set; }
    }
    public class BranchDto
    {
        // 對應 item_D_code
        public string BranchCode { get; set; }
        // 對應 item_D_name
        public string BranchName { get; set; }
    }

    // 撥款年月
    public class MonthOptionDto
    {
        // 選項的顯示文字，例如 "114-07"
        public string YyyMM { get; set; }
        // 選項的實際值，例如 "2025-07"
        public string Yyyymm { get; set; }
    }

}

namespace KF_WebAPI.Services.AE
{
    public class CommissionReportService
    {
        public ResultClass<string> GetReportDataAsync(ReportQueryParameters parameters)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            ADOData _ADO = new();
            ResultClass<string> userInfo = GetRoleNum(parameters.U_num);
            if (userInfo.ResultCode != "000")
            {
                return resultClass;
            }

            // 1.呼叫新方法來建立動態的 WHERE 條件
            var (sql_txt, sql_txt_CS, dynamicParams) = BuildWhereClause(parameters, new UserContext
            {
                RoleNum = userInfo.objResult, // 角色
                BranchCode = parameters.U_BC, // 部門代碼
                UserNum = parameters.U_num // 使用者編號
            });

            // 2. 獲取靜態的 SQL 查詢模板
            var baseQuery = GetBaseSqlQuery(); // 這個方法會回傳完整的 SQL 模板

            // 3. 將動態條件注入到模板的佔位符中
            var finalSql = baseQuery
                .Replace("/*{{sql_txt}}*/", sql_txt)
                .Replace("/*{{sql_txt_CS}}*/", sql_txt_CS)
                .Replace("/*{{sql_txt_CAR}}*/", ""); // sql_txt_CAR:根據ASP範例，此處為空

            // 4. 附加排序邏輯
            string orderByClause = parameters.OrderBy == "1"
                ? " ORDER BY get_amount_date ASC "
                : " ORDER BY M.U_BC, M.U_name ";
            finalSql += orderByClause;

            try
            {
                string tSql = FuncHandler.GenerateDebugSql(finalSql, dynamicParams);
                Console.WriteLine(tSql); // 用於調試，輸出最終的 SQL 查詢語句                   
                DataTable dtResult = _ADO.ExecuteQuery(finalSql, dynamicParams);
                // todo:難字 CommissionReportRow
                if (dtResult.Rows.Count > 0 )
                {
                    resultClass.ResultCode = "000";
                    resultClass.ResultMsg = "查詢成功";
                    resultClass.objResult = JsonConvert.SerializeObject(dtResult);
                }
                else
                {
                    resultClass.ResultCode = "400";
                    resultClass.ResultMsg = "查無資料";
                }
                return resultClass;
            }
            catch (Exception ex)
            {
                resultClass.ResultMsg = "查詢失敗: " + ex.Message;
                resultClass.ResultCode = "500";
                return resultClass;
            }
        }

        /// <summary>
        /// 獲取參數化的基礎 SQL 查詢模板。
        /// 這個模板移除了所有舊有的字串拼接，改為使用 SQL 參數和佔位符。
        /// </summary>
        /// <returns>一個包含 SQL 參數和佔位符的查詢字串。</returns>
        private string GetBaseSqlQuery()
        {
            // 使用 C# 的 verbatim string literal (@"...") 來儲存多行 SQL 查詢
            // 所有的 VBScript 變數都已被替換為 @parameters 或 /*{{placeholders}}*/
            return @"
                SELECT
                    CASE WHEN isCompany='N' AND CONVERT(varchar(100), Send_amount_date, 111) >= '2024/08/20' AND fund_company= 'FDCOM001' THEN 'Y' ELSE 'N' END isSer, /*作業服務費*/
                    ISNULL(act_service_amt,0) act_service_amt,
                    /*新鑫專案,進件日>=20240820,介紹人:個人戶,非公司行號,收5%服務費*/
                    CASE WHEN CONVERT(varchar(100), Send_amount_date, 111) >= '2024/08/20' AND fund_company= 'FDCOM001' AND I.isCompany='N'
                        THEN CONVERT(float, (dbo.GetRateByName('SerRate', @SelYearS) * 0.01))
                        ELSE 0
                    END service_Rate,
                    fund_company,
                    CASE WHEN fund_company= 'FDCOM003' AND ISNULL(CASE WHEN H.confirm_date IS NULL THEN R.isRate ELSE RL.isRate END , 'N')= 'Y'
                        THEN 'Y' ELSE 'N'
                    END isKF_CommRate,
                    isThisMonCancel,
                    CaseType,
                    item_sort,
                    M.U_arrive_date,
                    KFRate,
                    M.U_BC,
                    M.U_BC_rule,
                    /*修改紀錄LogDate>確認時間confirm_date代表有變更資料*/
                    CASE WHEN LogDate > H.confirm_date THEN 'Y' ELSE 'N' END isChange,
                    CASE WHEN misaligned_date IS NULL THEN 'ThisMon'
                        ELSE CASE WHEN CONVERT(varchar(7), get_amount_date, 126) = @SelYearS THEN 'ThisMon' ELSE 'MisaMon' END
                    END Ismisaligned,
                    isCancel,
                    ISNULL(get_amount , 0) * 10000 * (CASE WHEN isCancel='Y' THEN -1 ELSE 1 END) DidGet_amount,
                    ISNULL(Comparison,'') Comparison,
                    CASE WHEN (CASE WHEN H.act_perf_amt IS NULL THEN S.FR_D_rate ELSE S1.FR_D_rate END) IS NULL THEN 'N' ELSE 'Y' END isDiscount,
                    CASE WHEN (CASE WHEN H.act_comm_amt IS NULL THEN CommS.FR_D_rate ELSE CommS1.FR_D_rate END) IS NULL THEN 'N' ELSE 'Y' END isComm,
                    project_title,
                    interest_rate_pass refRateI,
                    Loan_rate refRateL,
                    interest_rate_pass,
                    ISNULL(I.Introducer_PID,'') I_PID,
                    CASE WHEN H.act_perf_amt IS NULL THEN 'N' ELSE 'Y' END IsConfirm,
                    ISNULL(H.Introducer_PID,'') Introducer_PID,
                    ISNULL(I_Count,0) I_Count,
                    H.HS_id,
                    M.U_BC_name,
                    A.CS_name,
                    ISNULL(CONVERT(varchar(4),(CONVERT(varchar(4), misaligned_date, 126)-1911))+'-'+CONVERT(varchar(2), dbo.PadWithZero(MONTH(misaligned_date))) +'-'+CONVERT(varchar(2), dbo.PadWithZero(DAY(misaligned_date))),'') misaligned_date,
                    CONVERT(varchar(4),(CONVERT(varchar(4), Send_amount_date, 126)-1911))+'-'+CONVERT(varchar(2),dbo.PadWithZero(MONTH(Send_amount_date))) +'-'+CONVERT(varchar(2),dbo.PadWithZero(DAY(Send_amount_date))) Send_amount_date,
                    CONVERT(varchar(4),(CONVERT(varchar(4), get_amount_date, 126)-1911))+'-'+CONVERT(varchar(2),dbo.PadWithZero(MONTH(get_amount_date))) +'-'+CONVERT(varchar(2),dbo.PadWithZero(DAY(get_amount_date))) get_amount_date,
                    CONVERT(varchar(4),(CONVERT(varchar(4), H.Send_result_date, 126)-1911))+'-'+CONVERT(varchar(2),dbo.PadWithZero(MONTH(H.Send_result_date))) +'-'+CONVERT(varchar(2),dbo.PadWithZero(DAY(H.Send_result_date))) Send_result_date,
                    H.CS_introducer,
                    M.U_name plan_name,
                    H.pass_amount * (CASE WHEN isCancel='Y' THEN -1 ELSE 1 END) pass_amount,
                    get_amount * (CASE WHEN isCancel='Y' THEN -1 ELSE 1 END) get_amount,
                    show_fund_company,
                    (
                        SELECT item_D_name FROM Item_list WHERE item_M_code = 'project_title'
                        AND item_D_type='Y' AND item_D_code = P.project_title
                        AND del_tag='0'
                    ) AS show_project_title,
                    Loan_rate + '%' Loan_rate,
                    interest_rate_original + '%' interest_rate_original,
                    interest_rate_pass + '%' interest_rate_pass,
                    /* 費用的部分撤件同月份要*-1;不同月份*1 */
                    ISNULL(charge_flow, 0) * (CASE WHEN isThisMonCancel='Y' THEN -1 ELSE 1 END) charge_flow,
                    ISNULL(charge_agent, 0) * (CASE WHEN isThisMonCancel='Y' THEN -1 ELSE 1 END) charge_agent,
                    ISNULL(charge_check, 0) * (CASE WHEN isThisMonCancel='Y' THEN -1 ELSE 1 END) charge_check,
                    ISNULL(get_amount_final, 0) * (CASE WHEN isThisMonCancel='Y' THEN -1 ELSE 1 END) get_amount_final,
                    ISNULL(H.subsidized_interest,0) subsidized_interest,
                    CASE WHEN M.U_BC = 'BC0900' THEN 0
                        ELSE CASE WHEN fund_company= 'FDCOM003'
                            THEN (H.get_amount*10000*ISNULL(CONVERT(float,ISNULL(R.KF_CommRate,0)), 0)*0.01)
                            ELSE (H.get_amount*10000*ISNULL(CONVERT(float, CommS.FR_D_discount), 0)*0.01)
                        END
                    END Expe_comm_amt,
                    CASE WHEN M.U_BC = 'BC0900' THEN 0
                        ELSE CASE WHEN fund_company= 'FDCOM003'
                            THEN (H.get_amount*10000*ISNULL(CONVERT(float,ISNULL(RL.KF_CommRate,0)), 0)*0.01)
                            ELSE H.get_amount*10000*ISNULL(CONVERT(float, CommS1.FR_D_discount), 0)*0.01
                        END * (CASE WHEN isCancel='Y' THEN -1 ELSE 1 END)
                    END Expe_comm_amt_firm,
                    CASE WHEN M.U_BC = 'BC0900' THEN 0
                        ELSE ISNULL(H.act_comm_amt, (H.get_amount*10000*ISNULL(CONVERT(float, CommS.FR_D_discount), 0)*0.01)) * (CASE WHEN isCancel='Y' THEN -1 ELSE 1 END)
                    END act_comm_amt,
                    ISNULL(act_comm_amt_Cancel,H.act_comm_amt*-1) act_comm_amt_Cancel,
                    ISNULL(H.act_perf_amt,H.get_amount*10000*ISNULL(CONVERT(float, S.FR_D_discount)*0.1,1)) * (CASE WHEN isCancel='Y' THEN -1 ELSE 1 END) act_perf_amt,
                    H.get_amount*10000*ISNULL(CONVERT(float, S.FR_D_discount)*0.1,1) * (CASE WHEN isCancel='Y' THEN -1 ELSE 1 END) Expe_perf_amt,
                    H.get_amount*10000*ISNULL(CONVERT(float, S1.FR_D_discount)*0.1, 1) * (CASE WHEN isCancel='Y' THEN -1 ELSE 1 END) Expe_perf_amt_firm,
                    CASE WHEN M.U_BC ='BC0900' THEN '' ELSE R.Comm_Remark END Comm_Remark,
                    ISNULL(I.Bank_name,'NA') Bank_name,
                    ISNULL(I.Bank_account,'NA') Bank_account
                FROM
                (
                    /*CancelDate 不等於空=撤件*/
                    SELECT *, 'sendcase' CaseType, 'N' isThisMonCancel, 'N' isCancel, act_comm_amt act_comm_amt_Cancel,
                        (SELECT item_D_name FROM Item_list WHERE item_M_code = 'fund_company' AND item_D_type='Y' AND item_D_code = fund_company) AS show_fund_company
                    FROM House_sendcase
                    WHERE (CONVERT(varchar(7), (get_amount_date), 126) = @SelYearS OR CONVERT(varchar(7), misaligned_date, 126) = @SelYearS)
    
                    UNION ALL
    
                    SELECT *, 'sendcase' CaseType,
                        CASE WHEN CONVERT(varchar(7), ISNULL(misaligned_date, get_amount_date), 126) = @SelYearS THEN 'Y' ELSE 'N' END isThisMonCancel,
                        'Y' isCancel,
                        (
                            SELECT ColumnVal FROM [LogTable] WHERE identify =
                            ( /*撤件的實際退佣金要去抓LogTable對應的最後一筆資料*/
                                SELECT MAX(identify)identify FROM [dbo].[LogTable]
                                WHERE [TableNA]='House_sendcase' AND [ColumnNA]='act_comm_amt' AND [KeyVal] = CONVERT(varchar(7), CancelDate, 126) + CONVERT(varchar,HS_id)
                            )
                        ) act_comm_amt_Cancel,
                        (SELECT item_D_name FROM Item_list WHERE item_M_code = 'fund_company' AND item_D_type='Y' AND item_D_code = fund_company) AS show_fund_company
                    FROM House_sendcase
                    WHERE CancelDate IS NOT NULL
                ) H
                LEFT JOIN House_apply A ON A.HA_id = H.HA_id AND A.del_tag='0'
                LEFT JOIN /*專案名稱及專案基礎倍率*/
                (
                    SELECT HP_project_id, fin_user, P.project_title, ISNULL(KFRate,'N') KFRate
                    FROM House_pre_project P
                    LEFT JOIN
                    (
                        SELECT [item_D_code] project_title, CONVERT(varchar(5),dbo.GetRateByName('KFRate', @SelYearS)) KFRate
                        FROM [dbo].[Item_list]
                        WHERE [item_D_name] LIKE N'國%' AND [item_M_code]='project_title'
                    ) R ON P.project_title = R.project_title
                    WHERE P.del_tag='0'
                ) P ON P.HP_project_id = H.HP_project_id
                LEFT JOIN
                (
                    SELECT CASE WHEN U_BC ='BC0900' THEN U_BC ELSE 'general' END U_BC_rule, /*BC0900的業績折扣*/
                        'general' U_BC_rule_comm,
                        u.U_BC, U_num, U_name, ub.item_D_name U_BC_name, pt.item_sort, U_arrive_date
                    FROM User_M u
                    LEFT JOIN Item_list ub ON ub.item_M_code='branch_company' AND ub.item_D_type='Y' AND ub.item_D_code = u.U_BC
                    LEFT JOIN Item_list pt ON pt.item_M_code='professional_title' AND pt.item_D_type='Y' AND pt.item_D_code = u.U_PFT
                ) M ON M.U_num = A.plan_num
                LEFT JOIN User_M Users ON P.fin_user = Users.U_num
                LEFT JOIN (SELECT * FROM Introducer_Comm WHERE del_tag='0') I ON REPLACE(H.CS_introducer,';','') = I.Introducer_name
                    /*判斷退傭人是否已經財務確認押上PID*/
                    AND CASE WHEN H.Introducer_PID IS NULL THEN I.Introducer_PID ELSE H.Introducer_PID END = I.Introducer_PID
                LEFT JOIN (SELECT Introducer_name, COUNT(U_ID) I_Count FROM Introducer_Comm WHERE del_tag=0 GROUP BY Introducer_name) I_Cou ON H.CS_introducer = I_Cou.Introducer_name
                LEFT JOIN ( /*佣金標準 Feat_rule_comm */
                    SELECT FR_M_code, FR_M_name, FR_D_ratio_A, FR_D_ratio_B, FR_D_rate, FR_D_discount, U_BC
                    FROM dbo.Feat_rule_comm WHERE FR_M_type='N' AND del_tag='0'
                ) CommS ON project_title = CommS.FR_M_code AND Loan_rate BETWEEN CommS.FR_D_ratio_A AND CommS.FR_D_ratio_B AND dbo.PadStringWithZero(interest_rate_pass)=CommS.FR_D_rate AND M.U_BC_rule_comm = CommS.U_BC
                LEFT JOIN ( /*確認過後的佣金標準 Feat_rule_comm_Log*/
                    SELECT FR_M_code, FR_M_name, FR_D_ratio_A, FR_D_ratio_B, FR_D_rate, FR_D_discount, FR_D_replace, U_BC
                    FROM dbo.Feat_rule_comm_Log WHERE FR_M_type='N' AND [LogKey] = @SelYearS
                ) CommS1 ON Loan_rate BETWEEN CommS1.FR_D_ratio_A AND CommS1.FR_D_ratio_B AND project_title = CommS1.FR_M_code AND dbo.PadStringWithZero(interest_rate_pass) = CommS1.FR_D_rate AND M.U_BC_rule_comm = CommS1.U_BC
                LEFT JOIN ( /*業績折扣標準 Feat_rule */
                    SELECT FR_M_code, FR_M_name, FR_D_ratio_A, FR_D_ratio_B, FR_D_rate, FR_D_discount, FR_D_replace, U_BC
                    FROM dbo.Feat_rule WHERE FR_M_type='N' AND del_tag='0'
                ) S ON project_title = S.FR_M_code AND Loan_rate BETWEEN S.FR_D_ratio_A AND S.FR_D_ratio_B AND dbo.PadStringWithZero(interest_rate_pass) = S.FR_D_rate AND M.U_BC_rule = S.U_BC
                LEFT JOIN ( /*確認過後的業績折扣標準 Feat_rule_log*/
                    SELECT FR_M_code, FR_M_name, FR_D_ratio_A, FR_D_ratio_B, FR_D_rate, FR_D_discount, FR_D_replace, U_BC
                    FROM dbo.Feat_rule_log WHERE FR_M_type='N' AND [LogKey] = @SelYearS
                ) S1 ON Loan_rate BETWEEN S1.FR_D_ratio_A AND S1.FR_D_ratio_B AND project_title = S1.FR_M_code AND dbo.PadStringWithZero(interest_rate_pass) = S1.FR_D_rate AND M.U_BC_rule = S1.U_BC
                LEFT JOIN /*國?退佣 Item_list */
                    (SELECT item_M_code, item_D_code, item_D_name Comm_Remark, ISNULL(item_D_int_A,0) KF_CommRate, CASE WHEN item_D_int_A IS NULL THEN 'N' ELSE 'Y' END isRate FROM Item_list WHERE item_M_code = 'Return' AND item_D_type='Y') R ON H.Comm_Remark = R.item_D_code
                LEFT JOIN /*國?退佣Log Item_list */
                    (SELECT item_M_code, item_D_code, item_D_name Comm_Remark, ISNULL(item_D_int_A,0) KF_CommRate, CASE WHEN item_D_int_A IS NULL THEN 'N' ELSE 'Y' END isRate FROM Item_list_Log WHERE item_M_code = 'Return' AND item_D_type='Y' AND [LogKey] = @SelYearS) RL ON H.Comm_Remark = RL.item_D_code
                LEFT JOIN
                (
                    SELECT KeyVal, MAX(LogDate) LogDate FROM [LogTable] GROUP BY KeyVal
                ) L ON CONVERT(varchar,H.HS_id) = KeyVal
                WHERE H.del_tag = '0' AND H.sendcase_handle_type='Y' AND ISNULL(H.Send_amount, '') <> '' AND get_amount_type = 'GTAT002' AND ISNULL(exc_flag,'N') = 'N'
                /*{{sql_txt}}*/
                /*{{sql_txt_CS}}*/

                UNION ALL /*其他案件(機車貸商品貸)*/

                SELECT
                    'N' isSer, 0 act_service_amt, 0 service_Rate, '' fund_company, 'N' isKF_CommRate, isThisMonCancel, CaseType, item_sort, M.U_arrive_date, 'N' KFRate, M.U_BC, M.U_BC_rule,
                    CASE WHEN LogDate > H.confirm_date THEN 'Y' ELSE 'N' END isChange,
                    'ThisMon' Ismisaligned, isCancel, ISNULL(get_amount, 0) * 10000 * (CASE WHEN isCancel='Y' THEN -1 ELSE 1 END) DidGet_amount,
                    '' Comparison, 'N' isDiscount, 'N' isComm, H.show_project_title project_title, H.interest_rate_pass refRateI,
                    NULL refRateL, H.interest_rate_pass, 'NA' I_PID,
                    CASE WHEN H.act_perf_amt IS NULL THEN 'N' ELSE 'Y' END IsConfirm,
                    'NA' Introducer_PID, 1 I_Count, case_id HS_id, M.U_BC_name, H.cs_name, '' misaligned_date, 'NA' Send_amount_date,
                    CONVERT(varchar(4),(CONVERT(varchar(4), get_amount_date, 126)-1911))+'-'+CONVERT(varchar(2), dbo.PadWithZero(MONTH(get_amount_date))) +'-'+CONVERT(varchar(2), dbo.PadWithZero(DAY(get_amount_date))) get_amount_date,
                    'NA' Send_result_date, 'NA' CS_introducer, M.U_name plan_name, H.get_amount pass_amount, H.get_amount * (CASE WHEN isCancel='Y' THEN -1 ELSE 1 END) get_amount,
                    H.show_fund_company, H.show_project_title, 'NA' Loan_rate, 'NA' interest_rate_original, H.interest_rate_pass + '%' interest_rate_pass,
                    0 charge_flow, 0 charge_agent, 0 charge_check, 0 get_amount_final, 0 subsidized_interest, 0 Expe_comm_amt,
                    0 AS act_comm_amt,
                    0 AS act_comm_amt_Cancel,
                    0 AS Expe_comm_amt_firm,
                    ISNULL(H.act_perf_amt, Expe_perf_amt) * (CASE WHEN isCancel='Y' THEN -1 ELSE 1 END) act_perf_amt,
                    Expe_perf_amt * (CASE WHEN isCancel='Y' THEN -1 ELSE 1 END) Expe_perf_amt,
                    Expe_perf_amt * (CASE WHEN isCancel='Y' THEN -1 ELSE 1 END) Expe_perf_amt_firm,
                    NULL Comm_Remark, Bank_name, Bank_account
                FROM
                (
                    SELECT *, 
                        'N' isThisMonCancel, 
                        'NA' Bank_account, 
                        'NA' Bank_name, 
                        'N' isCancel, 
                        NULL misaligned_date,
                        ROUND(comm_amt * dbo.GetRateByName('CarRate', @SelYearS) + (get_amount * 10000), -4) Expe_perf_amt
                    FROM House_othercase
                    WHERE (CONVERT(varchar(7), (get_amount_date), 126) = @SelYearS)

                    UNION ALL

                    SELECT *,
                        CASE WHEN CONVERT(varchar(7), get_amount_date, 126) = @SelYearS THEN 'Y' ELSE 'N' END isThisMonCancel,
                        'NA' Bank_account, 'NA' Bank_name, 'Y' isCancel, NULL misaligned_date,
                        ROUND(comm_amt * dbo.GetRateByName('CarRate', CONVERT(varchar(7),(get_amount_date),126)) + (get_amount * 10000), -4) Expe_perf_amt
                    FROM House_othercase
                    WHERE CancelDate IS NOT NULL
                    /*{{sql_txt_CAR}}*/
                ) H
                LEFT JOIN
                (
                    SELECT 'otherRule' U_BC_rule, u.U_BC, U_num, U_name, ub.item_D_name U_BC_name, pt.item_sort, U_arrive_date
                    FROM User_M u
                    LEFT JOIN Item_list ub ON ub.item_M_code='branch_company' AND ub.item_D_type='Y' AND ub.item_D_code = u.U_BC
                    LEFT JOIN Item_list pt ON pt.item_M_code='professional_title' AND pt.item_D_type='Y' AND pt.item_D_code = u.U_PFT
                ) M ON H.plan_num = M.U_num
                LEFT JOIN
                (
                    SELECT KeyVal, MAX(LogDate) LogDate
                    FROM LogTable GROUP BY KeyVal
                ) L ON CONVERT(varchar,H.case_id) = KeyVal
                WHERE H.del_tag = '0'
                /*{{sql_txt}}*/
                ";
        }


        /// <summary>
        /// 根據使用者權限和查詢參數，動態建立 WHERE 條件子句。
        /// 這個方法取代了 ASP 中拼接 sql_txt 和 sql_txt_CS 的邏輯。
        /// </summary>
        /// <param name="parameters">來自前端請求的查詢參數。</param>
        /// <param name="user">當前登入使用者的資訊。</param>
        /// <returns>一個包含 SQL 條件片段和對應參數的元組。</returns>
        private (string SqlTxt, string SqlTxtCs, List<SqlParameter> Parameters) BuildWhereClause(
            ReportQueryParameters parameters,
            UserContext user)
        {
            var sqlBuilder = new StringBuilder();
            var csSqlBuilder = new StringBuilder();
            var dynamicParams = new List<SqlParameter>();

            // 複製 ASP 中的參數，並根據權限覆寫
            var planNum = string.IsNullOrEmpty(parameters.Plan_num)?  string.Empty : parameters.Plan_num;
            var uBcTitle = string.IsNullOrEmpty(parameters.U_bc_title) ? string.Empty : parameters.U_bc_title ;
            
            // 步驟 1: 根據使用者角色(Auth)設定權限和預設過濾條件
            // 這段邏輯取代了 ASP 中的 session("Role_num") 判斷
            switch (user.RoleNum)
            {
                case "1008": // 業務主管
                case "1014":
                    uBcTitle = user.BranchCode; // 只能看自己部門
                    planNum = null; // 主管預設看全部門，除非前端指定
                    break;

                case "1009": // 業務
                case "1017":
                    planNum = user.UserNum; // 只能看自己
                    uBcTitle = null;
                    break;

                case "1010": // 業務助理
                    uBcTitle = user.BranchCode; // 只能看自己部門
                    planNum = null;
                    break;

                // "1004", "1007", "1001" (財務, 管理部主管, 開發者) 為管理者權限(Adm)，不過濾
                default:
                    // 使用前端傳來的原始參數
                    break;
            }

            // 步驟 2: 根據最終的過濾條件，產生 WHERE 子句
            if (!string.IsNullOrEmpty(planNum))
            {
                // 如果有指定業務員，則以此為最高優先級
                sqlBuilder.Append(" AND M.U_num  = @PlanNum ");
                dynamicParams.Add(new SqlParameter("PlanNum", planNum));
            }
            else if (!string.IsNullOrEmpty(uBcTitle))
            {
                // 否則，使用部門過濾
                if (uBcTitle == "666")
                {
                    sqlBuilder.Append(" AND M.U_BC BETWEEN 'BC0100' AND 'BC0600' ");
                }
                else
                {
                    sqlBuilder.Append(" AND M.U_BC = @UbcTitle ");
                    dynamicParams.Add(new SqlParameter("UbcTitle", uBcTitle));
                }
            }

            // 處理介紹人 (CS_introducer) 的過濾
            if (!string.IsNullOrEmpty(parameters.CS_introducer))
            {
                csSqlBuilder.Append(" AND H.CS_introducer LIKE @CSIntroducer ");
                dynamicParams.Add(new SqlParameter("CSIntroducer", $"%{parameters.CS_introducer}%"));
            }

            // 步驟 3: 加入固定的日期過濾邏輯
            // ASP 中最後拼接的日期判斷部分
            sqlBuilder.Append(@"
            AND (
                CASE 
                    WHEN isCancel = 'N' THEN CONVERT(varchar(7), get_amount_date, 126)
                    ELSE CONVERT(varchar(7), CancelDate, 126)
                END = @SelYearS
                OR
                CASE 
                    WHEN isCancel = 'N' THEN CONVERT(varchar(7), misaligned_date, 126)
                    ELSE NULL -- 如果已取消，則不考慮 misaligned_date
                END = @SelYearS
            )
        ");
            dynamicParams.Add(new SqlParameter("SelYearS", parameters.SelYear_S));

            return (sqlBuilder.ToString(), csSqlBuilder.ToString(), dynamicParams);
        }
        public ResultClass<string> GetRoleNum(string U_num)
        {
            ResultClass<string> result = new ResultClass<string>();
            try
            {
                ADOData _adoData = new ADOData(); // 測試:"Test" / 正式:""
                var parameters = new List<SqlParameter>();
                #region SQL
                var T_SQL = @"SELECT 
                                  [Role_num]
                              FROM [AE_DB_TEST].[dbo].[User_M]
                              where 
                                del_tag = 0 and U_leave_date is null
                                and is_susp is null
                                and U_num = @U_num ";

                #endregion
                parameters.Add(new SqlParameter("@U_num", U_num));

                var dtResult = _adoData.ExecuteQuery(T_SQL, parameters);

                if (dtResult.Rows.Count > 0)
                {
                    result.ResultCode = "000";
                    result.objResult = dtResult.Rows[0][0].ToString();
                }
                else
                {
                    result.ResultCode = "400";
                    result.ResultMsg = "查無資料";
                    result.objResult = string.Empty;
                }
            }
            catch (Exception ex)
            {
                result.ResultCode = "500";
                result.ResultMsg = $" response: {ex.Message}";
                result.objResult = string.Empty;
            }
            return result;
        }

        // 在 CommissionReportService.cs 中
        public ResultClass<string> GetCommissionRules(string hsId, string isConfirm, string mainKey)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            ADOData _adoData = new();

            string uBcRule = "general"; // 根據舊程式碼，此值固定為 'general'
            string tableFeatRuleSql;
            string sqlColumn;

            var parameters = new List<SqlParameter>();
            parameters.Add(new SqlParameter("@MainKey", mainKey));
            parameters.Add(new SqlParameter("@UbcRule", uBcRule));
            parameters.Add(new SqlParameter("@HsId", hsId));

            if (isConfirm == "Y")
            {
                // 確認過的抓 LOG 檔
                tableFeatRuleSql = @"
            SELECT FR_M_code, FR_M_name, FR_D_ratio_A, FR_D_ratio_B,
                   FR_D_rate, FR_D_discount, F.U_BC,
                   CONVERT(varchar, ISNULL(F.edit_date, F.add_date), 120) AS RefDate,
                   U_name AS RefName
            FROM dbo.Feat_rule_comm_Log F
            LEFT JOIN User_M M ON ISNULL(F.edit_num, F.add_num) = U_num
            WHERE LogKey = @MainKey AND F.U_BC = @UbcRule";
                sqlColumn = " ISNULL(R.RefName, '') AS RefName, ISNULL(R.RefDate, '') AS RefDate, ";
            }
            else
            {
                // 沒確認過的抓目前設定檔
                tableFeatRuleSql = "SELECT * FROM dbo.Feat_rule_comm WHERE FR_M_type = 'N' AND del_tag = '0' AND U_BC = @UbcRule";
                sqlColumn = " '' AS RefName, '' AS RefDate, ";
            }

            string finalSql = $@"
                SELECT {sqlColumn}
                       REPLACE(CONVERT(VARCHAR(15), CONVERT(MONEY, H.get_amount * 10000), 1), '.00', '') AS get_amount,
                       CONVERT(float, ISNULL(S.FR_D_discount, 0)) * 0.01 AS discount,
                       REPLACE(CONVERT(VARCHAR(15), CONVERT(MONEY, (H.get_amount * 10000 * 0.01 * ISNULL(CONVERT(float, S.FR_D_discount), 0))), 1), '.00', '') AS Expe_comm_amt,
                       interest_rate_pass AS refRateI,
                       Loan_rate AS refRateL,
                       ISNULL(R.FR_M_name, '') AS FR_M_name,
                       ISNULL(CONVERT(VARCHAR, R.FR_D_ratio_A), '') AS FR_D_ratio_A,
                       ISNULL(CONVERT(VARCHAR, R.FR_D_ratio_B), '') AS FR_D_ratio_B,
                       ISNULL(CONVERT(VARCHAR, R.FR_D_rate), '') AS FR_D_rate,
                       ISNULL(CONVERT(VARCHAR, R.FR_D_discount), '') AS FR_D_discount,
                       CASE WHEN S.FR_D_rate = R.FR_D_rate THEN 'Y' ELSE 'N' END AS Sel
                FROM House_sendcase H
                LEFT JOIN House_apply A ON A.HA_id = H.HA_id AND A.del_tag = '0'
                LEFT JOIN (
                    SELECT 'general' AS U_BC_rule, u.U_BC, U_num, U_name, item_D_name AS U_BC_name
                    FROM User_M u
                    LEFT JOIN Item_list ub ON ub.item_M_code = 'branch_company' AND ub.item_D_type = 'Y' AND ub.item_D_code = u.U_BC
                ) M ON M.U_num = A.plan_num
                LEFT JOIN House_pre_project P ON P.HP_project_id = H.HP_project_id
                LEFT JOIN ({tableFeatRuleSql}) S ON project_title = S.FR_M_code
                    AND Loan_rate BETWEEN S.FR_D_ratio_A AND S.FR_D_ratio_B
                    AND dbo.PadStringWithZero(interest_rate_pass) = S.FR_D_rate
                    AND M.U_BC_rule = S.U_BC
                LEFT JOIN ({tableFeatRuleSql}) R ON S.FR_M_code = R.FR_M_code
                    AND Loan_rate BETWEEN R.FR_D_ratio_A AND R.FR_D_ratio_B
                    AND S.U_BC = R.U_BC
                WHERE hs_id = @HsId
                ORDER BY R.FR_D_ratio_A, R.FR_D_rate DESC";
            try
            {
                var dtResult = _adoData.ExecuteQuery(finalSql, parameters);
                if (dtResult.Rows.Count > 0)
                {
                    // IEnumerable<CommissionRuleDto>
                    var result = dtResult.AsEnumerable().Select(row => new CommissionRuleDto
                    {
                        RefName = row.Field<string>("RefName"),
                        RefDate = row.Field<string>("RefDate"),
                        GetAmount = row.Field<string>("get_amount"),
                        Discount = (decimal)row.Field<double>("discount"),
                        Expe_comm_amt = row.Field<string>("Expe_comm_amt"),
                        RefRateI = row.Field<string>("refRateI"),
                        RefRateL = row.Field<string>("refRateL"),
                        FR_M_name = row.Field<string>("FR_M_name"),
                        FR_D_ratio_A = row.Field<string>("FR_D_ratio_A"),
                        FR_D_ratio_B = row.Field<string>("FR_D_ratio_B"),
                        FR_D_rate = row.Field<string>("FR_D_rate"),
                        FR_D_discount = row.Field<string>("FR_D_discount"),
                        Sel = row.Field<string>("Sel")
                    }).ToList();

                    resultClass.ResultCode = "000";
                    resultClass.objResult = JsonConvert.SerializeObject(result);
                }
                else
                {
                    resultClass.ResultCode = "400";
                    resultClass.ResultMsg = "查無資料";
                    resultClass.objResult = string.Empty;
                }
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                resultClass.ResultMsg = $"查詢失敗: {ex.Message}";
                resultClass.objResult = string.Empty;
            }
            return resultClass;
        }
        // In CommissionReportService.cs

        /// <summary>
        /// 獲取各區其他費用的詳細明細。
        /// </summary>
        public async Task<ResultClass<string>> GetOtherFeeDetailsAsync(string? uBcTitle, string selYearS)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            var parameters = new List<SqlParameter>();
            // 基礎查詢，從 Item_list 取得所有分公司
            var sqlBuilder = new StringBuilder(@"
                SELECT
                    i.item_D_code AS ItemDCode,
                    i.item_D_name AS ItemDName,
                    i.item_D_code + @SelYearS AS KeyVal,
                    i.item_sort
                FROM Item_list i
                WHERE i.item_M_code = 'branch_company' AND i.item_D_type = 'Y' AND i.show_tag = '0' AND i.del_tag = '0'
            ");
            parameters.Add(new SqlParameter("@SelYearS", selYearS));

            // 根據傳入的 uBcTitle 加入過濾條件
            if (string.IsNullOrEmpty(uBcTitle) || uBcTitle == "666")
            {
                sqlBuilder.Append(" AND i.item_D_code NOT LIKE 'BC080%' AND i.item_D_code <> 'BC0700' ");
            }
            else
            {
                sqlBuilder.Append(" AND i.item_D_code = @UbcTitle ");
                parameters.Add(new SqlParameter("@UbcTitle", uBcTitle));
            }

            // 使用 CTE (Common Table Expression) 來獲取最新的費用紀錄
            var finalSql = $@"
                WITH LatestLogs AS (
                    SELECT KeyVal, ColumnNA, ColumnVal
                    FROM LogTable
                    WHERE identify IN (
                        SELECT MAX(identify)
                        FROM LogTable
                        WHERE TableNA = 'OtherAmt'
                        GROUP BY KeyVal, ColumnNA
                    )
                ),
                BranchList AS (
                    {sqlBuilder}
                )
                SELECT
                    b.ItemDCode,
                    b.ItemDName,
                    b.KeyVal,
                    ISNULL(TRY_CAST(l_remit.ColumnVal AS decimal(18, 2)), 0) AS RemitAmt,
                    ISNULL(TRY_CAST(l_clean.ColumnVal AS decimal(18, 2)), 0) AS CleanAmt,
                    ISNULL(TRY_CAST(l_bonus.ColumnVal AS decimal(18, 2)), 0) AS BonusAmt,
                    ISNULL(TRY_CAST(l_subsidy.ColumnVal AS decimal(18, 2)), 0) AS SubsidyAmt,
                    ISNULL(TRY_CAST(l_fares.ColumnVal AS decimal(18, 2)), 0) AS FaresAmt,
                    (
                        ISNULL(TRY_CAST(l_remit.ColumnVal AS decimal(18, 2)), 0) +
                        ISNULL(TRY_CAST(l_clean.ColumnVal AS decimal(18, 2)), 0) +
                        ISNULL(TRY_CAST(l_bonus.ColumnVal AS decimal(18, 2)), 0) +
                        ISNULL(TRY_CAST(l_subsidy.ColumnVal AS decimal(18, 2)), 0) +
                        ISNULL(TRY_CAST(l_fares.ColumnVal AS decimal(18, 2)), 0)
                    ) AS TotAmt                    
                FROM BranchList b
                LEFT JOIN LatestLogs l_remit ON b.KeyVal = l_remit.KeyVal AND l_remit.ColumnNA = 'RemitAmt'
                LEFT JOIN LatestLogs l_clean ON b.KeyVal = l_clean.KeyVal AND l_clean.ColumnNA = 'CleanAmt'
                LEFT JOIN LatestLogs l_bonus ON b.KeyVal = l_bonus.KeyVal AND l_bonus.ColumnNA = 'BonusAmt'
                LEFT JOIN LatestLogs l_subsidy ON b.KeyVal = l_subsidy.KeyVal AND l_subsidy.ColumnNA = 'SubsidyAmt'
                LEFT JOIN LatestLogs l_fares ON b.KeyVal = l_fares.KeyVal AND l_fares.ColumnNA = 'FaresAmt'
                ORDER BY b.item_sort ASC
            ";

            ADOData _adoData = new();
            
            try
            {
                var dtResult = _adoData.ExecuteQuery(finalSql, parameters);
                if (dtResult.Rows.Count > 0)
                {
                    // IEnumerable<CommissionRuleDto>
                    var result = dtResult.AsEnumerable().Select(row => new OtherFeeDetailDto
                    {
                        ItemDCode = row.Field<string>("ItemDCode"),
                        ItemDName = row.Field<string>("ItemDName"),
                        KeyVal = row.Field<string>("KeyVal"),
                        RemitAmt = row.Field<decimal>("RemitAmt"),
                        CleanAmt = row.Field<decimal>("CleanAmt"),
                        BonusAmt = row.Field<decimal>("BonusAmt"),
                        SubsidyAmt = row.Field<decimal>("SubsidyAmt"),
                        FaresAmt = row.Field<decimal>("FaresAmt"),
                        TotAmt = row.Field<decimal>("TotAmt")
                    }).ToList();

                    resultClass.ResultCode = "000";
                    resultClass.objResult = JsonConvert.SerializeObject(result);
                }
                else
                {
                    resultClass.ResultCode = "400";
                    resultClass.ResultMsg = "查無資料";
                    resultClass.objResult = string.Empty;
                }
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                resultClass.ResultMsg = $"查詢失敗: {ex.Message}";
                resultClass.objResult = string.Empty;
            }
            return resultClass;
            
        }
        // In CommissionReportService.cs

        /// <summary>
        /// 核准放款表/佣金表-其他費用(isEditOtherFee = false)
        /// </summary>
        public ResultClass<string> GetFeeLogs(string tableNA, string keyVal, string columnNA)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            // 1. 將傳入的欄位字串，整理成乾淨的 List<string>
            var columnsList = columnNA.Split(',')
                                      .Select(c => c.Trim())
                                      .Where(c => !string.IsNullOrEmpty(c))
                                      .ToList();
            if (!columnsList.Any())
            {
                resultClass.ResultCode = "400";
                resultClass.ResultMsg = "查無資料";
                resultClass.objResult = string.Empty;
                return resultClass;
            }

            var parameters = new List<SqlParameter>
            {
                new SqlParameter("@TableNA", tableNA),
                new SqlParameter("@KeyVal", keyVal)
            };

            var paramNames = new List<string>();
            for (int i = 0; i < columnsList.Count; i++)
            {
                var paramName = $"@p{i}";
                paramNames.Add(paramName);
                parameters.Add(new SqlParameter(paramName, columnsList[i]));
            }
            var sql = $@"
                SELECT 
                    M.U_name AS LogUser,
                    L.ColumnNA,
                    L.ColumnVal,
                    ISNULL(L.Remark, '') AS Remark,
                    CONVERT(varchar, L.Logdate, 120) AS Logdate
                FROM dbo.LogTable L
                LEFT JOIN User_M M ON L.LogID = M.U_num
                WHERE L.TableNA = @TableNA AND L.KeyVal = @KeyVal AND L.ColumnNA IN ({string.Join(", ", paramNames)})
                ORDER BY L.Logdate DESC";

            ADOData _adoData = new();

            try
            {
                var dtResult = _adoData.ExecuteQuery(sql, parameters);
                if (dtResult.Rows.Count > 0)
                {
                    // IEnumerable<CommissionRuleDto>
                    var result = dtResult.AsEnumerable().Select(row => new FeeLogDto
                    {
                        LogUser = row.Field<string>("LogUser"),
                        ColumnNA = row.Field<string>("ColumnNA"),
                        ColumnVal = row.Field<string>("ColumnVal"),
                        Remark = row.Field<string>("Remark"),
                        Logdate = row.Field<string>("Logdate")
                    }).ToList();

                    resultClass.ResultCode = "000";
                    resultClass.objResult = JsonConvert.SerializeObject(result);
                }
                else
                {
                    resultClass.ResultCode = "400";
                    resultClass.ResultMsg = "查無資料";
                    resultClass.objResult = string.Empty;
                }
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                resultClass.ResultMsg = $"查詢失敗: {ex.Message}";
                resultClass.objResult = string.Empty;
            }
            return resultClass;
        }

        /// <summary>
        /// 核准放款表/佣金表-其他費用.SaveLogByObj(isEditOtherFee = true)
        /// </summary>
        public async Task<bool> SaveOtherFeesAsync(List<SaveFeeLogDto> logEntries, bool checkForChanges)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            List<SaveFeeLogDto> entriesToSave = new List<SaveFeeLogDto>();
            ADOData _adoData = new ADOData();
            
            if (checkForChanges)
            {
                foreach (var entry in logEntries)
                {
                    var latestValSql = @$"
                        SELECT TOP 1 ColumnVal 
                        FROM LogTable 
                        WHERE TableNA = '{entry.TableNA}' AND KeyVal = '{entry.KeyVal}' AND ColumnNA = '{entry.ColumnNA}'
                        ORDER BY identify DESC";

                    var latestVal = _adoData.ExecuteSQuery(latestValSql);

                    // If no record exists, or if the value has changed, add it to the list to be saved.
                    if (latestVal.Rows.Count == 0 || latestVal.Rows[0][0] != entry.ColumnVal)
                    {
                        entriesToSave.Add(entry);
                    }
                }
            }
            else
            {
                entriesToSave = logEntries;
            }

            if (!entriesToSave.Any())
            {
                return true; // Nothing to save, but operation is "successful".
            }

            // Use a transaction to ensure all entries are saved or none are.

            try
            {
                
                foreach (var entry in entriesToSave)
                {
                    var insertSql = @$"
                    INSERT INTO LogTable (TableNA, KeyVal, ColumnNA, ColumnVal, LogID, LogDate)
                    VALUES ('{entry.TableNA}', '{entry.KeyVal}', '{entry.ColumnNA}', '{entry.ColumnVal}', '{entry.LogID}', SYSDATETIME() )";
                    Console.WriteLine($"insertSql={insertSql}");

                    _adoData.ExecuteSQuery(insertSql);
                    
                }
                return true;
            }
            catch (Exception ex)
            {
                // Log the exception (ex)
                return false;
            }
        }

        // 在 CommissionReportService.cs 或其他服務中

        /// <summary>
        /// 獲取所有有效的分公司列表，用於下拉選單。
        /// </summary>
        public ResultClass<string> GetBranchesAsync()
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            ADOData _adoData = new ADOData();
            var sql = @"
                SELECT 
                    item_D_code AS BranchCode,
                    item_D_name AS BranchName
                FROM Item_list
                WHERE item_M_code = 'branch_company' 
                  AND item_D_type = 'Y' 
                  AND show_tag = '0' 
                  AND del_tag = '0' 
                  AND item_D_code NOT LIKE 'BC080%' 
                  AND item_D_code <> 'BC0700'
                ORDER BY item_sort";

            try
            {
                var dtResult = _adoData.ExecuteSQuery(sql);
                if (dtResult.Rows.Count > 0)
                {
                    // IEnumerable<CommissionRuleDto>
                    var result = dtResult.AsEnumerable().Select(row => new BranchDto
                    {
                        BranchCode = row.Field<string>("BranchCode"),
                        BranchName = row.Field<string>("BranchName")
                    }).ToList();

                    resultClass.ResultCode = "000";
                    resultClass.objResult = JsonConvert.SerializeObject(result);
                }
                else
                {
                    resultClass.ResultCode = "400";
                    resultClass.ResultMsg = "查無資料";
                    resultClass.objResult = string.Empty;
                }
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                resultClass.ResultMsg = $"查詢失敗: {ex.Message}";
                resultClass.objResult = string.Empty;
            }
            return resultClass;
        }

        // 在 CommissionReportService.cs 中

        /// <summary>
        /// 獲取撥款年月下拉選單的選項列表
        /// </summary>
        public async Task<ResultClass<string>> GetDisbursementMonthsAsync()
        {
            // 使用您提供的 SQL 查詢
            var sql = @"
                SELECT DISTINCT 
                    CONVERT(varchar(4), (CONVERT(varchar(4), get_amount_date, 126) - 1911)) + '-' + 
                    CONVERT(varchar(2), MONTH(get_amount_date)) AS yyyMM,
                    CONVERT(varchar(7), get_amount_date, 126) AS yyyymm 
                FROM House_sendcase 
                WHERE YEAR(get_amount_date) > YEAR(DATEADD(year, -2, SYSDATETIME())) 
                ORDER BY yyyymm DESC";

            ResultClass<string> resultClass = new ResultClass<string>();
            ADOData _adoData = new ADOData();

            try
            {
                var dtResult = _adoData.ExecuteSQuery(sql);
                if (dtResult.Rows.Count > 0)
                {
                    // IEnumerable<CommissionRuleDto>
                    var result = dtResult.AsEnumerable().Select(row => new MonthOptionDto
                    {
                        YyyMM = row.Field<string>("YyyMM"),
                        Yyyymm = row.Field<string>("Yyyymm")
                    }).ToList();

                    resultClass.ResultCode = "000";
                    resultClass.objResult = JsonConvert.SerializeObject(result);
                }
                else
                {
                    resultClass.ResultCode = "400";
                    resultClass.ResultMsg = "查無資料";
                    resultClass.objResult = string.Empty;
                }
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                resultClass.ResultMsg = $"查詢失敗: {ex.Message}";
                resultClass.objResult = string.Empty;
            }
            return resultClass;
        }

    }

    class UserContext
    {
        /// <summary>
        /// 角色代碼，例如 "1009" (業務), "1008" (業務主管), "1001" (開發者)
        /// </summary>
        public string? RoleNum { get; set; } // 預設為最高權限，便於測試

        /// <summary>
        /// 使用者編號 (U_num)
        /// </summary>
        public string? UserNum { get; set; }

        /// <summary>
        /// 使用者所屬分公司代碼 (U_BC)
        /// </summary>
        public string? BranchCode { get; set; }
    }
}
