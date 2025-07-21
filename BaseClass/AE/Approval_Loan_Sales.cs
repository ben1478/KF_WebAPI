using System.Text.Json.Serialization;

namespace KF_WebAPI.BaseClass.AE
{

    /// <summary>
    /// 報表資料的 C# 資料傳輸物件 (DTO)。
    /// C# 屬性名稱已完全匹配 SQL 查詢別名的大小寫。
    /// 使用 [JsonPropertyName] 屬性，確保序列化為 JSON 時，其鍵名 (key)
    /// 也與 C# 屬性名稱完全一致，覆蓋掉 ASP.NET Core 的預設 camelCase 策略。
    /// </summary>
    public class CommissionReportRow
    {
        [JsonPropertyName("isSer")]
        public string isSer { get; set; }

        [JsonPropertyName("act_service_amt")]
        public decimal act_service_amt { get; set; }

        [JsonPropertyName("service_Rate")]
        public decimal service_Rate { get; set; }

        [JsonPropertyName("fund_company")]
        public string fund_company { get; set; }

        [JsonPropertyName("isKF_CommRate")]
        public string isKF_CommRate { get; set; }

        [JsonPropertyName("isThisMonCancel")]
        public string isThisMonCancel { get; set; }

        [JsonPropertyName("CaseType")]
        public string CaseType { get; set; }

        [JsonPropertyName("item_sort")]
        public int? item_sort { get; set; }

        [JsonPropertyName("U_arrive_date")]
        public DateTime? U_arrive_date { get; set; }

        [JsonPropertyName("KFRate")]
        public string KFRate { get; set; }

        [JsonPropertyName("U_BC")]
        public string U_BC { get; set; }

        [JsonPropertyName("U_BC_rule")]
        public string U_BC_rule { get; set; }

        [JsonPropertyName("isChange")]
        public string isChange { get; set; }

        [JsonPropertyName("Ismisaligned")]
        public string Ismisaligned { get; set; }

        [JsonPropertyName("isCancel")]
        public string isCancel { get; set; }

        [JsonPropertyName("DidGet_amount")]
        public decimal DidGet_amount { get; set; }

        [JsonPropertyName("Comparison")]
        public string Comparison { get; set; }

        [JsonPropertyName("IsDiscount")]
        public string IsDiscount { get; set; }

        [JsonPropertyName("isComm")]
        public string isComm { get; set; }

        [JsonPropertyName("project_title")]
        public string project_title { get; set; }

        [JsonPropertyName("refRateI")]
        public string refRateI { get; set; }

        [JsonPropertyName("refRateL")]
        public string refRateL { get; set; }

        [JsonPropertyName("interest_rate_pass")]
        public string interest_rate_pass { get; set; }

        [JsonPropertyName("I_PID")]
        public string I_PID { get; set; }

        [JsonPropertyName("IsConfirm")]
        public string IsConfirm { get; set; }

        [JsonPropertyName("Introducer_PID")]
        public string Introducer_PID { get; set; }

        [JsonPropertyName("I_Count")]
        public int I_Count { get; set; }

        [JsonPropertyName("HS_id")]
        public long HS_id { get; set; }

        [JsonPropertyName("U_BC_name")]
        public string U_BC_name { get; set; }

        [JsonPropertyName("CS_name")]
        public string CS_name { get; set; }

        [JsonPropertyName("misaligned_date")]
        public string misaligned_date { get; set; }

        [JsonPropertyName("Send_amount_date")]
        public string Send_amount_date { get; set; }

        [JsonPropertyName("get_amount_date")]
        public string get_amount_date { get; set; }

        [JsonPropertyName("Send_result_date")]
        public string Send_result_date { get; set; }

        [JsonPropertyName("CS_introducer")]
        public string CS_introducer { get; set; }

        [JsonPropertyName("plan_name")]
        public string plan_name { get; set; }

        [JsonPropertyName("pass_amount")]
        public decimal pass_amount { get; set; }

        [JsonPropertyName("get_amount")]
        public decimal get_amount { get; set; }

        [JsonPropertyName("show_fund_company")]
        public string show_fund_company { get; set; }

        [JsonPropertyName("show_project_title")]
        public string show_project_title { get; set; }

        [JsonPropertyName("Loan_rate")]
        public string Loan_rate { get; set; }

        [JsonPropertyName("interest_rate_original")]
        public string interest_rate_original { get; set; }

        [JsonPropertyName("charge_flow")]
        public decimal charge_flow { get; set; }

        [JsonPropertyName("charge_agent")]
        public decimal charge_agent { get; set; }

        [JsonPropertyName("charge_check")]
        public decimal charge_check { get; set; }

        [JsonPropertyName("Subsidy_agent")]
        public decimal Subsidy_agent { get; set; }

        [JsonPropertyName("Subsidy_amt")]
        public decimal Subsidy_amt { get; set; }

        [JsonPropertyName("get_amount_final")]
        public decimal get_amount_final { get; set; }

        [JsonPropertyName("subsidized_interest")]
        public decimal subsidized_interest { get; set; }

        [JsonPropertyName("Expe_comm_amt")]
        public decimal Expe_comm_amt { get; set; }

        [JsonPropertyName("Expe_comm_amt_firm")]
        public decimal Expe_comm_amt_firm { get; set; }

        [JsonPropertyName("act_comm_amt")]
        public decimal act_comm_amt { get; set; }

        [JsonPropertyName("act_comm_amt_Cancel")]
        public decimal? act_comm_amt_Cancel { get; set; }

        [JsonPropertyName("act_perf_amt")]
        public decimal act_perf_amt { get; set; }

        [JsonPropertyName("Expe_perf_amt")]
        public decimal Expe_perf_amt { get; set; }

        [JsonPropertyName("Expe_perf_amt_firm")]
        public decimal Expe_perf_amt_firm { get; set; }

        [JsonPropertyName("Comm_Remark")]
        public string Comm_Remark { get; set; }

        [JsonPropertyName("Bank_name")]
        public string Bank_name { get; set; }

        [JsonPropertyName("Bank_account")]
        public string Bank_account { get; set; }
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
        [JsonPropertyName("ItemDCode")]
        public string ItemDCode { get; set; }

        [JsonPropertyName("ItemDName")]
        public string ItemDName { get; set; }

        [JsonPropertyName("KeyVal")]
        public string KeyVal { get; set; }

        [JsonPropertyName("RemitAmt")]
        public decimal RemitAmt { get; set; }

        [JsonPropertyName("CleanAmt")]
        public decimal CleanAmt { get; set; }

        [JsonPropertyName("BonusAmt")]
        public decimal BonusAmt { get; set; }

        [JsonPropertyName("SubsidyAmt")]
        public decimal SubsidyAmt { get; set; }

        [JsonPropertyName("FaresAmt")]
        public decimal FaresAmt { get; set; }

        [JsonPropertyName("TotAmt")]
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
