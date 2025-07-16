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
    // Models/CommissionReportRow.cs
    public class CommissionReportRow
    {
        // 為了讓 C# 的屬性名稱能對應到 JSON 中不規則的大小寫名稱，
        // 我們使用 [JsonProperty] 這個屬性來做精確的映射。

        [JsonProperty("isSer")]
        public string isSer { get; set; }

        [JsonProperty("act_service_amt")]
        public decimal act_service_amt { get; set; }

        [JsonProperty("service_Rate")]
        public decimal service_Rate { get; set; }

        [JsonProperty("fund_company")]
        public string fund_company { get; set; }

        [JsonProperty("isKF_CommRate")]
        public string isKF_CommRate { get; set; }

        [JsonProperty("isThisMonCancel")]
        public string IsThisMonCancel { get; set; }

        [JsonProperty("CaseType")]
        public string CaseType { get; set; }

        [JsonProperty("item_sort")]
        public int? item_sort { get; set; }

        [JsonProperty("U_arrive_date")]
        public DateTime? U_arrive_date { get; set; }

        [JsonProperty("KFRate")]
        public string KFRate { get; set; }

        [JsonProperty("U_BC")]
        public string U_BC { get; set; }

        [JsonProperty("U_BC_rule")]
        public string U_BC_rule { get; set; }

        [JsonProperty("isChange")]
        public string isChange { get; set; }

        [JsonProperty("Ismisaligned")]
        public string Ismisaligned { get; set; }

        [JsonProperty("isCancel")]
        public string isCancel { get; set; }

        [JsonProperty("DidGet_amount")]
        public decimal DidGet_amount { get; set; }

        [JsonProperty("Comparison")]
        public string Comparison { get; set; }

        [JsonProperty("isDiscount")]
        public string IsDiscount { get; set; }

        [JsonProperty("isComm")]
        public string isComm { get; set; }

        [JsonProperty("project_title")]
        public string project_title { get; set; }

        [JsonProperty("refRateI")]
        public string refRateI { get; set; }

        [JsonProperty("refRateL")]
        public string refRateL { get; set; }

        [JsonProperty("interest_rate_pass")]
        public string interest_rate_pass { get; set; }

        [JsonProperty("I_PID")]
        public string I_PID { get; set; }

        [JsonProperty("IsConfirm")]
        public string IsConfirm { get; set; }

        [JsonProperty("Introducer_PID")]
        public string Introducer_PID { get; set; }

        [JsonProperty("I_Count")]
        public int I_Count { get; set; }

        [JsonProperty("HS_id")]
        public int HS_id { get; set; }

        [JsonProperty("U_BC_name")]
        public string U_BC_name { get; set; }

        [JsonProperty("CS_name")]
        public string CS_name { get; set; }

        [JsonProperty("misaligned_date")]
        public string misaligned_date { get; set; }

        [JsonProperty("Send_amount_date")]
        public string Send_amount_date { get; set; }

        [JsonProperty("get_amount_date")]
        public string get_amount_date { get; set; }

        [JsonProperty("Send_result_date")]
        public string Send_result_date { get; set; }

        [JsonProperty("CS_introducer")]
        public string CS_introducer { get; set; }

        [JsonProperty("plan_name")]
        public string plan_name { get; set; }

        [JsonProperty("pass_amount")]
        public decimal pass_amount { get; set; }

        [JsonProperty("get_amount")]
        public decimal get_amount { get; set; }

        [JsonProperty("show_fund_company")]
        public string show_fund_company { get; set; }

        [JsonProperty("show_project_title")]
        public string show_project_title { get; set; }

        [JsonProperty("Loan_rate")]
        public string Loan_rate { get; set; }

        [JsonProperty("interest_rate_original")]
        public string interest_rate_original { get; set; }

        [JsonProperty("charge_flow")]
        public decimal charge_flow { get; set; }

        [JsonProperty("charge_agent")]
        public decimal charge_agent { get; set; }

        [JsonProperty("charge_check")]
        public decimal charge_check { get; set; }

        [JsonProperty("Subsidy_agent")]
        public decimal Subsidy_agent { get; set; } // 代書費

        [JsonProperty("Subsidy_amt")]
        public decimal Subsidy_amt { get; set; }   // 過車費

        [JsonProperty("get_amount_final")]
        public decimal get_amount_final { get; set; }

        [JsonProperty("subsidized_interest")]
        public decimal subsidized_interest { get; set; }

        [JsonProperty("Expe_comm_amt")]
        public decimal Expe_comm_amt { get; set; }

        [JsonProperty("Expe_comm_amt_firm")]
        public decimal Expe_comm_amt_firm { get; set; }

        [JsonProperty("act_comm_amt")]
        public decimal act_comm_amt { get; set; }

        [JsonProperty("act_comm_amt_Cancel")]
        public decimal? act_comm_amt_Cancel { get; set; }

        [JsonProperty("act_perf_amt")]
        public decimal act_perf_amt { get; set; }

        [JsonProperty("Expe_perf_amt")]
        public decimal Expe_perf_amt { get; set; }

        [JsonProperty("Expe_perf_amt_firm")]
        public decimal Expe_perf_amt_firm { get; set; }

        [JsonProperty("Comm_Remark")]
        public string Comm_Remark { get; set; }

        [JsonProperty("Bank_name")]
        public string Bank_name { get; set; }

        [JsonProperty("Bank_account")]
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
