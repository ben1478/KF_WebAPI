using Newtonsoft.Json;

namespace KF_WebAPI.BaseClass.AE
{
    public class UserGroup
    {
        public string U_num { get; set; }
    }
    public class Feat_user_list_req
    {
        public string start_date { get; set; }
        public string end_date { get; set; }

        public string U_num { get; set; }
    }
    public class Feat_user_list_res
    {
        public string CS_name { get; set; }
        public decimal HS_id { get; set; }
        public string plan_name { get; set; }
        public string show_appraise_company { get; set; }
        public string show_project_title { get; set; }
        public DateTime get_amount_date { get; set; }
        public string get_amount { get; set; }
        public string Loan_rate { get; set; }
        public string interest_rate_original { get; set; }
        public string interest_rate_pass { get; set; }
        public string charge_M { get; set; }
        public string charge_flow { get; set; }
        public string charge_agent { get; set; }
        public string charge_check { get; set; }
        public string get_amount_final { get; set; }
        public string CS_introducer { get; set; }
        public string HS_note { get; set; }
        public string project_title { get; set; }
        public string exception_type { get; set; }
        public string exception_rate { get; set; }
        public string FR_D_discount { get; set; }
        public string Exceptions { get; set; }
        public string? Performance_Discount { get; set; }
        public string? Performance_Amount { get; set; }
        public Feat_rule_Detail Feat_rule_Detail { get; set; }
    }
    //備註其他資料資訊
    public class Feat_rule_Detail
    {
        public decimal FR_D_ratio_A { get; set; }
        public decimal FR_D_ratio_B { get; set; }
        public string FR_D_rate { get; set; }
        public int show_FR_D_discount { get; set; }
        public string FR_D_replace { get; set; }

    }
    //總計跟業績獎金計算資訊
    public class Feat_rule_Totail
    {
        public decimal total_check_amount { get; set; }
        public string? range_D_base { get; set; }
        public string? check_rate { get; set; }
        public string? range_D_ratio_A { get; set; }
        public string? range_D_ratio_B { get; set; }
        public string? range_D_rate { get; set; }
        public string? range_D_reward { get; set; }
        public int? total_get_amount { get; set; }
        public decimal show_cash { get; set; }

    }

    public class Feat_leader_list_req
    {
        public string start_date { get; set; }
        public string end_date { get; set; }
        public string leaders { get; set; }
    }

    public class Feat_target
    {
        public tbInfo? tbInfo { get; set; }
        public decimal? target_id { get; set; }
        public int? target_ym { get; set; }
        public int? target_quota { get; set; }
        public int? group_id { get; set; }
        public string? U_num { get; set; }

    }

    public class Feat_Target_Upd
    {
        public string U_num { get; set; }
        public int target_ym { get; set; }
        public int target_quota { get; set; }
        public decimal group_M_id { get; set; }

    }

    public class Target_YYYY
    {
        public string U_BC { get; set; }
        public int Target { get; set; }
    }

    public class MonthQuota_res
    {
        public decimal group_M_id { get; set; }
        public string group_M_name { get; set; }
        public string group_D_name { get; set; }
        public string group_D_code { get; set; }
        public int U_PFT_sort { get; set; }
        public string U_PFT_name { get; set; }
        public int target_quota { get; set; }
        public string target_ym { get; set; }

    }
    public class FeatDailyPerson_req
    {
        public string reportDate_n { get; set; }
        public string company { get; set; }
        public string U_BC { get; set; }
    }
    public class FeatDailyReport_req
    {
        public string reportDate_n { get; set; }
        public string U_BC { get; set; }
    }
    public class FeatDailyReport_excel
    {
        public string plan_name { get; set; }
        public string U_PFT_name { get; set; }
        public int day_incase_num_FDCOM001 { get; set; }
        public int month_incase_num_FDCOM001 { get; set; }
        public int day_incase_num_FDCOM003 { get; set; }
        public int month_incase_num_FDCOM003 { get; set; }
        public int day_get_amount_num { get; set; }
        public int day_get_amount { get; set; }
        public int month_pass_num { get; set; }
        public int month_get_amount_num { get; set; }
        public int month_get_amount_FDCOM001 { get; set; }
        public int month_get_amount_FDCOM003 { get; set; }
        public int month_pass_amount_FDCOM001 { get; set; }
        public int month_pre_amount_FDCOM001 { get; set; }
        public int month_pass_amount_FDCOM003 { get; set; }
        public decimal advance_payment_AE { get; set; }
        public int target_quota { get; set; }
        public string AchievementRate { get; set; }
    }

    public class FeatDailyReport_excel_Total
    {
        public int day_incase_num_FDCOM001_total { get; set; }
        public int month_incase_num_FDCOM001_total { get; set; }
        public int day_incase_num_FDCOM003_total { get; set; }
        public int month_incase_num_FDCOM003_total { get; set; }
        public int day_get_amount_num_total { get; set; }
        public int day_get_amount_total { get; set; }
        public int month_pass_num_total { get; set; }
        public int month_get_amount_num_total { get; set; }
        public int month_get_amount_FDCOM001_total { get; set; }
        public int month_get_amount_FDCOM003_total { get; set; }
        public int month_pass_amount_FDCOM001_total { get; set; }
        public int month_pre_amount_FDCOM001_total { get; set; }
        public int month_pass_amount_FDCOM003_total { get; set; }
        public decimal advance_payment_AE_total { get; set; }
        public int target_quota_total { get; set; }
    }

    public class SendCaseStatu_req
    {
        public DateTime Send_Date_S { get; set; }
        public DateTime Send_Date_E { get; set; }
        public string Company { get; set; }
        public string? status { get; set; }

    }

    public class SendCaseStatu_res
    {
        public string fund_company { get; set; }
        public string Company_Name { get; set; }
        public int totleCount { get; set; }
        public int ApprCount { get; set; }
        public int unApprCount { get; set; }
        public int PayCount { get; set; }
        public int unPayCount { get; set; }
        public int WPayCount { get; set; }
        public int GUCount { get; set; }
        /// <summary>
        /// 送審中筆數
        /// </summary>
        public int Review_count { get; set; }
        /// <summary>
        /// 婉拒筆數
        /// </summary>
        public int Decline_count { get; set; }
        /// <summary>
        /// 待對保筆數
        /// </summary>
        public int Guarantee { get; set; }
        /// <summary>
        /// 不對保筆數
        /// </summary>
        public int GuaranteeNone { get; set; }
    }

    public class SendCaseStatu_Excel
    {
        public string Company_Name { get; set; }
        public int totleCount { get; set; }
        public int ApprCount { get; set; }
        public int unApprCount { get; set; }
        public int PayCount { get; set; }
        public int unPayCount { get; set; }
        public int WPayCount { get; set; }
        public int GUCount { get; set; }
    }

    public class SendCaseStatu_Det_Excel
    {
        public decimal HS_id { get; set; }
        public string show_fund_company { get; set; }
        public string Send_amount_date { get; set; }
        public string CS_name { get; set; }
        public string CS_ID { get; set; }
        public string CS_MTEL1 { get; set; }
        public string show_appraise_company { get; set; }
        public string show_project_title { get; set; }
        public string ShowType { get; set; }
        public string Send_amount { get; set; }
        public string pass_amount { get; set; }
        public string Loan_rate { get; set; }
        public string addr { get; set; }
    }

    public class CS_ListByJob_req
    {
        public DateTime Pre_Agency_Date_S { get; set; }
        public DateTime Pre_Agency_Date_E { get; set; }
        public string job_kind { get; set; }
        public int page { get; set; }
    }

    public class CS_ListByJob_Excel
    {
        public string CS_name { get; set; }
        public string CS_birthday { get; set;}
        public string CS_MTEL1 { get; set;}
        public string CS_company_name { get; set;}
        public string CS_company_tel { get;set;}
        public string job_kind_na { get; set;}
        public string CS_job_title { get; set;}
        public string CS_job_years { get; set;}
        public string income_na { get; set;}
        public string CS_income_everymonth { get;set;}
    }

    public class Approval_Loan_Sales_req 
    {
        public string OrderByStyle { get;set; }
        public string YYYYMM { get; set;}
        public string sales_num { get; set;}
    }

    public class Approval_Loan_Sales_res: Approval_Loan_Sales_Excel
    {
        public string isCancel { get; set; }
        public string Introducer_PID { get; set; }
        public int I_Count { get; set; }
    }

    public class Approval_Loan_Sales_Excel
    {
        public string U_BC_name { get; set; }
        public string Send_amount_date { get; set; }
        public string CS_name { get; set; }
        public string CS_introducer { get; set; }
        public string Bank_name { get; set; }
        public string Bank_account { get; set; }
        public string plan_name { get; set; }
        public string Send_result_date { get; set; }
        public string pass_amount { get; set; }
        public string get_amount_date { get; set; }
        public int get_amount { get; set; }
        public string show_project_title { get; set; }
        public string Loan_rate { get; set; }
        public string interest_rate_original { get; set; }
        public string interest_rate_pass { get; set; }
        public int charge_flow { get; set; }
        public int charge_agent { get; set; }
        public int charge_check { get; set; }
        public int get_amount_final { get; set; }
        public decimal subsidized_interest { get; set; }
        public string Comm_Remark { get; set; }
        /// <summary>
        /// 委對/外對
        /// </summary>
        public string Comparison { get; set; }
    }
    public class Flow_rest_report_req
    {
        public string Flow_Rest_Date_S { get; set; }
        public string Flow_Rest_Date_E { get; set; }
        public string U_BC { get; set; }
        public string U_num { get; set; }
        public string LE_tag { get; set; }
    }

    public class Flow_rest_report_res_early 
    {
        public string U_num { get; set; }
        public int Sum_early { get; set; }  
    }
    public class Flow_rest_report_Excel
    {
        public string U_BC_name { get; set; }
        public string FR_U_num { get; set; }
        public string FR_U_name { get; set; }
        public decimal SUM_FR_kind_FRK001 { get ; set; }
        public decimal SUM_FR_kind_FRK002 { get; set; }
        public decimal SUM_FR_kind_FRK003 { get; set; }
        public decimal SUM_FR_kind_FRK004 { get; set; }
        public decimal SUM_FR_kind_FRK005 { get; set; }
        public decimal SUM_FR_kind_FRK006 { get; set; }
        public decimal SUM_FR_kind_FRK007 { get; set; }
        public decimal SUM_FR_kind_FRK008 { get; set; }
        public decimal SUM_FR_kind_FRK009 { get; set; }
        public decimal SUM_FR_kind_FRK010 { get; set; }
        public decimal SUM_FR_kind_FRK011 { get; set; }
        public decimal SUM_FR_kind_FRK012 { get; set; }
        public decimal SUM_FR_kind_FRK013 { get; set; }
        public decimal SUM_FR_kind_FRK014 { get; set; }
        public decimal SUM_FR_kind_FRK015 { get; set; }
        public decimal SUM_FR_kind_FRK018 { get; set; }
        public decimal SUM_FR_kind_FRK016 { get; set; }
        public decimal SUM_FR_kind_FRK017 { get; set; }
        public decimal SUM_FR_kind_FRK019 { get; set; }
        public decimal SUM_FR_kind_FRK020 { get; set; }
        public decimal SUM_FR_kind_FRK999 { get; set; }
    }
    
}
