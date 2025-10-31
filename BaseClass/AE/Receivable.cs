namespace KF_WebAPI.BaseClass.AE
{
    public class Receivable_M
    {
        public tbInfo? tbInfo { get; set; }
        public decimal RCM_id { get; set; }
        public string? RCM_cknum { get; set; }
        public decimal? HS_id { get; set; }
        public decimal? HA_id { get; set; }
        public decimal? amount_total { get; set; }
        public int? month_total { get; set; }
        public decimal? amount_per_month { get; set; }
        public DateTime? date_begin { get; set; }
        public string? str_date_begin { get; set; }
        public string? RCM_note { get; set; }
        public int? loan_grace_num { get; set; }
        public string? cust_rate { get; set; }
        public decimal? cust_amount_per_month { get; set; }
        public string? court_sale { get; set; }

        public string? interest_rate_pass { get; set; }
    }
    public class Receivable_D_check_req
    {
        public decimal RCD_id { get; set; }
        public string check_pay_type { get; set; }
        public int PartiallySettled { get; set; }
        public string str_check_pay_date { get; set; }
        public string check_pay_num { get; set; }
        public string invoice_no { get; set; }
        public string str_invoice_date { get; set; }
        public string RC_note { get; set; }
        /// <summary>1:沖銷 2:呆帳</summary>
        public string Type { get; set; }
        public string bad_debt_type { get; set; }
        public string str_bad_debt_date { get; set; }
    }
    public class Receivable_MD_req
    {
        public decimal RCM_id { get; set; }
        public decimal amount_per_month { get; set; }
        public string str_date_begin { get; set; }
        public string? RCM_note { get; set; }
    }
    public class Receivable_Settle_req
    {
        public decimal RCM_id { get; set; }
        public string? court_sale { get; set; }
        public string str_begin_settle { get; set; }
        public string RCM_note { get; set; }
    }
    public class Receivable_req
    {
        public string name { get; set; }
        public string RC_Date_S { get; set; }
        public string RC_Date_E { get; set; }
        public string? check_type { get; set; }
        public string? NS_type { get; set; }
        public string? bad_type { get; set; }
        public string? RC_count { get; set; }
        public string? invoice_no { get; set; } 
        public string? invoice_date_S { get; set; }
        public string? invoice_date_E { get; set; }
        public string? pay_date_S { get; set; }
        public string? pay_date_E { get; set; }
    }
    public class Receivable_res
    {
        public decimal RCM_id { get; set; }
        public decimal RCD_id { get; set; }
        public string CS_name { get; set; }
        public string CS_PID { get; set; }
        public decimal amount_total { get; set; }
        public int month_total { get; set; }
        public int RC_count { get; set; }
        public string roc_RC_date { get; set; }
        public decimal amount_per_month { get; set; }
        public decimal interest { get; set; }
        public decimal Rmoney { get; set; }
        public int? HFees { get; set; }
        public decimal Ex_RemainingPrincipal { get; set; }
        public string RecPayAmt { get; set; }
        public string diffAMT { get; set; }
        public int? DelayDay { get; set; }
        public decimal? Delaymoney { get; set; }
        public string RC_Note { get; set; }
        public string check_pay_type { get; set; }
        public string? check_pay_date { get; set; }
        public string check_pay_name { get; set; }
        public string bad_debt_type { get; set; }
        public string? bad_debt_date { get; set; }
        public string bad_debt_name { get; set; }
        public string? invoice_no { get; set; }
        public string? invoice_date { get; set; }
    }
    public class Receivable_Excel
    {
        public string CS_name { get; set; }
        public decimal amount_total { get; set; }
        public int month_total { get; set; }
        public int RC_count { get; set; }
        public string RC_date { get; set; }
        public decimal RC_amount { get; set; }
        public decimal? interest { get; set; }
        public decimal? Rmoney { get; set; }
        public decimal RemainingAmount { get; set; }
        public int PartiallySettled { get; set; }
        public int? DelayDay { get; set; }
        public decimal? Delaymoney { get; set; }
        public string check_pay_type { get; set; }
        public string? check_pay_date { get; set; }
        public string check_pay_name { get; set; }
        public string RC_note { get; set; }
        public string bad_debt_type { get; set; }
        public string? bad_debt_date { get; set; }
        public string bad_debt_name { get; set; }
        public string? invoice_no { get; set; }
        public string? invoice_date { get; set; }
    }
    public class Receivable_Coll_req
    {
        public string name { get; set; }
        public string str_Date_S { get; set; }
        public string str_Date_E { get; set; }
        public string DiffDay_Type { get; set; }  
    }
    public class Receivable_Coll_res
    {
        public string CS_name { get; set; }
        public decimal amount_total { get; set; }
        public int month_total { get; set; }
        public int RC_count { get; set; }
        public string RC_date { get; set; }
        public int DiffDay { get; set; }
        public decimal RC_amount { get; set; }
        public string interest_rate_pass { get; set; }
        public int loan_grace_num { get; set; }
        public decimal interest { get;set; }
        public decimal Rmoney { get;set; }
        /// <summary>
        /// 本金餘額
        /// </summary>
        public decimal RemainingPrincipal { get; set; }
        /// <summary>
        /// 實際本金餘額
        /// </summary>
        public decimal RemainingPrincipal_1 { get; set; }
    }
    public class Receivable_Coll_Excel
    {
        public string CS_name { get; set; }
        public decimal amount_total { get; set; }
        public int month_total { get; set; }
        public int RC_count { get; set; }
        public string RC_date { get; set; }
        public int DiffDay { get; set; }
        public decimal RC_amount { get; set; }
        public decimal interest { get; set; }
        public decimal Rmoney { get; set; }
        /// <summary>
        /// 本金餘額
        /// </summary>
        public decimal RemainingPrincipal { get; set; }
        /// <summary>
        /// 實際本金餘額
        /// </summary>
        public decimal RemainingPrincipal_1 { get; set; }
    }
    public class Receivable_Late_Pay_req
    {
        public string name { get; set; }
        public string str_Date_E { get; set; }
        public string delay_type { get; set; }
        public string pay_type { get; set; }
    }
    public class Receivable_Late_Pay_res
    {
        public decimal RCM_id { get; set; }
        public decimal RCD_id { get; set; }
        public string CS_name { get; set; }
        public decimal amount_total { get; set; }
        public int month_total { get; set; }
        public int RC_count { get; set; }
        public string RC_date { get; set; }
        public decimal RC_amount { get; set; }
        public decimal interest { get; set; }
        public decimal Rmoney { get; set; }
        public decimal? RemainingPrincipal { get; set; }
        public int DelayDay { get; set; }
        public string interest_rate_pass { get; set; }
        public int loan_grace_num { get; set; }
    }
    public class Receivable_Late_Pay_Excel
    {
        public string CS_name { get; set; }
        public decimal amount_total { get; set; }
        public int month_total { get; set; }
        public int RC_count { get; set; }
        public string RC_date { get; set; }
        public decimal RC_amount { get; set; }
        public decimal interest { get; set; }
        public decimal Rmoney { get; set; }
        /// <summary>
        /// 本金餘額
        /// </summary>
        public decimal RemainingPrincipal { get; set; }
        public int DelayDay { get; set; }
    }
    public class Receivable_Over_Rel_Excel
    {
        public string BC_name { get; set; }
        public string u_name { get; set; }
        public int ToT_Count { get; set; }
        public decimal amount_total { get; set; }
        public int OV_Count { get; set; }
        public decimal OV_total { get; set; }
        public string OV_Rate { get; set; }
    }
    public class Receivable_ROC_YYYMM_SE
    {
        public List<Receivable_ROC_YYYMM_S> ROC_Date_S { get; set; }
        public List<Receivable_ROC_YYYMM_E> ROC_Date_E { get; set; }
    }
    public class Receivable_ROC_YYYMM_S
    {
        public string ROC_YYYMM { get; set; }
        public string Gre_YYYYMM { get;set; }
    }
    public class Receivable_ROC_YYYMM_E
    {
        public string ROC_YYYMM { get; set; }
        public string Gre_YYYYMM { get; set; }
    }
    public class Receivable_Repay_res
    {
        public string RC_date { get; set; }
        public string yyyMM { get; set; }
        public int ToCount { get; set; }
        public int YCount { get; set; }
        public int NCount { get; set; }
        public int BCount { get; set; }
        public int SCount { get; set; }
        public string str_interest_T { get; set; }
        public string str_interest { get; set; }
        public string str_interest_U { get; set; }
        public string str_Rmoney_T { get; set; }    
        public string str_Rmoney { get; set; }
        public string str_Rmoney_U { get; set; }
        public string str_RemainingPrincipal_BB { get; set; }
        public string str_S_AMT { get; set; }
        public string str_RemainingPrincipal { get; set; }
    }
    public class Receivable_Repay_Excel
    {
        public string yyyMM { get; set; }
        public int ToCount { get; set; }
        public int YCount { get; set; }
        public int NCount { get; set; }
        public int BCount { get; set; }
        public int SCount { get; set; }
        public decimal interest_T { get; set; }
        public decimal interest { get; set; }
        public decimal interest_U { get; set; }
        public decimal Rmoney_T { get; set; }
        public decimal Rmoney { get; set; }
        public decimal Rmoney_U { get; set; }
        public decimal RemainingPrincipal_BB { get; set; }
        public decimal S_AMT { get; set; }
        public decimal RemainingPrincipal { get; set; }
    }
    public class Receivable_Excess_Excel
    {
        public string diffType { get; set; }
        public string AmtTypeDesc { get; set; }
        public int Count { get; set; }
        public decimal amount_total { get; set; }
        public decimal Rate { get; set; }
    }
    public class Receivable_Excess_req
    {
        public string Forec { get; set; }
        public string DiffType { get; set; }
        public string AmtType { get; set; }
    }
    public class Receivable_Excess_Detail_Excel
    {
        public string diffType { get; set; }
        public string AmtTypeDesc { get; set; }
        public string Cs_name { get; set; } 
        public int DiffDay { get; set; }
        public decimal amount_total { get; set; }
        public string RCM_note { get;set; }
    }
    public class Receivable_Info_res 
    {
        public decimal RCM_id { get; set; }
        public string pay_type { get; set; }
        public string pay_date { get; set; }
        public int pay_money { get; set; }
        public string pay_text { get; set; }
        public string User { get; set; }
    }
    public class Receivable_Over
    {
        public string? u_name { get; set; }
        public string? plan_num { get; set; }
        public string? BC_name { get; set; }    
        public string? U_BC { get; set; }   
        public string? amount_type { get; set; }
        public string? pro_name { get; set; }
        public int TOT_Count { get; set; }
        public decimal amount_total { get; set; }
        public int OV_Count { get; set; }
        public decimal OV_total { get; set; }
        public decimal OV_Rate { get; set; }
        public int SCount { get; set; }
        public decimal RemainingPrincipal { get; set; }

        public int TOT_bad_Count { get; set; }
        public decimal TOT_bad_debt { get; set; }

        public decimal OV_bad_Rate { get; set; }
        public decimal TOT_total { get; set; }
        public decimal TOT_OV_Rate { get; set; }
    }
}
