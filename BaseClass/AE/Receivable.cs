﻿namespace KF_WebAPI.BaseClass.AE
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
    public class Receivable_D
    {
        public tbInfo? tbInfo { get; set; }
        public decimal RCD_id { get; set; }
        public string RCD_cknum { get; set; }
        public decimal RCM_id { get; set; }
        public int RC_count { get; set; }
        public decimal RC_amount { get; set; }
        public DateTime RC_date { get; set; }
        public string? RC_type { get; set; }
        public string? RC_note { get; set; }
        public string check_pay_type { get; set; }
        public DateTime? check_pay_date { get; set; }
        public string? check_pay_num { get; set; }
        public string bad_debt_type { get; set; }
        public DateTime? bad_debt_date { get; set; }
        public string? bad_debt_num { get; set; }
        public string cancel_type { get; set; }
        public DateTime? cancel_date { get; set; }
        public string? cancel_num { get; set; }
        public string? invoice_no { get; set; }
        public DateTime? invoice_date { get; set; }
        public decimal interest { get; set; }
        public decimal Rmoney { get; set; }
        public decimal? RemainingPrincipal { get; set; }
        public decimal PartiallySettled { get; set; }
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
        public string str_Date_S { get; set; }
        public string str_Date_E { get; set; }
        public string check_type { get; set; }
        public string? NS_type { get; set; }
        public string bad_type { get; set; }
        public string RC_count { get; set; }
    }

    public class Receivable_res
    {
        public decimal RCM_id { get; set; }
        public decimal RCD_id { get; set; }
        public string CS_name { get; set; }
        public string str_amount_total { get; set; }
        public int RC_count { get; set; }
        public string RC_date { get; set; }
        public string str_RC_amount { get; set; }
        public string? str_interest { get; set; }
        public string? str_Rmoney { get; set; }
        public string str_RemainingAmount { get; set; }
        public string str_PartiallySettled { get; set; }
        public int? DelayDay { get; set; }
        public string check_pay_type { get; set; }
        public string? check_pay_date { get; set; }
        public string check_pay_name { get; set; }
        public string RC_note { get; set; }
        public string bad_debt_type { get; set; }
        public string? bad_debt_date { get; set; }
        public string bad_debt_name { get; set; }
        public string? invoice_no { get; set; }
        public string? invoice_date { get; set; }
        public string interest_rate_pass { get; set; }
        public string isOver_RC { get; set; }
        public string isNewFun { get; set; }
        public decimal? RemainingPrincipal { get; set; }
        public decimal? Fee { get; set; }
        public decimal? EXrate { get; set; }
        public decimal? Delaymoney { get; set; }
    }
    public class Receivable_Excel
    {

    }
}
