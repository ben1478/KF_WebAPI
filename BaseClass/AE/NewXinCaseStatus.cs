namespace KF_WebAPI.BaseClass.AE
{
    public class NewXinCaseStatus_res
    {
        public string? CS_ID { get; set; }
        public string? addr { get; set; }
        public string? Loan_rate { get; set; }
        public string? Send_amount_date { get; set; }
        public string? get_amount_date { get; set; }
        public string? HS_id { get; set; }
        public string? pass_amount { get; set; }
        public string? Send_amount { get; set; }
        public string? get_amount { get; set; }
        public string? project_title { get; set; }
        public string? CS_name { get; set; }
        public string? CS_MTEL1 { get; set; }
        public string? CS_MTEL2 { get; set; }
        public string? CS_introducer { get; set; }
        public string? plan_name { get; set; }
        public string? U_BC_name { get; set; }
        public string? show_Send_result_type { get; set; }
        public string? show_check_amount_type { get; set; }
        public string? show_get_amount_type { get; set; }
        public string? show_appraise_company { get; set; }
        public string? show_fund_company { get; set; }
        public string? show_project_title { get; set; }
        public string? fin_name { get; set; }
        public DateTime fin_date { get; set; }
    }

    public class NewXinCaseStatus_req
    {
        public string? start_date { get; set; }
        public string? end_date { get; set; }
        /// <summary>
        /// 撥款年月
        /// </summary>
        public string? selYear_S { get; set; }
    }

    public class NewXinCaseStatus_Excel
    {
        public string? Send_amount_date { get; set; }
        public string? get_amount_date { get; set; }
        public string? CS_name { get; set; }
        public string? CS_ID { get; set; }
        public string? show_project_title { get; set; }
        public string? get_amount { get; set; }
    }

    public class NewXinUpload
    {
        public string? A { get; set; }
        public string? B { get; set; }
        public string? C { get; set; }
        public string? D { get; set; }
        public string? E { get; set; }
        public string? F { get; set; }
        public string? G { get; set; }
        public string? H { get; set; }
        public string? I { get; set; }
        public string? J { get; set; }
        public string? K { get; set; }
        public string? L { get; set; }
        public string? M { get; set; }
        public string? N { get; set; }
        public string? O { get; set; }
        public string? P { get; set; }
        public string? Q { get; set; }
        public string? R { get; set; }
        public string? S { get; set; }
        public string? T { get; set; }
        public string? U { get; set; }
    }

}
