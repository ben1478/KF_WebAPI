namespace KF_WebAPI.BaseClass.AE
{
    public class UserGroup
    {
        public string U_num { get; set; }
    }
    public class Feat_user_list_req
    {
        public string start_date { get; set; }
        public string end_date { get; set;}

        public string U_num { get; set;}
    }
    public class Feat_user_list_res
    {
        public string CS_name {  get; set; }
        public decimal HS_id { get; set; }
        public string plan_name { get; set;}
        public string show_appraise_company { get; set;}
        public string show_project_title { get;set;}
        public DateTime get_amount_date { get; set; }
        public string get_amount { get; set; }
        public string Loan_rate { get; set; }  
        public string interest_rate_original { get; set; }
        public string interest_rate_pass { get;set; }
        public string charge_M { get;set; }
        public string charge_flow { get;set; }
        public string charge_agent { get;set; }
        public string charge_check { get;set; }
        public string get_amount_final { get;set; }
        public string CS_introducer { get; set; }
        public string HS_note { get; set; }
        public string project_title { get; set;}
        public string exception_type { get; set;}
        public string exception_rate { get;set; }
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
        public string? range_D_ratio_B { get;set; }
        public string? range_D_rate { get; set; }
        public string? range_D_reward { get; set; }
        public int? total_get_amount { get; set; }
        public decimal show_cash { get; set; }

    }
}
