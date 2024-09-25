namespace KF_WebAPI.BaseClass.AE
{
    public class MonthQuota_res
    {
        public decimal group_M_id { get; set; }
        public string group_M_name { get; set; }
        public string group_D_name { get; set; }
        public string group_D_code { get; set; }
        public int U_PFT_sort { get; set; }
        public string U_PFT_name { get; set; }
        public int target_quota { get; set;}
        public string target_ym { get; set; }

    }
    public class FeatDailyPerson_req
    {
        public string reportDate_n { get; set; }
        public string company { get; set; }
        public string U_BC { get; set;}
    }
    public class FeatDailyReport_req
    {
        public string reportDate_n { get; set; }
        public string U_BC { get; set; }
    }
    public class FeatDailyReport_excel
    {
        public string plan_name { get; set; } 
        public string U_PFT_name { get; set;}
        public int day_incase_num_FDCOM001 { get; set;}
        public int month_incase_num_FDCOM001 { get;set; }
        public int day_incase_num_FDCOM003 { get; set; }
        public int month_incase_num_FDCOM003 { get; set; }
        public int day_get_amount_num { get;set; }
        public int day_get_amount { get;set; }
        public int month_pass_num { get; set; }
        public int month_get_amount_num { get; set; }
        public int month_get_amount_FDCOM001 { get; set; }
        public int month_get_amount_FDCOM003 { get; set; }
        public int month_pass_amount_FDCOM001 { get; set; }
        public int month_pre_amount_FDCOM001 { get; set; }
        public int month_pass_amount_FDCOM003 { get; set; }
        public decimal advance_payment_AE { get; set; }
        public int target_quota { get; set; }
        public string AchievementRate { get;set; }
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
}
