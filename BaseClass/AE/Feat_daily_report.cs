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
    public class FeatDailyReport_res 
    {
        public string reportDate_n { get; set; }
        public string U_BC { get; set; }
    }
}
