namespace KF_WebAPI.BaseClass.AE
{
    public class User_group
    {
        public tbInfo? tbInfo { get; set; }
        public string? group_M_title { get; set; }
        public string? group_sort { get; set; }
        public string str_group_start_day { get; set; }
        public string str_group_end_day { get; set; }
        public string group_M_code { get; set; }
        public string group_M_name { get; set; }
    }

    public class User_group_Ins
    {
        public tbInfo? tbInfo { get; set; }
        public int group_id { get; set; }   
        public string? group_M_title { get; set; }
        public string? group_sort { get; set; }
        public string str_group_start_day { get; set; }
        public string str_group_end_day { get; set; }
        public string group_M_code { get; set; }
        public string group_M_name { get; set; }
    }

    public class User_group_Upd : User_group_Ins
    {

    }

    public class User_group_DUpd
    {
        public decimal group_M_id { get; set; }
        public string User { get; set; }
        public decimal group_id { get; set; }
        public string group_M_code { get; set; }
        public string group_M_name { get;set; }
        public string group_D_code { get; set;}
        public string group_D_name { get; set; }
        public string group_start_day { get; set; }
        public string group_end_day { get; set; }
    }
}
