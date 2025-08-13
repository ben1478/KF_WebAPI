namespace KF_WebAPI.BaseClass.AE
{
    public class Role_M
    {
        public string R_name { get; set; }
        public string User { get;set; }
        public string? del_tag { get; set; }
        public string? R_num { get; set; }
    }

    public class MS_PerMission
    {
        public decimal? Map_id { get; set; }
        public decimal menu_id { get; set; }
        public string? R_num { get; set; }
        public string? menu_name { get; set; }
        public string per_read { get; set;}
        public string per_add { get; set; }
        public string per_edit { get; set; }
        public string per_del { get; set; }
        public string per_query { get; set; }
        public string User { get; set; }
    }

}
