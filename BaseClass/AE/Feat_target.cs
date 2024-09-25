namespace KF_WebAPI.BaseClass.AE
{
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
        public int Target { get;set; }
    }
}
