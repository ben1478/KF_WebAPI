namespace KF_WebAPI.BaseClass.AE
{
    public class User_Hday
    {
        public tbInfo? tbInfo { get; set; }
        public decimal? UH_id { get; set; }
        public string? UH_cknum { get; set; }
        public string? U_num { get; set; }
        public string? H_year_count { get; set; }
        public string? H_year_count_note { get; set; }
        public DateTime H_begin { get; set; }
        public DateTime H_end { get; set; }
        public decimal? H_day_base { get; set; }
        public decimal? H_day_adjust { get; set; }
        public string? H_day_adjust_note { get; set; }
        public decimal? H_day_total { get; set; }
        public int? H_spent_count { get; set; }
    }

    public class User_Hday_res
    {
        public decimal? H_day_total_now { get; set; }
        public DateTime? H_begin_now { get; set; }
        public DateTime? H_end_now { get; set; }
        /// <summary>
        /// 本年度特休時數
        /// </summary>
        public decimal? FR_kind_FRK005_now { get;set; }
        public decimal? H_day_total_last { get; set; }
        public DateTime? H_begin_last { get; set; }
        public DateTime? H_end_last { get; set; }
        /// <summary>
        /// 前年度特休時數
        /// </summary>
        public decimal? FR_kind_FRK005_last { get; set; }
    }

    public class User_Hday_req
    {
        public int page { get; set; }
        public string? U_Num_Name { get; set; }
        public string? U_BC { get; set; }
        public string? Job_Status { get; set; }
    }

    public class User_Hday_Upd
    {
        public tbInfo? tbInfo { get; set; }
        public string? U_num { get; set; }
        public string? H_year_count { get; set; }
        public string? str_H_begin { get; set; }
        public string? str_H_end { get; set; }
        public decimal? H_day_base { get; set; }
        public decimal? H_day_adjust { get; set; }
        public decimal? H_hours { get; set; }
        public string? H_day_adjust_note { get; set; }
        public decimal? H_day_total { get; set; }
    }
}
