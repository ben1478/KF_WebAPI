using System;

namespace KF_WebAPI.BaseClass.AE
{
    public class Flow_rest 
    {
        public tbInfo? tbInfo { get; set; }
        public string? FR_id { get; set; }
        public string? FR_cknum { get; set; }
        public string? FR_version { get; set; }
        public string? FR_kind { get; set; }
        public DateTime? FR_date_begin { get; set; }
        public DateTime? FR_date_end { get; set; }
        public decimal FR_total_hour { get; set; }
        public string? FR_note { get; set; }
        public string? FR_sign_type { get; set; }
        public string? FR_cancel { get; set; }
        public string? FR_U_num { get; set; }
        public string? FR_step_now { get; set; }
        public string? FR_step_01_type { get; set; }
        public string? FR_step_01_num { get; set; }
        public string? FR_step_01_sign { get; set; }
        public string? FR_step_01_note { get; set; }
        public DateTime? FR_step_01_date { get; set; }
        public string? FR_step_02_type { get; set; }
        public string? FR_step_02_num { get; set; }
        public string? FR_step_02_sign { get; set; }
        public string? FR_step_02_note { get; set; }
        public DateTime? FR_step_02_date { get; set; }
        public string? FR_step_03_type { get; set; }
        public string? FR_step_03_num { get; set; }
        public string? FR_step_03_sign { get; set; }
        public string? FR_step_03_note { get; set; }
        public DateTime? FR_step_03_date { get; set; }
        public string? FR_step_04_type { get; set; }
        public string? FR_step_04_num { get; set; }
        public string? FR_step_04_sign { get; set; }
        public string? FR_step_04_note { get; set; }
        public DateTime? FR_step_04_date { get; set; }
        public string? FR_step_05_type { get; set; }
        public string? FR_step_05_num { get; set; }
        public string? FR_step_05_sign { get; set; }
        public string? FR_step_05_note { get; set; }
        public DateTime? FR_step_05_date { get; set; }
        public string? FR_step_HR_type { get; set; }
        public string? FR_step_HR_num { get; set; }
        public string? FR_step_HR_sign { get; set; }
        public string? FR_step_HR_note { get; set; }
        public DateTime? FR_step_HR_date { get; set; }
        public DateTime? cancel_date { get; set; }
        public string? cancel_num { get; set; }
        public string? cancel_ip { get; set; }
    }

    public class Flow_rest_req
    {
        public int page { get; set; } 
        public DateTime? FR_date_begin { get; set; }
        public DateTime? FR_date_end { get; set; }
        public string? U_BC { get; set; }    
        public string? FR_sign_type { get;set; }
        public string? FR_kind { get; set; }
        public string? Rest_Num { get; set; }
    }

    public class Flow_rest_HR_excel
    {
        public string FR_id { get; set; }
        public string U_name { get; set; }
        public string FR_kind { get; set; }
        public string FR_note { get; set; }
        public string FR_Kind_name { get; set; }
        public DateTime FR_Date_S { get; set; }
        public DateTime FR_Date_E { get; set; }
        public decimal FR_total_hour { get; set; }
        public string Sign_name { get; set; }
        public string? FR_U_num { get; set; }
        public string? FR_ot_compen { get; set; }
    }

}
