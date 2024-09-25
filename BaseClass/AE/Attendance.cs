using System.Collections.Generic;

namespace KF_WebAPI.BaseClass.AE
{
    public class Attendance
    {
        public int? id { get; set; }
        public string? user_name { get; set; }
        public string? userID { get; set; }
        public string? yyyymm { get; set; }
        public string? user_apart { get; set; }
        public string? attendance_date { get; set; }
        public string? work_time { get; set; }
        public string? getoffwork_time { get; set; }
        public DateTime? inputdate { get; set; }
        public string? user_num { get; set; }
        public string? ac_flag { get; set; }

    }

    public class Attendance_req 
    {
        public string yyyymm { get; set; }
        public int AttStatus { get; set;}
        public string? U_num { get;set; }
        public string? U_BU { get; set; }
        public string? U_name { get;set; }

    }

    public class Attendance_res
    {
        public string? U_Na { get; set;}
        public string? userID { get; set; }
        public string? user_name { get; set; }
        public string? attendance_date { get; set; }
        public string? work_time { get; set; }
        public int? Late { get; set; }
        public string? work_status { get; set; }
        public string? getoffwork_time { get; set; }
        public int? early { get; set; }
        public string? offwork_status { get; set; }
        public string? U_BC { get; set; }
        public int? RestCount { get; set; }
        public string? FR_sign_type_name_desc { get; set; }
        public decimal? FR_total_hour { get; set; }

    }

    public class Attendance_res_excel
    {
        public string? U_Na { get; set; }
        public string? userID { get; set; }
        public string? user_name { get; set; }
        public string? attendance_date { get; set; }
        public string? work_time { get; set; }
        public int? Late { get; set; }
        public string? work_status { get; set; }
        public string? getoffwork_time { get; set; }
        public int? early { get; set; }
        public string? offwork_status { get; set; }
        public string? FR_sign_type_name_desc { get; set; }
        public decimal? FR_total_hour { get; set; }

    }
}
