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
        public string? U_BC { get; set; }
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

        public DateTime? arrive_date { get; set; }
        public DateTime? leave_date { get; set; }
        public int item_sort { get; set; }
        public string Role_num { get; set; }
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

    public class DayWeek
    {
        public string ymd { get; set; }
        public string strweek { get; set; }
        /// <summary> 休假日類別 </summary>
        public string typename { get; set; }
    }

    public class Attendance_report_excel
    {
        public string? U_Na { get; set; }
        public string? userID { get; set; }
        public string? user_name { get; set; }
        public string? attendance_date { get; set; }
        public string? attendance_week { get; set; }
        public string? work_time { get; set; }
        public string? getoffwork_time { get; set; }
        public int? Late { get; set; }
        public int? early { get; set; }
        public string? U_BC { get; set; }
        /// <summary> 休假日類別 </summary>
        public string? type { get; set; }
        public string typename { get; set; }
        /// <summary> 曠職判定 </summary>
        public string? absenteeism { get; set; }
        public DateTime? arrive_date { get; set; }
        public DateTime? leave_date { get; set; }

        public int item_sort { get; set; }
        public string Role_num { get; set; }
    }
}
