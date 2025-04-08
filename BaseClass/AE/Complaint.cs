using System;

namespace KF_WebAPI.BaseClass.AE
{
    /// <summary>
    /// 客訴資料
    /// </summary>
    public class Complaint_Ins
    {
        public string? Comp_Id { get; set; }
        public string? CS_name { get; set; }
        public string? Sales_num { get; set; }
        public string? Complaint { get; set; }
        public string? CompDate { get; set; }
        public string? CompTime { get; set; }
        public string? Remark { get; set; }
        public DateTime? add_date { get; set; }
        public string? add_num { get; set; }
        public DateTime? edit_date { get; set; }
        public string? edit_num { get; set; }
        public DateTime? del_date { get; set; }
        public string? del_num { get; set; }


    }

    public class Complaint_Req
    {
        
        public string Date_S { get; set; }
        public string Date_E { get; set; }
        
        //部門
        public string U_BC { get; set; }

    }
    
}
