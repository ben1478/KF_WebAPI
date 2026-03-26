using System;

namespace KF_WebAPI.BaseClass.AE
{
    /// <summary>
    /// 客訴資料
    /// </summary>
    public class Complaint_M
    {
        public tbInfo? tbInfo { get; set; }
        public string? Comp_Id { get; set; }
        public string? CS_name { get; set; }
        public string? CS_PID { get; set; }
        public string? PassID { get; set; }
        public string? Sales_num { get; set; }
        public string? ContractID { get; set; }
        public string? CompSou { get; set; }
        public string? Complaint { get; set; }
        public string? CompDate { get; set; }
        public string? CompTime { get; set; }
        public string? Remark { get; set; }
        public string? RiskLevel { get; set; }
        public string? ResponsBC { get; set; }
        public string? ResponsWay { get; set; }
        public string? ResponsStates { get; set; }
        public string IsClose { get; set; }
        public string? CloseDate { get; set; }
        public int? DealDay { get; set; }
    }

    public class Complaint_M_req
    {
        public string RoleType { get; set; }
        public string? CheckDateS { get; set; }
        public string? CheckDateE { get; set;}
        public string? CS_name { get; set; }
        public string? CS_PID { get; set; }
        public string? Sales_num { get; set; }
        public string? RiskLevel { get;set; }
        public string? CompSou { get; set; }
        public string? IsClose { get;set; }
        public string U_BC { get;set; }
        public string User { get; set; }    
    }
    public class Complaint_Req
    {   
        public string? BC_code { get; set; }     
        public string? selYear_S { get; set; }   
        public string? UserNum { get; set; }      
        public string? UserRole { get; set; }     
        public string? UserBC { get; set; } 
    }
}
