namespace KF_WebAPI.BaseClass.AE
{
    public class AuditFlow_Ins
    {
        public string AF_ID { get; set; }
        public string AF_Name { get; set; }
        public string AF_Caption { get; set; }
        public string AF_Step { get; set; }
        public string AF_Step_Caption { get; set; }
        public string AF_Step_Person { get; set; }
        public string add_num { get; set; }
        public string edit_num { get; set; }

    }

    public class AuditFlow_M_Req
    {
        public string AF_ID { get; set; }
        public string FM_Source_ID { get; set; }
    }

    public class RevFlow_Req
    {
        public string? U_BC { get; set; }
        public string? User { get; set; }
        public string? RF_Date_S { get; set; }
        public string? RF_Date_E { get; set; }
    }

    public class AuditFlow_D_Upd
    {
        public tbInfo? tbInfo { get; set; }
        public string? FM_Source_ID { get; set; }
        public string? FD_Sign_Countersign { get; set; }
        public string? FD_Step_SignType { get; set; }
        public string? User { get; set; }
        public string? FD_Step_desc { get; set; }
        public string? FM_Step { get; set; }
    }

    public class AuditFlowReason
    {
        public string FD_ID { get; set; }
        public string Reason { get; set; }
        public string FD_Step_num { get; set; }
        public string User { get; set; }
    }

    public class Counter_Ins
    {
        public string[] arr_Unm { get; set; }
        public string AF_ID { get; set; }
        public string FM_Step { get; set; }
        public string FM_Source_ID { get; set; }
        public string User { get; set; }
    }

    public class LF_AF_Confirm
    {
        public string Source_ID { get; set; }
        public string FM_Step { get; set; }
        public string Confirm { get; set; }
        public string User { get; set; }
    }
}
