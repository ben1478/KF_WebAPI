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

    //以下待修正
    public class AuditFlow_D_Ins
    {
        public tbInfo? tbInfo { get; set; }
        public string? FD_Source_ID { get; set; }
        public string? FM_Step { get; set; }
        public string? FM_Step_Sign { get; set; }
        public string? FD_Step1_Desc { get; set; }
        public string? FD_Step1_num { get; set; }
        public DateTime? FD_Step1_date { get; set; }
        public string? FD_Step1_SignType { get; set; }
        public string? FD_Step2_Desc { get; set; }
        public string? FD_Step2_num { get; set; }
        public DateTime? FD_Step2_date { get; set; }
        public string? FD_Step2_SignType { get; set; }
        public string? FD_Step3_Desc { get; set; }
        public string? FD_Step3_num { get; set; }
        public DateTime? FD_Step3_date { get; set; }
        public string? FD_Step3_SignType { get; set; }
        public string? FD_Step4_Desc { get; set; }
        public string? FD_Step4_num { get; set; }
        public DateTime? FD_Step4_date { get; set; }
        public string? FD_Step4_SignType { get; set; }
        public string? FD_Step5_Desc { get; set; }
        public string? FD_Step5_num { get; set; }
        public DateTime? FD_Step5_date { get; set; }
        public string? FD_Step5_SignType { get; set; }
        public string? FD_Step6_Desc { get; set; }
        public string? FD_Step6_num { get; set; }
        public DateTime? FD_Step6_date { get; set; }
        public string? FD_Step6_SignType { get; set; }
        public string? FD_Step7_Desc { get; set; }
        public string? FD_Step7_num { get; set; }
        public DateTime? FD_Step7_date { get; set; }
        public string? FD_Step7_SignType { get; set; }
        public string? FD_Step8_Desc { get; set; }
        public string? FD_Step8_num { get; set; }
        public DateTime? FD_Step8_date { get; set; }
        public string? FD_Step8_SignType { get; set; }
        public string? FD_Step9_Desc { get; set; }
        public string? FD_Step9_num { get; set; }
        public DateTime? FD_Step9_date { get; set; }
        public string? FD_Step9_SignType { get; set; }
    }

    public class AuditFlow_Req
    {
        public string FM_ID { get; set; }
        public string FD_Source_ID { get; set;}
    }

    public class Flow_Req
    {
        public string FM_ID { get; set; }
        public string FD_Source_ID { get; set; }
        public string FM_Step_Now { get; set; }
        public string FM_Step { get; set; }
        public string FD_step_sign { get; set; }
        public string FD_step_note { get; set; }
        public string PM_U_num { get; set; }
        public string Msg_kind { get; set; }
    }

    public class AuditFlowDetail_Ins
    {
        public string edit_num { get; set; }
        public string FM_ID { get; set; }
        public string FD_Source_ID { get; set; }
        public string? FD_Step1_Desc { get; set; }  
        public string? FD_Step1_num { get; set; }
        public string? FD_Step2_Desc { get; set; }
        public string? FD_Step2_num { get; set; }
        public string? FD_Step3_Desc { get; set; }
        public string? FD_Step3_num { get; set; }
        public string? FD_Step4_Desc { get; set; }
        public string? FD_Step4_num { get; set; }
        public string? FD_Step5_Desc { get; set; }
        public string? FD_Step5_num { get; set; }
        public string? FD_Step6_Desc { get; set; }
        public string? FD_Step6_num { get; set; }
        public string? FD_Step7_Desc { get; set; }
        public string? FD_Step7_num { get; set; }
        public string? FD_Step8_Desc { get; set; }
        public string? FD_Step8_num { get; set; }
        public string? FD_Step9_Desc { get; set; }
        public string? FD_Step9_num { get; set; }
    }

    public class AuditFlowReason
    {
        public string FM_ID { get; set; }
        public string FD_Source_ID { get; set; }
        public string Type { get; set; }
        public string Reason { get; set; }
        public string User { get; set; }
    }
}
