using Microsoft.OpenApi.Writers;

namespace KF_WebAPI.BaseClass.AE
{
    public class InvPrepay_Ins
    {
        public tbInfo? tbInfo { get; set; }
        public List<InvPrepay_D_Ins>? Ins_List { get; set; }
        public string VP_ID { get; set; }   
        public string VP_BC { get; set; }
        public string[]? VP_type { get; set; }
        public string VP_Pay_Type { get; set; }
        public string VP_Total_Money { get; set; }
        public string bank_code { get; set; }
        public string bank_name { get; set; }
        public string branch_name { get; set; }
        public string bank_account { get; set; }
        public string payee_name { get; set; }
        public string VP_MFG_Date { get; set; }
        public string User { get; set; }
    }

    public class InvPrepay_D_Ins
    {
        public string FormID { get; set; }
        public string FormCaption { get; set; }
        public string FormMoney { get; set; }
        public string ChangeReason { get; set; }
        public string? VD_Account_code { get; set; }
        public string? VD_Account { get; set; }
    }

    public class Invp_Print
    {
        public List<Invp_Print_Deatil>? VP_Deatil_List { get; set; }
        public List<Invp_Print_Flow>? VP_Flow_List { get; set; }
        public string BC_Name { get; set; }
        public string U_name { get; set; }
        public string VP_ID { get; set; }
        public string VP_AppDate { get; set; }
        public string VP_Pay_Type { get; set; }
        public string? bank_code { get; set; }
        public string? bank_name { get; set; }
        public string? branch_name { get; set; }
        public string? bank_account { get; set; }
        public string? payee_name { get; set; }
        public string VP_Total_Money { get; set; }
        public string VP_type { get; set; }
        public string? VP_type_PO { get; set; }
        public string add_date { get; set; }
    }

    public class Invp_Print_Deatil
    {
        public string? Form_ID { get; set; }
        public string VD_Fee_Summary { get; set; }
        public string VD_Fee { get; set; }
    }

    public class Invp_Print_Flow
    {
        public string FD_Sign_Countersign { get; set; }
        public string FD_Step { get; set; }
        public string U_name { get; set; }
        public string FD_Step_date { get; set; }
        public string? FD_Step_desc { get; set; }
    }
}
