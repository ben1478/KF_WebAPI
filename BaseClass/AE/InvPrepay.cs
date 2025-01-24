namespace KF_WebAPI.BaseClass.AE
{
    public class InvPrepay_Ins
    {
        public tbInfo? tbInfo { get; set; }
        public List<InvPrepay_D_Ins>? Ins_List { get; set; }
        public string VP_BC { get; set; }
        public string[]? VP_type { get; set; }
        public string VP_Pay_Type { get; set; }
        public string VP_Total_Money { get; set; }
        public string bank_code { get; set; }
        public string bank_name { get; set; }
        public string branch_name { get; set; }
        public string bank_account { get; set; }
        public string payee_name { get; set; }
        public string User { get; set; }
    }

    public class InvPrepay_D_Ins
    {
        public string FormID { get; set; }
        public string FormCaption { get; set; }
        public string FormMoney { get; set; }
        public string ChangeReason { get; set; }
    }
}
