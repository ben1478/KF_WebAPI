namespace KF_WebAPI.BaseClass.AE
{
    public class StagnationDebt_M : tbInfo
    {
        public tbInfo? tbInfo { get; set; }
        public int? sdm_id { get; set; }
        public decimal rcm_id { get; set; }
        public decimal? total_due_amount { get; set; }
        public decimal? total_paid_amount { get; set; }
        public decimal? total_bad_debt { get; set; }
        public string remarks { get; set; }
    }

    public class StagnationDebt_Res : StagnationDebt_M
    {
        public string SD_Name { get; set; }
        public string SD_CID { get; set;}
        public string str_total_due_amount { get; set; }
        public string str_total_paid_amount { get; set; }
        public string str_total_bad_debt { get; set; }
        public List<StagnationDebt_D>? SDList { get; set; }
    }

    public class StagnationDebt_D
    {
        public tbInfo? tbInfo { get; set; }
        public int? sdm_id { get; set; }
        public int? sdd_id { get; set; }
        public decimal? payment_amount { get; set; }
        public string? str_payment_amount { get; set; }
        public string? collection_date { get; set; }
        public string? collection_date_roc { get; set; }
        public string? collection_not { get; set; }
    }

}
