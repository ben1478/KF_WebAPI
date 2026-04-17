namespace KF_WebAPI.BaseClass.AE
{
    public class Incoming_req
    {
        public string? TelAsk { get; set; }
        public string? TelSour { get; set; }
        public string? Fin_type { get; set; }
        public string? checkDateS { get; set; }
        public string? checkDateE { get; set; }
    }

    public class Motocase_req
    {
        public string? checkDateS { get; set; }
        public string? checkDateE { get; set; }
        public string? project { get; set; }
    }

    public class MotocaseSummary
    {
        public string YYYYMM { get; set; }
        public int SendCount { get; set; }
        public int PassCount { get; set; }
        public int GetCount { get; set; }
        public decimal PassAmount { get; set; }
        public decimal GetAmount { get; set; }
        public decimal PerAmount { get; set; }
        public decimal RemAmount { get; set; }
        public int SettCount { get; set; }
        public decimal SettAmount { get; set; }
        public int BadCount { get; set; }
        public decimal BadAmount { get; set; }
        public string PassRate { get; set; }
        public string GetRate { get; set; }
    }

    public class Carcase_req
    {
        public string? checkDateS { get; set; }
        public string? checkDateE { get; set; }
    }

    public class CarcaseSummary
    {
        public string YYYYMM { get; set; }
        public int SendCount { get; set; }
        public int PassCount { get; set; }
        public int GetCount { get; set; }
        public decimal PassAmount { get; set; }
        public decimal GetAmount { get; set; }
        public decimal PerAmount { get; set; }
        public decimal RemAmount { get; set; }
        public int SettCount { get; set; }
        public decimal SettAmount { get; set; }
        public int BadCount { get; set; }
        public decimal BadAmount { get; set; }
        public string PassRate { get; set; }
        public string GetRate { get; set; }
    }
    public class OverdueRate
    {
        public string CS_name { get; set; }
        public string RC_date { get; set; }
        public string Pro_name { get; set; }
        public int DiffDay { get; set; }
        public decimal Amount_total { get; set; }
        public string DiffType { get; set; }
        
    }

    public class Housecase_req
    {
        public string? checkDateS { get; set; }
        public string? checkDateE { get; set; }
    }

    public class HousecaseSummary
    {
        public string YYYYMM { get; set; }
        public int SendCount { get; set; }
        public int PassCount { get; set; }
        public int GetCount { get; set; }
        public decimal PassAmount { get; set; }
        public decimal GetAmount { get; set; }
        public decimal PerAmount { get; set; }
        public decimal RemAmount { get; set; }
        public int SettCount { get; set; }
        public decimal SettAmount { get; set; }
        public int BadCount { get; set; }
        public decimal BadAmount { get; set; }
        public string PassRate { get; set; }
        public string GetRate { get; set; }
    }

    public class SettDetailList
    {
        public string CS_name { get; set; } 
        public string str_get_amount_date { get; set; }
        public decimal capital_AMT { get; set; }
        public decimal fee_total { get;set; }
        public decimal Interest_total { get; set; }
        public decimal Delay_AMT { get; set; }
        public string get_amount { get; set; }
        public string str_date_begin_settle { get; set; }
        public decimal Sett_total { get; set; } 
        public string pj_name { get; set; }
    }
}
