namespace KF_WebAPI.BaseClass.AE
{
    public class Manufacturer
    {

        public int? MF_ID { get; set; }
        public string? MF_cknum { get; set; }
        public string? Company_name { get; set; }
        public string? Company_number { get; set; }
        public string? Company_addr { get; set; }
        public string? Company_busin { get; set; }
        public string? Invoice_Iss { get; set; }
        public string? Overseas { get; set; }
        public DateTime? add_date { get; set; }
        public string? add_num { get; set; }
        public string? add_ip { get; set; }
        public DateTime? edit_date { get; set; }
        public string? edit_num { get; set; }
        public string? edit_ip { get; set; }

    }

    public class Manufacturer_Ins
    {
        public string? MF_ID { get; set; }
        public string? Company_name { get; set; }
        public string? Company_number { get; set; }
        public string? Company_addr { get; set; }
        public string? Company_busin { get; set; }
        public string Company_tel { get; set; }
        public string Company_fax { get; set; }
        public string? Invoice_Iss { get; set; }
        public string? Overseas { get; set; }
        public string? add_num { get; set; }
        public string? edit_num { get; set; }
    }
}
