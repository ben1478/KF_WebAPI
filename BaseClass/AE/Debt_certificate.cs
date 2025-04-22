namespace KF_WebAPI.BaseClass.AE
{
    public class Debt_Certificate
    {
        public string cs_name { get; set; }
        public string CS_PID { get; set; }
        public decimal loan_amount { get; set; }
        public DateTime certificate_date_S { get; set; }
        public DateTime certificate_date_E { get; set; }
        public string Remark { get; set; }
        public DateTime? add_date { get; set; }
        public string? add_num { get; set; }
        public string? add_ip { get; set; }
        public DateTime? edit_date { get; set; }
        public string? edit_num { get; set; }
        public string? edit_ip { get; set; }
        public string? del_tag { get; set; }
        public DateTime? del_date { get; set; }
        public string? del_num { get; set; }
        public string? del_ip { get; set; }

    }

    public class Debt_Certificate_res
    {
        public string cs_name { get; set; }
        public string CS_PID { get; set; }
        public string? str_loan_amount { get; set; }
        public string str_certificate_date_S { get; set; }
        public string str_certificate_date_E { get; set; }
        public string Remark { get; set; }
        public decimal? loan_amount { get; set; }
    }

    public class Debt_Certificate_Lres
    {
        public int Debt_ID { get; set; }
        public string cs_name { get; set; }
        public string CS_PID { get; set; }
        public string? str_loan_amount { get; set; }
        public string str_certificate_date_S { get; set; }
        public string str_certificate_date_E { get; set; }
        public string Remark { get; set; }
    }

    public class Debt_Certificate_Excel
    {
        public string cs_name { get; set; }
        public string CS_PID { get; set; }
        public decimal loan_amount { get; set; }
        public string str_certificate_date_S { get; set; }
        public string str_certificate_date_E { get; set; }
        public string Remark { get; set; }
    }

    public class Debt_Certificate_req
    {
        public string Debt_ID { get; set; }
        public string cs_name { get; set; }
        public string CS_PID { get; set; }
        public decimal loan_amount { get; set; }
        public string str_certificate_date_S { get; set; }
        public string str_certificate_date_E { get; set; }
        public string Remark { get; set; }

        //public string edit_num { get; set; }

    }
}
