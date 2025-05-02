namespace KF_WebAPI.BaseClass.AE
{
    public class Introducer_Comm_req
    {
        public string? Introducer { get; set; }
        public string? isCompany { get; set; }
    }

    public class Introducer_Comm_res
    {
        public string? U_ID { get; set; }
        public string? ID { get; set; }
        public string? Introducer_name { get; set; }
        public string? Introducer_name1 { get; set; }
        public string? Introducer_HBD { get; set; } 
        public string? Introducer_PID { get; set; }
        public string? Bank_account { get; set; }
        public string? Bank_head { get; set; }
        public string? Bank_branches { get; set; }
        public string? Bank_name { get; set; }
        public string? Introducer_addr { get; set; }
        public string? Remark { get; set; }
        public string? Introducer_tel { get; set; }
        public string? isCompany { get; set; } 
        public string? PW_Line { get; set; }
        public string? LINE_Num { get; set; }
        public string? CompanyKey { get; set; }
        public string? Contract_Date { get; set; } 
        public string? add_date { get; set; } 
        public string? add_num { get; set; }
        public string? add_ip { get; set; }
        public string? edit_date { get; set; } 
        public string? edit_num { get; set; }
        public string? edit_ip { get; set; }
        public string? del_tag { get; set; } // 假設為布林值
        public string? del_date { get; set; } // 假設為日期
        public string? del_num { get; set; }
        public string? del_ip { get; set; }
    }

    public class Introducer_Comm_Del
    {
        public string? del_tag { get; set; } 
        public string? del_date { get; set; } 
        public string? del_num { get; set; }
        public string? del_ip { get; set; }
        public string? U_ID { get; set; }

    }

    public class Introducer_Comm_Excel
    {
        public string? Introducer_name { get; set; }
        public string? Introducer_HBD { get; set; }
        public string? Introducer_PID { get; set; }
        public string? Bank_account { get; set; }
        public string? Bank_head { get; set; }
        public string? Bank_branches { get; set; }
        public string? Bank_name { get; set; }
        public string? Introducer_addr { get; set; }
    }


}
