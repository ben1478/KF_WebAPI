namespace KF_WebAPI.BaseClass.AE
{
    public class Pro_Target_Ins
    {
        public string PR_title { get; set; }
        public string PR_Date { get; set; }
        public int PR_target { get; set; }
        public string? user { get; set; }
    }

    public class Pro_Target : Pro_Target_Ins
    {
        public string PR_ID { get; set; }
    }

    public class Per_Target_Ins
    {
        public string? PE_title { get;set; }
        public string PE_num { get; set; }
        public int PE_target { get; set; }
        public string PE_Date { get; set; } 
        public string? user { get; set; }
    }

    public class Per_Target : Per_Target_Ins
    {
        public int PE_ID { get; set; }
    }

    public class Per_Achieve
    {
        public string month { get;set; }
        public string U_BC_NEW { get; set; }
        public Double total_target { get; set; }   
        public Double total_perf { get; set; }
        public Double total_perf_after_discount { get; set; }
        public string achieve_rate { get; set; }
        public string achieve_rate_after_discount { get; set; }
        public int Subord { get; set; }
        public int Leader { get; set; } 
    }

    public class Per_Target_res
    {
        public int PE_ID { get; set; }
        public string titleName { get; set; }
        public string U_name { get; set; }
        public string PE_num { get; set; }
        public int PE_target { get; set; }
        public string PE_Date { get; set; }
        public string PE_Date_Minguo { get; set; }
    }
}
