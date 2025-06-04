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

}
