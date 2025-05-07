namespace KF_WebAPI.BaseClass.AE
{
    public class Pro_Target_Ins
    {
        public string title { get; set; }
        public string startMonth { get; set; }
        public string endMonth { get; set; }
        public int amount { get; set; }
        public string? user { get; set; }
    }

    public class Pro_Target : Pro_Target_Ins
    {
        public string PR_ID { get; set; }
    }
}
