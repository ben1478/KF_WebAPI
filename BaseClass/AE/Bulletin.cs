namespace KF_WebAPI.BaseClass.AE
{
    public class Bulletin_Ins
    {
        public tbInfo? tbInfo { get; set; }
        public int id { get; set; } 
        public string notice_date { get; set; }
        public string notice_type { get; set; }
        public string notice_mode { get; set; }
        public string title { get; set; }
        public string bulletin_content { get; set; }
    }
}
