namespace KF_WebAPI.BaseClass.AE
{
    public class Message
    {
        public string Name { get; set; }
        public string MsgID { get; set; }
        public int Count { get; set; }
    }

    public class Message_Req
    {
        public string StartDate { get; set; }
        public string EndDate { get; set; }
        public string Mes_Kind { get; set; }
        public string ReadType { get; set; }
        public string UserType { get; set; }
        public string User { get;set; }
    }

}
