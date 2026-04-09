namespace KF_WebAPI.BaseClass.AE
{
    public class BookingRequest
    {
        public int? BookingID { get; set; }
        public string Room { get; set; }
        public string? U_num{ get; set; }
        public DateTime StartTime { get; set; }
        public DateTime EndTime { get; set; }
        public string? Subject { get; set; }
    }
}
