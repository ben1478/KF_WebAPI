namespace KF_WebAPI.BaseClass.AE
{
    public class BookingRequest
    {
        public int RoomId { get; set; }
        public string? U_num{ get; set; }
        public DateTime StartTime { get; set; }
        public DateTime EndTime { get; set; }
        public string? Subject { get; set; }
    }
}
