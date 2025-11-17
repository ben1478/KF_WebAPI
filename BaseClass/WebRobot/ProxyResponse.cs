namespace KF_WebAPI.BaseClass.WebRobot
{
    public class ProxyResponse
    {
        public int code { get; set; }
        public List<string> data { get; set; }
        public string msg { get; set; }
        public bool success { get; set; }
    }
}
