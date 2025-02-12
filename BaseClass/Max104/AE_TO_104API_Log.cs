namespace KF_WebAPI.BaseClass.Max104
{
    public class AE_TO_104API_Log
    {
        public string? ACCESS_TOKEN { get; set; }
        public string? APIName { get; set; }
        public string? JSON { get; set; }
        public string? SESSION_KEY { get; set; }
        
        public string? Add_User { get; set; }

        public string? Result { get; set; }
        public string? Msg { get; set; }

        // 簡化構造函數
        public AE_TO_104API_Log() { }
    }
}
