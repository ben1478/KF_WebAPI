namespace KF_WebAPI.BaseClass.WebRobot
{
    public class RemoteAction
    {
        public string? ReStart { get; set; }//重新啟動
        public string? Shutdown { get; set; }//關閉
        public List<string>? UpdRunTimes { get; set; }//更新執行時間
        public List<string>? UpdKeyWords { get; set; }//更新關鍵字List
    }
}
