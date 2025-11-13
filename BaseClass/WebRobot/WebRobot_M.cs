namespace KF_WebAPI.BaseClass.WebRobot
{
    public class WebRobot_M
    {
        public string? ComputerInfo { get; set; }/*電腦資訊*/
        public string? KeyWords { get; set; }/*關鍵字*/
        public Int32? RunTime_S { get; set; }/*執行時間起*/
        public Int32? RunTime_E { get; set; }/*執行時間訖*/
        public Int32? numDelay { get; set; }/*多少秒執行一次*/
        public string? RunStatus { get; set; }/*目前執行狀態*/

    }
}
