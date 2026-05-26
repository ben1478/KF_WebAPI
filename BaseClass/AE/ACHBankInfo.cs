namespace KF_WebAPI.BaseClass.AE
{
    public class ACHBankInfo
    {
        //檔案名稱
        public string? FileName { get; set; }
        //客戶PID
        public string? CS_PID { get; set; }
        //銀行代號
        public string? BankNo { get; set; }
        //銀行帳號
        public string? AccountNo { get; set; }
        //成功=1;失敗=0;
        public Boolean IsSuccess { get; set; }
        //錯誤訊息
        public string? ResultMsg { get; set; }
    }
}
