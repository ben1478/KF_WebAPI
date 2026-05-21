namespace KF_WebAPI.BaseClass.AE
{
    public class ACHBatchInput
    {
        // 變動參數：發動日期
        public DateTime LaunchDate { get; set; }

        // 明細資料清單
        public List<ACHDetailInput> Details { get; set; }
    }

    public class ACHDetailInput
    {
        public string AccountNo { get; set; }   // 收受者帳號
        public string CS_PID { get; set; }        // 收受者統編 / 身分證字號
        public string Bank_No { get; set; }      // 提回銀行代碼 (7碼)
        public decimal Amount { get; set; }           // 金額
    }

}
