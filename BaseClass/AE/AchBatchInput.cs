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
        public string ReceiverAccount { get; set; }   // 收受者帳號
        public string ReceiverId { get; set; }        // 收受者統編 / 身分證字號
        public string RecvBankCode { get; set; }      // 提回銀行代碼 (7碼)
        public decimal Amount { get; set; }           // 金額
        public string UserNumber { get; set; }        // 用戶號碼
        public string OriginatorZone { get; set; }    // 發動者專用區 (選填，對帳用)
    }

}
