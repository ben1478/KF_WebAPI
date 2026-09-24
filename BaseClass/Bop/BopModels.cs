namespace KF_WebAPI.BaseClass.Bop
{

    public class RegisterAccountQRCodeRequest
    {
        public string UUID { get; set; } = Guid.NewGuid().ToString(); // 36 碼唯一序號
        public string VTACCT { get; set; }                           // 16 碼虛擬帳號
        public string CHKREPEAT { get; set; } = "Y";                 // 是否可重複入帳 (Y/N)
        public string CHKSINGLE { get; set; } = "N";                 // 是否檢核單筆金額 (Y/N)
        public string SINGLELIMIT { get; set; } = "0";               // 單筆金額
        public string CHKTOT { get; set; } = "N";                    // 是否檢核總金額 (Y/N)
        public string TOTLIMIT { get; set; } = "0";                  // 總金額
        public string QRCODEAMOUNT { get; set; } = "0";              // QRCode 金額
        public string QRCODENOTE { get; set; } = "";                 // QRCode 備註 (長度上限 20)
        public string CHKDEADLINE { get; set; } = "N";               // 是否檢核付款期限 (Y/N)
        public string DEADLINE { get; set; }                         // 付款期限 (YYYYMMDDhhmm)
        public string RECYCLEDATE { get; set; }                      // 帳號回收日 (YYYYMMDD)
        public string KeySN { get; set; }                            // 私鑰序號
    }

    public class RegisterAccountQRCodeResult
    {
        public string ReturnCode { get; set; }                       // "000000" 代表成功
        public string ReturnMsg { get; set; }
        public string UUID { get; set; }
        public string VTACCT { get; set; }
        public string QRCode { get; set; }                           // 板信回傳的 QRCode 字串
        public string SHA256 { get; set; }
        public string RequestJson { get; set; }
        public string ResponseJson { get; set; }
    }


    public class CreateVirtualAccountRequest
    {
        public string VTACCT { get; set; }//虛擬帳號
        public decimal Amount { get; set; }          // 交易金額
        public string Note { get; set; } = "";       // 備註說明
        public DateTime? Deadline { get; set; }      // 繳費截止時間 (選填)
        public DateTime? RecycleDate { get; set; }   // 帳號回收日期 (選填)
        public bool AllowRepeat { get; set; } = true;// 是否可重複入帳
    }

    public class CreateVirtualAccountResponse
    {
        public bool Success { get; set; }
        public string Message { get; set; }
        public string CustomerId { get; set; }
        public string UUID { get; set; }
        public string VirtualAccount { get; set; }
        public string QRCode { get; set; }
    }


}
