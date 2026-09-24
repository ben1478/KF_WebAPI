using System;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using KF_WebAPI.BaseClass;
using KF_WebAPI.BaseClass.Bop;

namespace KF_WebAPI.Service
{
    public class VirtualAccountService
    {
        private readonly BopBankService _bopBankService;
        private readonly IConfiguration _config;

        public VirtualAccountService(BopBankService bopBankService, IConfiguration config)
        {
            _bopBankService = bopBankService;
            _config = config;
        }

        public async Task<CreateVirtualAccountResponse> CreateVirtualAccountAsync(CreateVirtualAccountRequest request)
        {
            // 1. 取得虛擬帳號 (測試階段先讀設定檔，正式階段接資料庫編碼邏輯)
            string vtacct = request.VTACCT;
            if (string.IsNullOrEmpty(vtacct))
            {
                return new CreateVirtualAccountResponse { Success = false, Message = "無可用的虛擬帳號" };
            }

            // 2. 轉換為板信電文格式 (日期格式需符合銀行規範)
            var deadline = request.Deadline ?? DateTime.Now.AddDays(7);
            var recycleDate = request.RecycleDate ?? DateTime.Now.AddDays(30);

            var bankReq = new RegisterAccountQRCodeRequest
            {
                UUID = Guid.NewGuid().ToString(),
                VTACCT = _config["Bop:Account"] + vtacct,
                CHKREPEAT = request.AllowRepeat ? "Y" : "N",
                CHKSINGLE = "Y",
                SINGLELIMIT = ((int)request.Amount).ToString(),
                CHKTOT = "Y",
                TOTLIMIT = ((int)request.Amount).ToString(),
                QRCODEAMOUNT = ((int)request.Amount).ToString(),
                QRCODENOTE = string.IsNullOrEmpty(request.Note) ? "" : (request.Note.Length > 20 ? request.Note.Substring(0, 20) : request.Note),
                CHKDEADLINE = "Y",
                DEADLINE = deadline.ToString("yyyyMMddHHmm"),
                RECYCLEDATE = recycleDate.ToString("yyyyMMdd")
            };

            // 3. 呼叫板信服務
            var bankResult = await _bopBankService.RegisterAccountQRCodeAsync(bankReq);

            // 4. 判斷回傳結果 (板信成功通常是 "000000" 或 "0000")
            if (bankResult.ReturnCode == "000000" || bankResult.ReturnCode == "0000")
            {
                return new CreateVirtualAccountResponse
                {
                    Success = true,
                    Message = "虛擬帳號建立成功",
                    CustomerId = _config["Bop:CompanyId"],
                    UUID = bankResult.UUID,
                    VirtualAccount = bankResult.VTACCT,
                    QRCode = bankResult.QRCode
                };
            }

            return new CreateVirtualAccountResponse
            {
                Success = false,
                Message = $"銀行端失敗 [{bankResult.ReturnCode}]: {bankResult.ReturnMsg}"
            };
        }
    }
}
