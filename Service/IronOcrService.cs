
using IronOcr;
using KF_WebAPI.BaseClass.AE;
using System.Text.RegularExpressions;
namespace KF_WebAPI.Service
{
    public class IronOcrService
    {


        public ACHBankInfo ParsePDF_Base64(string file_body_encode)
        {
           
             ACHBankInfo _ACHBankInfo = new ACHBankInfo();
            var Ocr = new IronTesseract();

            // 先把 Base64 轉成 PDF 檔案


            byte[] pdfBytes = Convert.FromBase64String(file_body_encode);
            using (var ms = new MemoryStream(pdfBytes))
            {
                using (var Input = new OcrInput())
                {
                    Ocr.Language = OcrLanguage.ChineseTraditional;
                    Input.DeNoise();
                    Input.TargetDPI = 350;
                    Input.EnhanceResolution();
                    Input.Binarize();
                    Input.LoadPdf(ms); // 用暫存檔載入 PDF

                    var Result = Ocr.Read(Input);
                    string m_Text = Result.Text.Replace(" ", "").Replace("。", "").Replace("|", "").Replace("﹣", "");
                    string m_ErrMsg = "";

                    _ACHBankInfo.IsSuccess = true;

                    // 1. 金融機構代號後面7碼
                    var matchBank = Regex.Match(m_Text, @"金融機構代號(\d{7})");
                    if (matchBank.Success)
                        _ACHBankInfo.BankNo = matchBank.Groups[1].Value;
                    else
                        m_ErrMsg += "解析金融機構代號失敗!";

                    // 2. 委託代繳戶名稱後面的所有數字
                    var matchAccount = Regex.Match(m_Text, @"託代繳戶.*?(\d+)");
                    if (matchAccount.Success)
                        _ACHBankInfo.AccountNo = matchAccount.Groups[1].Value;
                    else
                        m_ErrMsg += "解析銀行帳號失敗!";

                    // 3. 身分證字號 (1字母 + 9數字)
                    var matchId = Regex.Match(m_Text, @"身分證字號.*?([A-Z]\d{9})");
                    if (matchId.Success)
                        _ACHBankInfo.CS_PID = matchId.Groups[1].Value;
                    else
                        m_ErrMsg += "解析身分證字號失敗!";

                    if (m_ErrMsg != "")
                    {
                        _ACHBankInfo.IsSuccess = false;
                        _ACHBankInfo.ResultMsg = m_ErrMsg;
                    }
                }
            }
            

         

            return _ACHBankInfo;
        }



        public ACHBankInfo ParsePDF(string filePath)
        {
            ACHBankInfo _ACHBankInfo = new ACHBankInfo();
            var Ocr = new IronTesseract();
            using (var Input = new OcrInput())
            {
                Ocr.Language = OcrLanguage.ChineseTraditional;
                Input.DeNoise();          // 去除雜訊
                Input.TargetDPI = 350;
                Input.EnhanceResolution(); // 提升解析度
                Input.Binarize();         // 轉成黑白，提高文字對比
                Input.LoadPdf(filePath); // 掃描型 PDF
                var Result = Ocr.Read(Input);
                Console.WriteLine(Result.Text);
                string m_Text = Result.Text.Replace(" ", "").Replace("。", "").Replace("|", "").Replace("﹣", "");
                string m_ErrMsg = "";

                _ACHBankInfo.IsSuccess = true;
                // 1. 金融機構代號後面7碼
                var matchBank = Regex.Match(m_Text, @"金融機構代號(\d{7})");
                if (matchBank.Success)
                {
                    _ACHBankInfo.BankNo = matchBank.Groups[1].Value;
                }
                else
                {
                    m_ErrMsg += "解析金融機構代號失敗!";
                    _ACHBankInfo.ResultMsg = _ACHBankInfo.BankNo;
                }

                // 2. 委託代繳戶名稱後面的所有數字
                var matchAccount = Regex.Match(m_Text, @"託代繳戶.*?(\d+)");
                if (matchAccount.Success)
                {
                    _ACHBankInfo.AccountNo = matchAccount.Groups[1].Value;
                }
                else
                {
                    m_ErrMsg += "解析銀行帳號失敗!";
                }

                // 3. 身分證字號 (1字母 + 9數字)
                var matchId = Regex.Match(m_Text, @"身分證字號.*?([A-Z]\d{9})");
                if (matchId.Success)
                {
                    _ACHBankInfo.CS_PID = matchId.Groups[1].Value;
                }
                else
                {
                    m_ErrMsg += "解析身分證字號失敗!";
                }
                if (m_ErrMsg != "")
                {
                    _ACHBankInfo.IsSuccess = false;
                    _ACHBankInfo.ResultMsg = m_ErrMsg;
                }

            }
            return _ACHBankInfo;
        }
    }
}
