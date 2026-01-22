using System.Text;
using System.Text.Json;

namespace KF_WebAPI.Service
{

    // 前端傳來的請求
    public class ReportRequest
    {
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }

        // 關鍵：決定要做什麼 (例如 "Excel", "Chart", "Insight")
        public string ActionType { get; set; }

        // 使用者額外的指令 (例如 "幫我看哪個人毛利最高")
        public string CustomPrompt { get; set; }
    }

    // 模擬 SQL Server 撈出來的資料結構
    public class SalesData
    {
        public string Date { get; set; }
        public string SalesPerson { get; set; } // 業務員 (Join過員工檔)
        public string Product { get; set; }
        public decimal Amount { get; set; }
        public int Quantity { get; set; }
    }

    public class GeminiService
    {

        private readonly HttpClient _httpClient;
        private readonly string _apiKey = "AIzaSyB9DPKK2zhHKpddNXZjvCokDA-wZbVdas8"; // 建議放在 appsettings.json
        private readonly string _endpoint = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent";

        public GeminiService(HttpClient httpClient)
        {
            _httpClient = httpClient;
        }

        public async Task<string> ProcessDataWithGemini(List<SalesData> data, string actionType, string customPrompt)
        {
            string dataJson = JsonSerializer.Serialize(data);
            string systemInstruction = "";

            // 策略路由：根據功能決定 Prompt
            switch (actionType.ToLower())
            {
                case "excel":
                    // 技巧：要求 AI 給 CSV，因為 C# 轉 CSV 到 Excel 很簡單
                    systemInstruction = @"
                    你是一個資料轉檔助手。請將輸入的 JSON 資料整理成標準的 CSV 格式。
                    規則：
                    1. 包含表頭 (日期, 業務員, 產品, 金額, 數量)。
                    2. 請依據『金額』由大到小排序。
                    3. 直接輸出 CSV 內容，不要有 markdown 標記，不要有其他廢話。";
                    break;

                case "chart":
                    // 技巧：要求回傳 Chart.js 可用的 JSON
                    systemInstruction = @"
                    你是一個前端圖表專家。請分析資料並回傳 JSON 格式供 Chart.js 使用。
                    需求：
                    1. 統計每位業務員的『總業績金額』。
                    2. 回傳格式：{ ""labels"": [""豪哥"", ""小美""], ""datasets"": [{ ""label"": ""業績"", ""data"": [650000, 300000] }] }
                    3. 只要 JSON，不要 markdown。";
                    break;

                default: // Insight / Text
                    systemInstruction = "你是一個資深業務經理。請根據數據提供簡短的業績分析與建議。";
                    break;
            }

            // 組合最終 Prompt
            string fullPrompt = $"{systemInstruction}\n\n使用者額外要求：{customPrompt}\n\n[數據資料開始]\n{dataJson}\n[數據資料結束]";

            return await CallGeminiApi(fullPrompt);
        }

        private async Task<string> CallGeminiApi(string prompt)
        {
            var url = _endpoint + $"?key={_apiKey}";

            var payload = new
            {
                contents = new[] { new { parts = new[] { new { text = prompt } } } }
            };

            var jsonPayload = JsonSerializer.Serialize(payload);
            var content = new StringContent(jsonPayload, Encoding.UTF8, "application/json");

            var response = await _httpClient.PostAsync(url, content);
            response.EnsureSuccessStatusCode();

            var responseString = await response.Content.ReadAsStringAsync();
            using var doc = JsonDocument.Parse(responseString);

            // 擷取 AI 回應文字
            return doc.RootElement.GetProperty("candidates")[0]
                                  .GetProperty("content")
                                  .GetProperty("parts")[0]
                                  .GetProperty("text")
                                  .GetString();
        }
    }

}

