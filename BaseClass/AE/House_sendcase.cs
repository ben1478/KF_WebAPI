
namespace KF_WebAPI.BaseClass.AE
{

    public class House_sendcase_LQuery
    {

        public decimal HS_id { get; set; }
        /// <summary>
        /// 區
        /// </summary>
        public string U_BC_Name { get; set; }

        /// <summary>
        /// 進件日
        /// </summary>
        public string Send_amount_date { get; set; }

        /// <summary>
        /// 申請人
        /// </summary>
        public string CS_name { get; set; }

        /// <summary>
        /// 介紹人
        /// </summary>
        public string CS_introducer { get; set; }

        /// <summary>
        /// 業務
        /// </summary>
        public string plan_name { get; set; }

        /// <summary>
        /// 撥款日
        /// </summary>
        public string get_amount_date { get; set; }

        /// <summary>
        /// 撥款金額(萬)
        /// </summary>
        public string get_amount { get; set; }

        /// <summary>
        /// 承作利率(%)
        /// </summary>
        public string interest_rate_pass { get; set; }


        /// <summary>
        /// 上傳檔案 key
        /// </summary>
        /// 
        public string File_ID { get; set; }
        /// <summary>
        /// 上傳Count
        /// </summary>
        public string upLoad_Count { get; set; }

    }

    /// <summary>
    /// 撥款及費用確認書列表.查詢條件
    /// </summary>
    public class House_sendcase_Req
    {
        /// <summary>
        /// 申請人
        /// </summary>
        public string CS_name { get; set; }

        /// <summary>
        /// 區
        /// </summary>
        public string BC_code { get; set; }

        /// <summary>
        /// 業務
        /// </summary>
        public string plan_name { get; set; }
        

        /// <summary>
        /// 撥款年月
        /// </summary>
        public string selYear_S { get; set; }
        

        /// <summary>
        /// 排序
        /// </summary>
        public string OrderByStr { get; set; }
    }
}
