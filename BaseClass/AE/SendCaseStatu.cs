namespace KF_WebAPI.BaseClass.AE
{
    public class SendCaseStatu_req
    {
        public DateTime Send_Date_S { get; set; }
        public DateTime Send_Date_E { get; set;}
        public string Company { get; set; }
        public string? status { get;set; }

    }

    public class SendCaseStatu_res
    {
        public string fund_company { get; set; }
        public string Company_Name { get; set; }
        public int totleCount { get; set; }
        public int ApprCount { get; set; }
        public int unApprCount { get; set; }
        public int PayCount { get; set; }
        public int unPayCount { get; set; }
        public int WPayCount { get; set; }
        public int GUCount { get; set; }
        /// <summary>
        /// 送審中筆數
        /// </summary>
        public int Review_count { get; set; }
        /// <summary>
        /// 婉拒筆數
        /// </summary>
        public int Decline_count { get; set;}
        /// <summary>
        /// 待對保筆數
        /// </summary>
        public int Guarantee { get; set; }
        /// <summary>
        /// 不對保筆數
        /// </summary>
        public int GuaranteeNone { get; set; }
    }

    public class SendCaseStatu_Excel
    {
        public string Company_Name { get; set; }
        public int totleCount { get; set; }
        public int ApprCount { get; set; }
        public int unApprCount { get; set; }
        public int PayCount { get; set; }
        public int unPayCount { get; set; }
        public int WPayCount { get; set; }
        public int GUCount { get; set; }
    }

    public class SendCaseStatu_Det_Excel
    {
        public decimal HS_id { get; set; }
        public string show_fund_company { get; set; }
        public string Send_amount_date { get; set; }
        public string CS_name { get; set;}
        public string CS_ID { get; set;}
        public string CS_MTEL1 { get; set; }
        public string show_appraise_company { get;set; }
        public string show_project_title { get; set; }
        public string ShowType { get; set; }
        public string Send_amount { get; set; }
        public string pass_amount { get; set;}
        public string Loan_rate { get; set;}
        public string addr { get; set;}
    }
}
