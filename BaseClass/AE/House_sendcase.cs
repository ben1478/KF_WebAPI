
namespace KF_WebAPI.BaseClass.AE
{
    /*
    public class House_sendcase
    {
        public decimal HS_id { get; set; }

        public string HS_cknum { get; set; }

        public decimal? HA_id { get; set; }

        public decimal? HP_project_id { get; set; }

        public string fund_company { get; set; }

        public string appraise_company { get; set; }

        public string sendcase_handle_num { get; set; }

        public string sendcase_handle_type { get; set; }

        public DateTime? sendcase_handle_date { get; set; }

        public string Send_amount { get; set; }

        public DateTime? Send_amount_date { get; set; }

        public string Send_result_type { get; set; }

        public DateTime? Send_result_date { get; set; }

        public string pass_amount { get; set; }

        public string check_amount_type { get; set; }

        public string check_amount { get; set; }

        public DateTime? check_amount_date { get; set; }

        public string get_amount_type { get; set; }

        public string get_amount { get; set; }

        public DateTime? get_amount_date { get; set; }

        public string Loan_rate { get; set; }

        public string interest_rate_original { get; set; }

        public string interest_rate_pass { get; set; }

        public string charge_M { get; set; }

        public string charge_flow { get; set; }

        public string charge_agent { get; set; }

        public string charge_check { get; set; }

        public string get_amount_final { get; set; }

        public string introducer { get; set; }

        public string HS_note { get; set; }

        public string exception_type { get; set; }

        public string exception_set_num { get; set; }

        public DateTime? exception_set_date { get; set; }

        public string exception_rate { get; set; }

        public string del_tag { get; set; }

        public DateTime? add_date { get; set; }

        public string add_num { get; set; }

        public string add_ip { get; set; }

        public DateTime? edit_date { get; set; }

        public string edit_num { get; set; }

        public string edit_ip { get; set; }

        public DateTime? del_date { get; set; }

        public string del_num { get; set; }

        public string del_ip { get; set; }

        public decimal? advance_payment_AE { get; set; }

        public decimal? subsidized_interest { get; set; }

        public decimal? act_comm_amt { get; set; }

        public string Introducer_PID { get; set; }

        public string confirm_num { get; set; }

        public DateTime? confirm_date { get; set; }

        public DateTime? misaligned_date { get; set; }

        public decimal? act_perf_amt { get; set; }

        public decimal? act_service_amt { get; set; }

        public string Comm_Remark { get; set; }

        public string CS_introducer { get; set; }

        public string Comparison { get; set; }

        public DateTime? CancelDate { get; set; }

        public string Cancel_num { get; set; }

        public string exc_flag { get; set; }

        public string exc_flag0 { get; set; }

        public string exc_flag1 { get; set; }

        public string exc_flag2 { get; set; }

    }
    */

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

        public string Date_S { get; set; }
        public string Date_E { get; set; }

        /// <summary>
        /// 排序
        /// </summary>
        public string OrderByStr { get; set; }
    }
}
