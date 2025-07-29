using System;

namespace KF_WebAPI.BaseClass.AE
{
    public class House_othercase
    {
        public string? case_id { get; set; }
        public string? CaseType { get; set; }
        public string? show_fund_company { get; set; }
        public string? show_project_title { get; set; }
        public string? cs_name { get; set; }
        public string? cs_id { get; set; }
        public string? get_amount { get; set; }
        public string? period { get; set; }
        public string? interest_rate_pass { get; set; }
        public DateTime? get_amount_date { get; set; }
        public string? case_remark { get; set; }
        public int? comm_amt { get; set; }
        public int? act_perf_amt { get; set; }
        public string? plan_num { get; set; }
        
    }

    public class House_othercase_Ins
    {
        public string? case_id { get; set; }
        public string? CaseType { get; set; }
        public string? show_fund_company { get; set; }
        public string? show_project_title { get; set; }
        public string? cs_name { get; set; }
        public string? cs_id { get; set; }
        public string? get_amount { get; set; }
        public string? period { get; set; }
        public string? interest_rate_pass { get; set; }
        public string? get_amount_date { get; set; }
        public string? case_remark { get; set; }
        public int? comm_amt { get; set; }
        public int? act_perf_amt { get; set; }
        public string? plan_num { get; set; }
        public string? add_num { get; set; }
        public string? edit_num { get; set; }
    }

    public class House_othercase_Req
    {
        /// <summary>
        /// 區
        /// </summary>
        public string? BC_code { get; set; }

        /// <summary>
        /// 撥款年月：(YYYY-MM)
        /// </summary>
        
        public string? selYear_S { get; set; }

        /// <summary>
        /// Logged-in user's ID
        /// </summary>
        public string? UserNum { get; set; }

        /// <summary>
        /// Logged-in user's Role ID
        /// </summary>
        public string? UserRole { get; set; }

        /// <summary>
        /// Logged-in user's Branch Code
        /// </summary>
        public string? UserBC { get; set; }
    }

    public class HouseOthercaseQueryResponse
    {
        /// <summary>
        /// The list of case data.
        /// </summary>
        public object Data { get; set; }

        /// <summary>
        /// Indicates if the selected period is financially closed.
        /// </summary>
        public bool IsPeriodClosed { get; set; }
    }

}
