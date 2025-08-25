using System;

namespace KF_WebAPI.BaseClass.AE
{
    public class Performance_req
    {
        /// <summary>
        /// 區
        /// </summary>
        public string? u_bc_title { get; set; }
        /// <summary>
        /// 起訖年月 start
        /// </summary>
        public string? selYear_S { get; set; }
        /// <summary>
        /// 起訖年月 end
        /// </summary>
        public string? selYear_E { get; set; } // 

        /// <summary>
        /// 到職日基準日
        /// </summary>
        public string? start_date { get; set; }
        /// <summary>
        /// 在職狀態
        /// </summary>
        public string? Enable { get; set; }

        /// <summary>
        /// 折數後業績
        /// </summary>
        public string? isACT { get; set; }

        /// <summary>
        /// 排序
        /// </summary>
        public string? OrderBy { get; set; }
    }

    public class Performance_Cens_req
    {
        /// <summary>
        /// 起訖年月 start
        /// </summary>
        public string? Date_S { get; set; }
        /// <summary>
        /// 起訖年月 end
        /// </summary>
        public string? Date_E { get; set; } // 

        /// <summary>
        /// 折數後業績
        /// </summary>
        public string? isACT { get; set; }

        /// <summary>
        /// 業務員編號
        /// </summary>
        //public string? plan_num { get; set; }

        /// <summary>
        /// U_num + yyyy + "-月"
        /// </summary>
        public string? key { get; set; }

    }

    public class Performance_res
    {
        public string? U_num { get; set; }
        public string? Cal_Arrive { get; set; }
        public string? U_BC { get; set; }
        public string? U_name { get; set; }
        public string? U_arrive_date { get; set; }
        public string? U_leave_date { get; set; }
        public string? enable { get; set; }
        public string? U_BC_name { get; set; }
        public string? title { get; set; }
        public string? U_PFT { get; set; }
        public string? plan_num { get; set; }
        public string? yyyy { get; set; }
        public string? Jan { get; set; }
        public string? Feb { get; set; }
        public string? Mar { get; set; }
        public string? Apr { get; set; }
        public string? May { get; set; }
        public string? Jun { get; set; }
        public string? Jul { get; set; }
        public string? Aug { get; set; }
        public string? Sep { get; set; }
        public string? Oct { get; set; }
        public string? Nov { get; set; }
        public string? Dec { get; set; }
        public Int32? Jan_C { get; set; }
        public Int32? Feb_C { get; set; }
        public Int32? Mar_C { get; set; }
        public Int32? Apr_C { get; set; }
        public Int32? May_C { get; set; }
        public Int32? Jun_C { get; set; }
        public Int32? Jul_C { get; set; }
        public Int32? Aug_C { get; set; }
        public Int32? Sep_C { get; set; }
        public Int32? Oct_C { get; set; }
        public Int32? Nov_C { get; set; }
        public Int32? Dec_C { get; set; }
        public string? Totle { get; set; }
        public string? MonAVG { get; set; }
        public string? YearAVG { get; set; }
        public string? cal_yearAvg { get; set; }
    }

    public class Performance_Excel
    {
        public string U_name { get; set; }
        public string U_arrive_date { get; set; }
        public string U_BC_name { get; set; }
        public string title { get; set; }
        public int Jan { get; set; }
        public int Feb { get; set; }
        public int Mar { get; set; }
        public int Apr { get; set; }
        public int May { get; set; }
        public int Jun { get; set; }
        public int Jul { get; set; }
        public int Aug { get; set; }
        public int Sep { get; set; }
        public int Oct { get; set; }
        public int Nov { get; set; }
        public int Dec { get; set; }
        public double Totle { get; set; }
        public int MonAVG { get; set; }
        public int YearAVG { get; set; }
    }

    public class Performance_Cen_res
    {
        /// <summary>
        /// 撥款日
        /// </summary>
        public string? yyyymmdd { get; set; }
        /// <summary>
        /// 業績
        /// </summary>
        public decimal? get_amount { get; set; }
        /// <summary>
        /// 撤件金額
        /// </summary>
        public decimal? Cancel_amount { get; set; }
        /// <summary>
        /// 撤件基準日
        /// </summary>
        public string? CancelDate { get; set; }
        /// <summary>
        /// 撤件人
        /// </summary>
        public string? Cancel_Na { get; set; }
    }
}

