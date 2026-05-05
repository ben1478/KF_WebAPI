using KF_WebAPI.BaseClass.AE;
using Newtonsoft.Json.Linq;
using System.Collections;

namespace KF_WebAPI.BaseClass.Winton
{
    public class ApietokenResponse
    {
        public string Status { get; set; }
        public string Error { get; set; }
        public string Data { get; set; }
    }

    public class ApiResponse
    {
        public string Status { get; set; }
        public string Error { get; set; }
        public ApiResponseData Data { get; set; }
    }

    public class ApiResponseData
    {
        public string failure { get; set; }
        public int success { get; set; }
        public string total { get; set; }
        public JArray dataset { get; set; }
    }

    public class Summons_req
    {
        public Summons_M_req ADataSetMaster { get; set; }
        public List<Summons_D_req> ADataSetDetail { get; set; }
        public string AToken { get; set; }
        public string AUpdateType { get; set; }
        public string ACreateCB01 { get; set; }

    }

    public class Summons_M_req
    {
        public string MFGL003 { get; set; }
        public string MFGL004 { get; set; }
        public string MFGL005 { get; set; }
        public string MFGL006 { get; set; }
        public string MFGL009 { get; set; }
        public string? MFGL801 { get; set; }
        public string? MFGL802 { get; set; }
        public string? MFGL803 { get; set; }
        public string? MFGL804 { get; set; }
        public string? MFGL805 { get; set; }
        public double? MFGL806 { get; set; }
        public double? MFGL807 { get; set; }
        public double? MFGL808 { get; set; }
        public double? MFGL809 { get; set; }
        public double? MFGL810 { get; set; }
        public DateTime? MFGL811 { get; set; }
        public DateTime? MFGL812 { get; set; }
    }

    public class Summons_D_req
    {
        public string DTGL004 { get; set; }
        public string DTGL005 { get; set; }
        public string DTGL008 { get; set; }
        public string DTGL009 { get; set; }
        public string DTGL011 { get; set; }
        public string DTGL012 { get; set; }
        public double DTGL013 { get; set; }
        public string? DTGL014 { get; set; }
        public string? DTGL015 { get; set; }
        public double DTGL021 { get; set; }
        public string? DTGL027 { get; set; }
        public double DTGL028 { get; set; }
        public string? DTGL801 { get; set; }
        public string? DTGL802 { get; set; }
        public string? DTGL803 { get; set; }
        public string? DTGL804 { get; set; }
        public double? DTGL805 { get; set; }
        public double? DTGL806 { get; set; }
        public double? DTGL807 { get; set; }
        public double? DTGL808 { get; set; }
        public double? DTGL809 { get; set; }
        public double? DTGL810 { get; set; }
        public DateTime? DTGL811 { get; set; }
        public DateTime? DTGL812 { get; set; }
    }

    public class Manufacturer_req
    {
        public string AToken { get; set; }
        public string AUpdateType { get; set; }
        public Manufacturer_M_req ADataSet { get; set; }
    }

    public class Manufacturer_M_req
    {
        public string SU01001 { get; set; }
        public string? SU01002 { get; set; }
        public string SU01003 { get; set; }
        public string SU01004 { get; set; }
        public string? SU01005 { get; set; }
        public string? SU01007 { get; set; }
        public string? SU01008 { get; set; }
        public string? SU01009 { get; set; }
        public string? SU01010 { get; set; }
        public string? SU01011 { get; set; }
        public string? SU01012 { get; set; }
        public string? SU01013 { get; set; }
        public string? SU01014 { get; set; }
        public string? SU01015 { get; set; }
        public string? SU01016 { get; set; }
        public string? SU01017 { get; set; }
        public string? SU01018 { get; set; }
        public string? SU01019 { get; set; }
        public string? SU01020 { get; set; }
        public string? SU01021 { get; set; }
        public string? SU01022 { get; set; }
        public string? SU01023 { get; set; }
        public double? SU01025 { get; set; }
        public string? SU01026 { get; set; }
        public string? SU01027 { get; set; }
        public double? SU01028 { get; set; }
        public string? SU01029 { get; set; }
        public string? SU01030 { get; set; }
        public string? SU01031 { get; set; }
        public string? SU01032 { get; set; }
        public string? SU01033 { get; set; }
        public double? SU01034 { get; set; }
        public double? SU01037 { get; set; }
        public string? SU01038 { get; set; }
        public string? SU01039 { get; set; }
        public string? SU01040 { get; set; }
        public string? SU01041 { get; set; }
        public double? SU01043 { get; set; }
        public double? SU01044 { get; set; }
        public double? SU01045 { get; set; }
        public double? SU01046 { get; set; }
        public double? SU01047 { get; set; }
        public double? SU01048 { get; set; }
        public double? SU01049 { get; set; }
        public double? SU01050 { get; set; }
        public double? SU01051 { get; set; }
        public double? SU01052 { get; set; }
        public double? SU01053 { get; set; }
        public double? SU01054 { get; set; }
        public string? SU01055 { get; set; }
        public string? SU01056 { get; set; }
        public string? SU01057 { get; set; }
        public double? SU01076 { get; set; }
        public double? SU01077 { get; set; }
        public double? SU01078 { get; set; }
        public string? SU01081 { get; set; }
        public string SU01082 { get; set; }
        public string? SU01085 { get; set; }
        public string? SU01087 { get; set; }
        public double? SU01090 { get; set; }
        public string? SU01091 { get; set; }
        public string? SU01092 { get; set; }
        public string? SU01093 { get; set; }
        public string? SU01095 { get; set; }
        public string? SU01096 { get; set; }
        public string? SU01097 { get; set; }
        public string? SU01100 { get; set; }
        public string? SU01101 { get; set; }
        public string? SU01102 { get; set; }
        public string? SU01103 { get; set; }
        public string? SU01105 { get; set; }
        public string? SU01106 { get; set; }
        public string? SU01107 { get; set; }
        public string? SU01108 { get; set; }
        public string? SU01109 { get; set; }
        public string? SU01110 { get; set; }
        public string? SU01112 { get; set; }
        public string? SU01113 { get; set; }
        public string? SU01114 { get; set; }
        public string? SU01115 { get; set; }
        public string? SU01116 { get; set; }
        public string? SU01117 { get; set; }
        public string? SU01119 { get; set; }
        public int? SU01123 { get; set; }
        public string? SU01124 { get; set; }
        public string? SU01125 { get; set; }
        public string? SU01126 { get; set; }
        public string? SU01801 { get; set; }
        public string? SU01802 { get; set; }
        public string? SU01803 { get; set; }
        public string? SU01804 { get; set; }
        public string? SU01805 { get; set; }
        public double? SU01806 { get; set; }
        public double? SU01807 { get; set; }
        public double? SU01808 { get; set; }
        public string? SU01809 { get; set; }
        public string? SU01810 { get; set; }
        public double? SU01811 { get; set; }
        public double? SU01812 { get; set; }
    }

    public class ReceivableForInv_req
    {
        public string AToken { get; set; }
        public string ADocType { get; set; }
        public int AUpdateType { get; set; }
        public string InvoiceGroup { get; set; }
        public List<ReceivableForInv_M_req> ADataSetMaster { get; set; }

        public List<ReceivableForInv_D_req> ADataSetDetail { get; set; }
    }

    public class ReceivableForInv_M_req
    {
        public string MF10003 { get; set; }
        public string MF10004 { get; set; }
        public string MF10008 { get; set; }
        public string MF10010 { get; set; }
        public string MF10011 { get; set; }
        public string MF10012 { get; set; }
        public string MF10018 { get; set; }
        public string MF10022 { get; set; }
        public string MF10059 { get; set; }
        public string MF10066 { get; set; }
        public string MF10091 { get; set; }
        public string MF10093 { get; set; }
        public string MF10094 { get; set; }
    }

    public class ReceivableForInv_D_req
    {
        public string DT10004 { get; set; }
        public string DT10006 { get; set; }
        public int DT10030 { get; set; }
        public decimal DT10040 { get; set; }
        public string DT10021 { get; set; }
    }

    public class Receivable_Win_Inv
    {
        public decimal HS_id { get; set; }
        public decimal RCD_id { get; set; }
        public string CS_name { get; set; }
        public string CS_PID { get; set; }
        public int RC_count { get; set; }
        public string roc_RC_date { get; set; }
        public decimal amount_per_month { get; set; }
        public decimal interest { get; set; }
        public decimal Rmoney { get; set; }
        public int? HFees { get; set; }
        public decimal Ex_RemainingPrincipal { get; set; }
        public decimal amount_total { get; set; }
        public int month_total { get; set; }
        public string? Win_Msg { get; set; }
        public DateTime RecPayDate { get; set; }
        public string? User { get; set; }
        public string? RC_note { get; set; }
    }

    public class PayOff_Win_Inv: Receivable_Win_Inv
    {
        /// <summary>
        /// 違約金/作業費
        /// </summary>
        public decimal Break_AMT { get; set; }
        /// <summary>
        /// A:違約金;B:作業費
        /// </summary>
        public string Break_Type { get; set; }
        /// <summary>
        /// 結清利息
        /// </summary>
        public decimal Interest_AMT { get; set; }
        /// <summary>
        /// 延滯總利息
        /// </summary>
        public decimal Delay_AMT { get; set; }
       
        public int month_total { get;set; }

        public string project_title { get; set; }
    }

    public class PaySelf_Win_Inv: Receivable_Win_Inv
    {
        public string? CP_account_last { get; set; }
        public decimal? CP_Pay_Amt { get; set; }
        public string? CP_bus_remark { get; set; }
        public int FileCount { get; set; }
        public string? str_Pay_Date { get; set; }   
        public string? U_Name { get; set; }
        public string? CaseType { get; set; }
        public string str_add_date { get; set; }
    }

    public class WinInvFileRow
    {
        public string Col1 { get; set; }
        public string Col2 { get; set; }
        public string Col3 { get; set; }
    }

    public class SalesOrder_req
    {
        public string AToken { get; set; }
        public string ADocType { get; set; }
        public string AExpRange { get; set; }
        public string ANoB { get; set; }
        public string ANoE { get; set; }
    }
}
