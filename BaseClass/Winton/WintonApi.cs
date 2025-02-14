using KF_WebAPI.BaseClass.AE;
using Newtonsoft.Json.Linq;

namespace KF_WebAPI.BaseClass.Winton
{
    public class ApiResponse<T>
    {
        public string Status { get; set; }
        public string Error { get; set; }
        public T Data { get; set; }
    }

    public class ApiResponse
    {
        public string Status { get; set; }
        public string Error { get; set; }
        public string Data { get; set; }
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
}
