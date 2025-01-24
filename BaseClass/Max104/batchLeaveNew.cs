using System.Diagnostics.Eventing.Reader;

namespace KF_WebAPI.BaseClass.Max104
{
    public class batchLeaveNew
    {
        public int? CO_ID { get; set; }
        public string? SESSION_KEY { get; set; }
        public LEAVE_DATA[]? LEAVE_DATA { get; set; }
        public string? WF_NO { get; set; }

        // 簡化構造函數
        public batchLeaveNew() { }
    }



    /// <summary>
    /// 請假
    /// </summary>
    public class LEAVE_DATA
    {
        public string? EMP_ID { get; set; }
        public string? LEAVEITEM_ID { get; set; }
        public string? LEAVE_START { get; set; }
        public string? LEAVE_END { get; set; }
        public string? AGENT_IDS { get; set; }
        public string? REASON { get; set; }
        public LEAVE_DATA() { }
        // 簡化構造函數
        public LEAVE_DATA(string? empId, string? leaveitem_id, string? otStart, string? otEnd, string? agent_ids, string? reason) =>
            (EMP_ID, LEAVEITEM_ID, LEAVE_START, LEAVE_END, AGENT_IDS,  REASON) = (empId, leaveitem_id,otStart, otEnd, agent_ids, reason);


    }

}
