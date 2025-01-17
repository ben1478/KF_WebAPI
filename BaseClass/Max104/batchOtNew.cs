namespace KF_WebAPI.BaseClass.Max104
{

    public class batchOtNew
    {
        public int? CO_ID { get; set; }
        public string? SESSION_KEY { get; set; }
        public OT_DATA[]? OT_DATA { get; set; }
        public string? WF_NO { get; set; }

        // 簡化構造函數
        public batchOtNew() { }
    }


    public class OT_DATA
    {
        public string? EMP_ID { get; set; }
        public string? OT_START { get; set; }
        public string? OT_END { get; set; }
        /// <summary>
        /// (2:申請加班費 / 3:申請補休 / 7:2+3 )
        /// </summary>
        public string? PAY_TYPE { get; set; }
        /// <summary>
        /// 支領誤餐費(1:是 0:否)
        /// </summary>
        public string? IS_MEAL { get; set; }
        /// <summary>
        /// 是否比對刷卡(1:是 0:否)
        /// </summary>
        public string? IS_CARDMATCH { get; set; }
        public string? REASON { get; set; }
        public OT_DATA() { }
        // 簡化構造函數
        public OT_DATA(string? empId, string? otStart, string? otEnd, string? payType, string? isMeal, string? isCardMatch, string? reason) =>
            (EMP_ID, OT_START, OT_END, PAY_TYPE, IS_MEAL, IS_CARDMATCH, REASON) = (empId, otStart, otEnd, payType, isMeal, isCardMatch, reason);
    }

}
