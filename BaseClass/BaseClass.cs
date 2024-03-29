﻿using System.Collections;

namespace KF_WebAPI.BaseClass
{

    public class BaseClass
    {
        
    }

    public class ResultClass<T>
    {
        /// <summary>
        /// 000:成功;999:失敗
        /// </summary>
        public String ResultCode { get; set; } = "000";
        /// <summary>
        /// 錯誤訊息:ResultCode=999
        /// </summary>
        public String ResultMsg { get; set; } = "";
        public String transactionId { get; set; } = "";

        public T? objResult { get; set; }
    }







    public class APIErrorLog
    {
        /// <summary>
        /// 交易序號
        /// </summary>
        public String TransactionId { get; set; } = "";
        /// <summary>
        /// API名稱
        /// </summary>
     
        public String API_Name { get; set; } = "";
        /// <summary>
        /// 錯誤訊息
        /// </summary>
     
        public String ErrMSG { get; set; } = "";
    }


    public class APILog
    {
        /// <summary>
        /// 交易序號
        /// </summary>
        public String TransactionId { get; set; } = "";
        /// <summary>
        /// 交易序號
        /// </summary>
        public String PreTransactionId { get; set; } = "";
        /// <summary>
        /// 案件編號
        /// </summary>
        public String Form_No { get; set; } = "";
        /// <summary>
        /// API名稱
        /// </summary>
        public String API_Name { get; set; } = "";
        /// <summary>
        /// 傳入參數
        /// </summary>
        public String ParamJSON { get; set; } = "";
        /// <summary>
        /// 回傳參數
        /// </summary>
        public String ResultJSON { get; set; } = "";
        /// <summary>
        /// 呼叫人
        /// </summary>
        public String CallUser { get; set; } = "";
        /// <summary>
        /// 呼叫時間
        /// </summary>
        public String CallTime { get; set; } = "";
        /// <summary>
        /// 傳入參數111
        /// </summary>
        public String StatusCode { get; set; } = "";
    }

}
