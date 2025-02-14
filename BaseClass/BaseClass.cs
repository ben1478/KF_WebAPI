using System.Collections;

namespace KF_WebAPI.BaseClass
{

    public class BaseClass
    {
        
    }
    /// <summary>
    /// 交易Info
    /// </summary>
    /// <typeparam name="T"></typeparam>
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


    /// <summary>
    /// 基本欄位
    /// </summary>
    public class tbInfo
    {
        public string? add_date { get; set; }
        public string? add_num { get; set; }
        public string? add_ip { get; set; }
        public string? edit_date { get; set; }
        public string? edit_num { get; set; }
        public string? edit_ip{ get; set; }
        public string? del_date { get; set; }
        public string? del_num { get; set; }
        public string? del_ip { get; set; }
        public string? del_tag { get; set; }
        
    }


    public class UserLogin
    {
        public string Username { get; set; }
        public string Password { get; set; }
    }

    public class SpecialClass
    {
        public string BC_Strings { get; set; }
        public string special_check { get; set; }
        public string U_num { get; set; }
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

    public class ExAPILog
    {
        public String API_CODE { get; set; }
        public String API_NAME { get; set; }
        public String API_KEY { get; set; }
        public String ACCESS_TOKEN { get; set; }
        public String PARAM_JSON { get; set; }
        public String RESULT_CODE { get; set; }
        public String RESULT_MSG { get; set; }
        public String? HANDLER { get; set; }
        public String? HANDLE_NOTE { get; set; }
        public String? HANDLE_STATUS { get; set; }
        public DateTime Add_date { get; set; }
        public String Add_User { get; set; }
    }

}
