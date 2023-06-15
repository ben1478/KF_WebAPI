
namespace KF_WebAPI.BaseClass
{
    public class YuRichAPI_Class
    {
        /// <summary>
        /// 1.通路商編號
        /// </summary>
        public String dealerNo { get; set; } = "MM09";

        /// <summary>
        /// 2.進件來源 
        /// </summary>
        public String source { get; set; } = "22";

        /// <summary>
        /// 3.API 電文交易序號 ID 
        /// </summary>
        public String transactionId { get; set; } = "";

        /// <summary>
        /// 4.加密 API 欄位交易資料
        /// </summary>
        public String encryptEnterCase { get; set; } = "";

        /// <summary>
        /// 5.串接規格版本
        /// </summary>
        public String version { get; set; } = "2.0";

    }

    /// <summary>
    /// 案件進件(Receive)
    /// </summary>
    public class Receive
    {
        /// <summary>
        /// 1.申請人中文姓名
        /// </summary>
        public String? customer_name { get; set; }

        /// <summary>
        /// 2.申請人身分證字號
        /// </summary>
        public String? customer_idcard_no { get; set; }

        /// <summary>
        /// 3.身分證初補換
        /// </summary>
        public String? customer_id_number_status { get; set; }

        /// <summary>
        /// 4.身分證初補換日
        /// </summary>
        public String? customer_id_issue_date { get; set; }

        /// <summary>
        ///  5.發證地點
        /// </summary>
        public String? customer_id_number_areacode { get; set; }

        /// <summary>
        /// 6.出生日期
        /// </summary>
        public String? customer_birthday { get; set; }

        /// <summary>
        /// 7.戶籍電話區碼
        /// </summary>
        public String? customer_resident_tel_code { get; set; }
        /// <summary>
        /// 8.戶籍電話
        /// </summary>
        public String? customer_resident_tel_num { get; set; }

        /// <summary>
        /// 9.戶籍電話分機
        /// </summary>
        public String? customer_resident_tel_ext { get; set; }

        /// <summary>
        /// 10.通訊住址電話區碼
        /// </summary>
        public String? customer_mail_tel_code { get; set; }

        /// <summary>
        /// 11.通訊住址電話
        /// </summary>
        public String? customer_mail_tel_num { get; set; }

        /// <summary>
        /// 12.通訊住址電話分機
        /// </summary>
        public String? customer_mail_tel_ext { get; set; }

        /// <summary>
        /// 13.行動電話
        /// </summary>
        public String? customer_mobile_phone { get; set; }

        /// <summary>
        /// 14.教育程度
        /// </summary>
        public String? customer_edutation_status { get; set; }

        /// <summary>
        /// 15.戶籍地郵遞區號
        /// </summary>
        public String? customer_resident_postalcode { get; set; }

        /// <summary>
        /// 16.戶籍地址縣市
        /// </summary>
        public String? customer_resident_addcity { get; set; }

        /// <summary>
        /// 17.戶籍地址鄉鎮
        /// </summary>
        public String? customer_resident_addregion { get; set; }

        /// <summary>
        /// 18.戶籍地址
        /// </summary>
        public String? customer_resident_address { get; set; }

        /// <summary>
        /// 19.同戶籍地址
        /// </summary>
        public String? customer_mail_identical { get; set; }

        /// <summary>
        /// 20.通訊地址郵遞區號
        /// </summary>
        public String? customer_mail_postalcode { get; set; }

        /// <summary>
        /// 21.通訊地址縣市
        /// </summary>
        public String? customer_mail_addcity { get; set; }

        /// <summary>
        /// 22.通訊地址鄉鎮
        /// </summary>
        public String? customer_mail_addregion { get; set; }

        /// <summary>
        /// 23.通訊地址
        /// </summary>
        public String? customer_mail_address { get; set; }

        /// <summary>
        /// 24.居住時間(年)
        /// </summary>
        public String? customer_dwell_year { get; set; }

        /// <summary>
        /// 25.居住時間(月)
        /// </summary>
        public String? customer_dwell_month { get; set; }

        /// <summary>
        /// 27.居住狀況
        /// </summary>
        public String? customer_dwell_status { get; set; }

        /// <summary>
        /// 28.同戶籍或住宅地址
        /// </summary>
        public String? customer_check_identical { get; set; }

        /// <summary>
        /// 29.帳單地址郵遞區號
        /// </summary>
        public String? customer_check_postalcode { get; set; }

        /// <summary>
        /// 30.帳單地址縣市
        /// </summary>
        public String? customer_check_addcity { get; set; }

        /// <summary>
        /// 31.帳單地址鄉鎮
        /// </summary>
        public String? customer_check_addregion { get; set; }

        /// <summary>
        /// 32.帳單地址
        /// </summary>
        public String? customer_check_address { get; set; }

        /// <summary>
        /// 33.E-MAIL
        /// </summary>
        public String? customer_email { get; set; }

        /// <summary>
        /// 34.職業狀態
        /// </summary>
        public String? customer_profession_status { get; set; }

        /// <summary>
        /// 35.公司名稱
        /// </summary>
        public String? customer_company_name { get; set; }

        /// <summary>
        /// 36.公司電話區碼
        /// </summary>
        public String? customer_company_tel_code { get; set; }

        /// <summary>
        /// 37.公司電話號碼
        /// </summary>
        public String? customer_company_tel_num { get; set; }

        /// <summary>
        /// 38.公司電話分機
        /// </summary>
        public String? customer_company_tel_ext { get; set; }

        /// <summary>
        /// 39.公司地址郵遞區號
        /// </summary>
        public String? customer_company_postalcode { get; set; }

        /// <summary>
        /// 40.公司地址縣市
        /// </summary>
        public String? customer_company_addcity { get; set; }

        /// <summary>
        /// 41.公司地址鄉鎮
        /// </summary>
        public String? customer_company_addregion { get; set; }

        /// <summary>
        /// 42.公司地址
        /// </summary>
        public String? customer_company_address { get; set; }

        /// <summary>
        /// 43.職稱
        /// </summary>
        public String? customer_job_type { get; set; }

        /// <summary>
        /// 44.年資(年)
        /// </summary>
        public String? customer_work_year { get; set; }

        /// <summary>
        /// 45.年資(月)
        /// </summary>
        public String? customer_work_month { get; set; }

        /// <summary>
        /// 46.月薪
        /// </summary>
        public String? customer_month_salary { get; set; }

        /// <summary>
        /// 47.銀行資料狀態
        /// </summary>
        public String? customer_creditcard_status { get; set; }

        /// <summary>
        /// 48.銀行資料狀態說明
        /// </summary>
        public String? customer_creditcard_status_remark { get; set; }

        /// <summary>
        /// 49.發卡銀行
        /// </summary>
        public String? customer_creditcard_bank { get; set; }

        /// <summary>
        /// 50.有效日期(年)
        /// </summary>
        public String? customer_creditcard_validdate_year { get; set; }

        /// <summary>
        /// 51.有效日期(月)
        /// </summary>
        public String? customer_creditcard_validdate_month { get; set; }

        /// <summary>
        /// 52.撥款對象
        /// </summary>
        public String? payee_type { get; set; }

        /// <summary>
        /// 53.撥款人帳戶名稱
        /// </summary>
        public String? payee_account_name { get; set; }

        /// <summary>
        /// 54.撥款人身分證字號
        /// </summary>
        public String? payee_account_idno { get; set; }

        /// <summary>
        /// 55.撥款銀行代碼
        /// </summary>
        public String? payee_bank_code { get; set; }

        /// <summary>
        /// 56.撥款銀行分行代碼
        /// </summary>
        public String? payee_bank_detail_code { get; set; }

        /// <summary>
        /// 57.撥款帳號
        /// </summary>
        public String? payee_account_num { get; set; }

        /// <summary>
        /// 58.繳款方式
        /// </summary>
        public String? payment_mode { get; set; }

        /// <summary>
        /// 59.(連保/配偶)連保選項
        /// </summary>
        public String? guarantor_option { get; set; }

        /// <summary>
        /// 60.(連保/配偶)姓名
        /// </summary>
        public String? guarantor_name { get; set; }

        /// <summary>
        /// 61.(連保/配偶)關係
        /// </summary>
        public String? guarantor_relation { get; set; }

        /// <summary>
        /// 62.(連保/配偶)身分證字號
        /// </summary>
        public String? guarantor_idcard_no { get; set; }

        /// <summary>
        /// 63.(連保/配偶)出生日期
        /// </summary>
        public String? guarantor_birthday { get; set; }

        /// <summary>
        /// 64.(連保/配偶)住家電話區碼
        /// </summary>
        public String? guarantor_resident_tel_code { get; set; }

        /// <summary>
        /// 65.(連保/配偶)住家電話號碼
        /// </summary>
        public String? guarantor_resident_tel_num { get; set; }

        /// <summary>
        /// 66.(連保/配偶)住家電話號碼-分機
        /// </summary>
        public String? guarantor_resident_tel_ext { get; set; }

        /// <summary>
        /// 67.(連保/配偶)行動電話
        /// </summary>
        public String? guarantor_mobile_phone { get; set; }

        /// <summary>
        /// 68.(連保/配偶)公司名稱
        /// </summary>
        public String? guarantor_company_name { get; set; }

        /// <summary>
        /// 69.(連保/配偶)職稱
        /// </summary>
        public String? guarantor_job_type { get; set; }

        /// <summary>
        /// 70.(連保/配偶)公司電話-區碼
        /// </summary>
        public String? guarantor_company_tel_code { get; set; }

        /// <summary>
        /// 71.(連保/配偶)公司電話-號碼
        /// </summary>
        public String? guarantor_company_tel_num { get; set; }

        /// <summary>
        /// 72.(連保/配偶)公司電話-分機
        /// </summary>
        public String? guarantor_company_tel_ext { get; set; }

        /// <summary>
        /// 73.(連保/配偶)地址區碼
        /// </summary>
        public String? guarantor_postalcode { get; set; }

        /// <summary>
        /// 74.(連保/配偶)地址縣市
        /// </summary>
        public String? guarantor_addcity { get; set; }

        /// <summary>
        /// 75.(連保/配偶)地址鄉鎮
        /// </summary>
        public String? guarantor_addregion { get; set; }

        /// <summary>
        /// 76.(連保/配偶)地址
        /// </summary>
        public String? guarantor_address { get; set; }

        /// <summary>
        /// 77.聯絡人中文姓名
        /// </summary>
        public String? contact_person_name_i { get; set; }

        /// <summary>
        /// 78.聯絡人關係
        /// </summary>
        public String? contact_person_relation_i { get; set; }

        /// <summary>
        /// 79.聯絡人行動電話
        /// </summary>
        public String? contact_person_mobile_phone_i { get; set; }

        /// <summary>
        /// 80.聯絡人住宅電話-區碼
        /// </summary>
        public String? contact_person_areacode_i { get; set; }

        /// <summary>
        /// 81.聯絡人住宅電話-號碼
        /// </summary>
        public String? contact_person_tel_i { get; set; }

        /// <summary>
        /// 82.聯絡人住宅電話-分機
        /// </summary>
        public String? contact_person_tel_ext_i { get; set; }

        /// <summary>
        /// 83.聯絡人住宅電話-區碼
        /// </summary>
        public String? contact_person_company_areacode_i { get; set; }

        /// <summary>
        /// 84.聯絡人住宅電話-號碼
        /// </summary>
        public String? contact_person_company_tel_i { get; set; }

        /// <summary>
        /// 85.聯絡人住宅電話-分機
        /// </summary>
        public String? contact_person_company_tel_ext_i { get; set; }

        /// <summary>
        /// 86.聯絡人中文姓名
        /// </summary>
        public String? contact_person_name_ii { get; set; }

        /// <summary>
        /// 87.聯絡人關係
        /// </summary>
        public String? contact_person_relation_ii { get; set; }

        /// <summary>
        /// 88.聯絡人行動電話
        /// </summary>
        public String? contact_person_mobile_phone_ii { get; set; }

        /// <summary>
        /// 89.聯絡人住宅電話-區碼
        /// </summary>
        public String? contact_person_areacode_ii { get; set; }

        /// <summary>
        /// 90.聯絡人住宅電話-號碼
        /// </summary>
        public String? contact_person_tel_ii { get; set; }

        /// <summary>
        /// 91.聯絡人住宅電話-分機
        /// </summary>
        public String? contact_person_tel_ext_ii { get; set; }

        /// <summary>
        /// 92.聯絡人住宅電話-區碼
        /// </summary>
        public String? contact_person_company_areacode_ii { get; set; }

        /// <summary>
        /// 93.聯絡人住宅電話-號碼
        /// </summary>
        public String? contact_person_company_tel_ii { get; set; }

        /// <summary>
        /// 94.聯絡人住宅電話-分機
        /// </summary>
        public String? contact_person_company_tel_ext_ii { get; set; }

        /// <summary>
        /// 95.商品利率選項
        /// </summary>
        public String? product_rate_option { get; set; }

        /// <summary>
        /// 96.商品案件選項
        /// </summary>
        public String? product_case_option { get; set; }

        /// <summary>
        /// 97.商品類別
        /// </summary>
        public String? product_category_id { get; set; }

        /// <summary>
        /// 98.商品代碼
        /// </summary>
        public String? product_id { get; set; }

        /// <summary>
        /// 99.頭款(訂金)
        /// </summary>
        public String? deposit { get; set; }

        /// <summary>
        /// 100.辦理分期金額
        /// </summary>
        public String? staging_amount { get; set; }

        /// <summary>
        /// 101.促銷專案
        /// </summary>
        public String? promotion_no { get; set; }

        /// <summary>
        /// 102.期數
        /// </summary>
        public String? periods_num { get; set; }

        /// <summary>
        /// 103.每期應繳金額
        /// </summary>
        public String? payment { get; set; }

        /// <summary>
        /// 104.分期總額
        /// </summary>
        public String? staging_total_price { get; set; }

        /// <summary>
        /// 105.通路商代號
        /// </summary>
        public String? dealer_no { get; set; }

        /// <summary>
        /// 106.通路商統編
        /// </summary>
        public String? dealer_id_no { get; set; }

        /// <summary>
        /// 107.經銷商或債權讓與人名稱
        /// </summary>
        public String? dealer_name { get; set; }

        /// <summary>
        /// 108.經銷商或債權讓與人電話
        /// </summary>
        public String? dealer_tel { get; set; }

        /// <summary>
        /// 109.經銷商或債權讓與人傳真
        /// </summary>
        public String? dealer_fax { get; set; }

        /// <summary>
        /// 110.據點/人員 ID
        /// </summary>
        public String? contact_id_no { get; set; }

        /// <summary>
        /// 111.據點/人員
        /// </summary>
        public String? contact_name { get; set; }

        /// <summary>
        /// 112.據點代號
        /// </summary>
        public String? dealer_branch_no { get; set; }

        /// <summary>
        /// 113.同經銷商或債權讓與人
        /// </summary>
        public String? dealer_branch_name_identical { get; set; }

        /// <summary>
        /// 114.經辦店名稱
        /// </summary>
        public String? dealer_branch_name { get; set; }

        /// <summary>
        /// 115.經辦店電話
        /// </summary>
        public String? dealer_branch_tel { get; set; }

        /// <summary>
        /// 116.經辦人手機
        /// </summary>
        public String? contact_phone { get; set; }

        /// <summary>
        /// 117.經銷商備註選項
        /// </summary>
        public String? dealer_note_code { get; set; }

        /// <summary>
        /// 118.經銷商備註選項_指定照會時間
        /// </summary>
        public String? dealer_note_date { get; set; }

        /// <summary>
        ///119. 經銷商備註欄
        /// </summary>
        public String? dealer_note { get; set; }

        /// <summary>
        /// 120.業務別
        /// </summary>
        public String? bus_type { get; set; }

        /// <summary>
        /// 121.業務別名稱
        /// </summary>
        public String? bus_type_name { get; set; }

        /// <summary>
        /// 122.負責人ID
        /// </summary>
        public String? company_principal_id { get; set; }

        /// <summary>
        /// 123.負責人姓名
        /// </summary>
        public String? company_principal_name { get; set; }

        /// <summary>
        /// 124.撥佣對象
        /// </summary>
        public String? commission_target { get; set; }

        /// <summary>
        /// 撥佣對象
        /// </summary>
        public attachmentFile[]? attachmentFile { get; set; } =  { };


    }

    /// <summary>
    /// 案件進件Response
    /// </summary>
    public class Receive_Result
    {
        /// <summary>
        /// 1.審件編號
        /// </summary>
        public String? examineNo { get; set; }
        /// <summary>
        /// 1.審件狀態
        /// </summary>
        public String? examStatus { get; set; }
        /// <summary>
        /// 1.檔案補件路徑
        /// </summary>
        public String? fileUrl { get; set; }
    }


    /// <summary>
    /// 請款請求(RequestPayment)
    /// </summary>
    public class RequestPayment
    {
        /// <summary>
        /// 1.通路商編號
        /// </summary>
        public String? dealerNo { get; set; }

        /// <summary>
        /// 2.據點編號 
        /// </summary>
        public String? branchNo { get; set; }

        /// <summary>
        /// 3.業務人員 ID 
        /// </summary>
        public String? salesNo { get; set; }

        /// <summary>
        /// 4.審件編號 
        /// </summary>
        public String? examineNo { get; set; }

        /// <summary>
        /// 5.URICH 案件編號
        /// </summary>
        public String? caseNo { get; set; }

        /// <summary>
        /// 6.訂單編號
        /// </summary>
        public String? dealerCaseNo { get; set; }

        /// <summary>
        /// 7.進件來源 
        /// </summary>
        public String? source { get; set; }

        /// <summary>
        /// 8.附件資料陣列 
        /// </summary>
        public attachmentFile[]? attachmentFile { get; set; }
    }


    public class YuRichBase
    {
        /// <summary>
        /// 1.通路商編號
        /// </summary>
        public String? dealerNo { get; set; } 

        /// <summary>
        /// 2.據點編號
        /// </summary>
        public String? branchNo { get; set; } 

        /// <summary>
        /// 3.業務人員 ID
        /// </summary>
        public String? salesNo { get; set; }

        /// <summary>
        /// 4.審件編號
        /// </summary>
        public String? examineNo { get; set; }
    }


    public class NotifyCaseStatus : YuRichBase
    {
        public String? examStatusExplain { get; set; }
        public String? ModifyTime { get; set; }
    }

    public class NotifyAppropriation : YuRichBase
    {
        /// <summary>
        /// 5.撥款時間 
        /// </summary>
        public String? appropriationDate { get; set; }
        /// <summary>
        /// 6.撥款金額
        /// </summary>
        public String? appropriationAmt { get; set; }
        /// <summary>
        /// 7.繳款方式
        /// </summary>
        public String? repayKindName { get; set; }
        /// <summary>
        /// 8.撥款狀態
        /// </summary>
        public String? status { get; set; }
    }


    /// <summary>
    /// 撥款狀態查詢(QueryAppropriation)
    /// </summary>
    public class QueryAppropriation
    {
        /// <summary>
        /// 1.通路商編號
        /// </summary>
        public String? dealerNo { get; set; } 

        /// <summary>
        /// 2.據點編號
        /// </summary>
        public String? branchNo { get; set; } 

        /// <summary>
        /// 3.業務人員 ID
        /// </summary>
        public String? salesNo { get; set; }

        /// <summary>
        /// 4.來源
        /// </summary>
        public String? source { get; set; }

        /// <summary>
        /// 5.審件編號 
        /// </summary>
        public String? examineNo { get; set; }

    }

    /// <summary>
    /// 案件申覆爭取(RequestforExam)
    /// </summary>
    public class RequestforExam
    {
        /// <summary>
        /// 1.通路商編號
        /// </summary>
        public String? dealerNo { get; set; }

        /// <summary>
        /// 2.據點編號
        /// </summary>
        public String? branchNo { get; set; }

        /// <summary>
        /// 3.業務人員 ID
        /// </summary>
        public String? salesNo { get; set; }

        /// <summary>
        /// 4.審件編號 
        /// </summary>
        public String? examineNo { get; set; }

        /// <summary>
        /// 5.申覆內容 
        /// </summary>
        public String? comment { get; set; }

        /// <summary>
        /// 6.來源 
        /// </summary>
        public String? source { get; set; }

        /// <summary>
        /// 7.強制爭取 
        /// </summary>
        public String? forceTryForExam { get; set; }

        /// <summary>
        /// 8.附件資料陣列 
        /// </summary>
        public attachmentFile[]? attachmentFile { get; set; }
    }

    /// <summary>
    /// 審件補件(RequestSupplement)
    /// </summary>
    public class RequestSupplement
    {
        /// <summary>
        /// 1.通路商編號
        /// </summary>
        public String? dealerNo { get; set; }

        /// <summary>
        /// 2.據點編號
        /// </summary>
        public String? branchNo { get; set; }

        /// <summary>
        /// 3.業務人員 ID
        /// </summary>
        public String? salesNo { get; set; }

        /// <summary>
        /// 4.審件編號 
        /// </summary>
        public String? examineNo { get; set; }

        /// <summary>
        /// 5.來源 
        /// </summary>
        public String? source { get; set; }

        /// <summary>
        /// 6.補件說明陣列
        /// </summary>
        public supplement[]? supplement { get; set; }

        /// <summary>
        /// 7.補件檔案陣列 
        /// </summary>
        public attachmentFile[]? attachmentFile { get; set; }

    }

    /// <summary>
    /// 案件狀態查詢(QueryCaseStatus)
    /// </summary>
    public class QueryCaseStatus
    {
        /// <summary>
        /// 1.通路商編號
        /// </summary>
        public String? dealerNo { get; set; }

        /// <summary>
        /// 2.據點編號
        /// </summary>
        public String? branchNo { get; set; }

        /// <summary>
        /// 3.業務人員 ID
        /// </summary>
        public String? salesNo { get; set; }

        /// <summary>
        /// 4.審件編號 
        /// </summary>
        public String? examineNo { get; set; }

        /// <summary>
        /// 5.來源 
        /// </summary>
        public String? source { get; set; }

    }

    /// <summary>
    /// 重新照會通知(reCallout)
    /// </summary>
    public class reCallout
    {
        /// <summary>
        /// 1.通路商編號
        /// </summary>
        public String? dealerNo { get; set; }

        /// <summary>
        /// 2.據點編號
        /// </summary>
        public String? branchNo { get; set; }

        /// <summary>
        /// 3.業務人員 ID
        /// </summary>
        public String? salesNo { get; set; }

        /// <summary>
        /// 4.來源 
        /// </summary>
        public String? source { get; set; }

        /// <summary>
        /// 5.審件編號 
        /// </summary>
        public String? examineNo { get; set; }

        /// <summary>
        /// 6.重照時間 
        /// </summary>
        public String? calloutDate { get; set; }

        /// <summary>
        /// 7.連絡電話  
        /// </summary>
        public String? tel { get; set; }

        /// <summary>
        /// 8.備註描述  
        /// </summary>
        public String? descript { get; set; }

        /// <summary>
        /// 9.現可照  
        /// </summary>
        public String? nowCallout { get; set; }


    }


    public class PutFileToExamiePath
    {
        /// <summary>
        /// 1.通路商編號
        /// </summary>
        public String? dealerNo { get; set; }

        /// <summary>
        /// 2.據點編號
        /// </summary>
        public String? branchNo { get; set; }

        /// <summary>
        /// 3.業務人員 ID
        /// </summary>
        public String? salesNo { get; set; }

        /// <summary>
        /// 4.審件編號 
        /// </summary>
        public String? examineNo { get; set; }

        // <summary>
        /// 5.URICH 案件編號
        /// </summary>
        public String? caseNo { get; set; }

        /// <summary>
        /// 6.訂單編號
        /// </summary>
        public String? dealerCaseNo { get; set; }

        /// <summary>
        /// 7.來源 
        /// </summary>
        public String? source { get; set; }

        /// <summary>
        /// 8.附件資料陣列 
        /// </summary>
        public attachmentFile[]? attachmentFile { get; set; }
    }

    /// <summary>
    /// 補件說明陣列
    /// </summary>
    public class supplement
    {
        /// <summary>
        /// 補件項目種類 固定帶 00
        /// </summary>
        public String? item { get; set; }

        /// <summary>
        /// 補件說明 
        /// </summary>
        public String? comment { get; set; }
    }


    /// <summary>
    /// 申貸檢附資料
    /// </summary>
    public class attachmentFile
    {
        /// <summary>
        /// 檔案編碼索引
        /// </summary>
        public String? file_index { get; set; }

        /// <summary>
        /// 檔案主體
        /// </summary>
        public String? file_body_encode { get; set; }

        /// <summary>
        /// 檔案大小
        /// </summary>
        public String? file_size { get; set; }

        /// <summary>
        /// 檔案格式
        /// </summary>
        public String? content_type { get; set; }
        /// <summary>
        /// 檔案格式
        /// </summary>
        public String? file_name { get; set; }

        public string? TransactionId { get; set; } = "";
    }


    /// <summary>
    /// 回傳狀態
    /// </summary>
    public class BaseResult
    {
        /// <summary>
        /// 回覆代碼
        /// </summary>
        public string code { get; set; } 

        /// <summary>
        /// 回覆訊息 
        /// </summary>
        public string msg { get; set; } = "";
        /// <summary>
        /// 回傳TransactionId
        /// </summary>
        public string? TransactionId { get; set; } = "";
    }



    /// <summary>
    /// 3.2 案件進件(Receive)回傳參數
    /// </summary>
    public class Result_R : BaseResult
    {

        /// <summary>
        /// 審件編號
        /// </summary>
        public string? examineNo { get; set; } = "";
        /// <summary>
        /// 審件狀態 擴充預留欄位，目前不會回傳
        /// </summary>
        public string? examStatus { get; set; } = "";
        /// <summary>
        /// 檔案補件路徑 擴充預留欄位，目前不會回傳
        /// </summary>
        public string? fileUrl { get; set; } = "";
       
    }


    /// <summary>
    /// 3.4 撥款狀態查詢回傳參數
    /// </summary>
    public class Result_QA : BaseResult
    {

        /// <summary>
        /// 撥款資料
        /// </summary>
        public Appropriations[]? appropriations { get; set; } = null;

    }


    /// <summary>
    /// 3.6 案件申覆爭取(RequestforExam) 
    /// </summary>
    public class Result_RE : BaseResult
    {

        /// <summary>
        /// 申覆次數
        /// </summary>
        public string? negotiateTimes { get; set; } = "";
    }

    /// <summary>
    /// 3.9 案件狀態查詢(QueryCaseStatus) 
    /// </summary>
    public class Result_QCS : BaseResult
    {

        /// <summary>
        /// 案件狀態 (收件,核准,婉拒,附條件,待補,補件,申覆,自退)
        /// </summary>
        public string? examStatusExplain { get; set; } = "";
        /// <summary>
        /// 專案代碼 
        /// </summary>
        public string? promoNo { get; set; } = "";
        /// <summary>
        /// 專案名稱 
        /// </summary>
        public string? promoName { get; set; } = "";
        /// <summary>
        /// 客戶保人資料陣列(會有多筆，客戶或是保人)
        /// </summary>
        public Customer[] Customer { get; set; } = { };

        /// <summary>
        /// 申貸本金
        /// </summary>
        public string instCap { get; set; } = "";

        /// <summary>
        /// 期付資訊陣列
        /// </summary>
        public payment[] payment { get; set; } = { };

        /// <summary>
        /// 備註說明 
        /// </summary>
        public string examineComment { get; set; } = "";
  
        /// <summary>
        /// 審件原因陣列
        /// </summary>
        public reasonSuggestionDetail[] reasonSuggestionDetail { get; set; } = { };

        /// <summary>
        /// 核准日期 
        /// </summary>
        public string approveDate { get; set; } = "";

        /// <summary>
        /// 核准日期 
        /// </summary>
        public Payee? Payee { get; set; } = null;

        /// <summary>
        /// 撥款資訊陣列/// </summary>
        public  capitalApply[] capitalApply { get; set; } = {};

        /// <summary>
        /// 舊約編號 
        /// </summary>
        public string oldLoanNo { get; set; } = "";
        /// <summary>
        /// 佣金資訊 
        /// </summary>
        public string brokeragePersonal { get; set; } = "";
        /// <summary>
        /// 預計撥付金額
        /// </summary>
        public string netBorrowToPayBackAmt { get; set; } = "";

        /// <summary>
        /// 預計結清金額
        /// </summary>
        public string borrowToPayBackAmt { get; set; } = "";

        /// <summary>
        /// 產品(車型)品牌
        /// </summary>
        public string carBrand { get; set; } = "";

        /// <summary>
        /// 產品(車型)名稱
        /// </summary>
        public string carName { get; set; } = "";

        /// <summary>
        /// 牌照號碼
        /// </summary>
        public string carNo { get; set; } = "";
    }

    /// <summary>
    /// 撥款資訊
    /// </summary>
    public class capitalApply
    {
        /// <summary>
        /// 撥款日期  格式 yyyyMMdd
        /// </summary>
        public string appropriateDate { get; set; } = "";
        /// <summary>
        /// 撥款金額 
        /// </summary>
        public string remitAmount { get; set; } = "";
        /// <summary>
        /// 撥款對象說明;欄位為文字型態，可使
        /// 用此欄位，判斷借新還
        ///  舊案件撥款金額( remitAmount)種類 ，邏輯如下
        ///  1. 代碼為「裕富」時 remitAmount 內金額為【前貸金額】。 
        /// 2.代碼為「客人」時，remitAmount 內金額為【撥款新案的剩餘金額】。 
        /// </summary>
        public string payeeTypeName { get; set; } = "";

    }




    /// <summary>
    /// 撥款資訊
    /// </summary>
    public class Payee
    {
        /// <summary>
        /// 客戶帳戶 1 ;據點帳戶 3 ;通路商帳戶 10 ;業務員帳戶 13 
        /// </summary>
        public string payeeType { get; set; } = "";

        /// <summary>
        /// 撥款銀行總行代碼
        /// </summary>
        public string bankCode { get; set; } = "";

        /// <summary>
        /// 撥款銀行分行代碼
        /// </summary>
        public string bankDetailCode { get; set; } = "";

        /// <summary>
        /// 撥款帳戶帳號
        /// </summary>
        public string accountNo { get; set; } = "";
    }
    /// <summary>
    /// 審件原因陣列
    /// </summary>
    public class reasonSuggestionDetail
    {
        /// <summary>
        /// 原因/建議  
        /// </summary>
        public string kind { get; set; } = "";
        /// <summary>
        /// 審件狀態  
        /// </summary>
        public string explain { get; set; } = "";
        /// <summary>
        /// 原因/建議-備註說明
        /// </summary>
        public string comment { get; set; } = "";

    }

    /// <summary>
    /// 期付資訊
    /// </summary>
    public class payment
    {
        /// <summary>
        /// 序號
        /// </summary>
        public string seqNo { get; set; } = "";
        /// <summary>
        /// 期數
        /// </summary>
        public string instNo { get; set; } = "";
        /// <summary>
        /// 期付金額
        /// </summary>
        public string instAmt { get; set; } = "";
    }

    public class Customer
    {
        /// <summary>
        /// 客戶 ID 
        /// </summary>
        public string idno { get; set; } = "";

        /// <summary>
        /// 客戶名稱  
        /// </summary>
        public string name { get; set; } = "";

        /// <summary>
        /// 客戶生日  (yyyMMdd)
        /// </summary>
        public string birthday { get; set; } = "";

        /// <summary>
        /// 0 客戶本人 ;1 保人(㇐);2 保人(二);3 保人(三)
        /// </summary>
        public string index { get; set; } = "";

        /// <summary>
        /// 客戶行動電話陣列
        /// </summary>
        public mobilePhone[] mobilePhone { get; set; } = { };
        /// <summary>
        /// 照會結果
        /// </summary>
        public calloutResult calloutResult { get; set; } = null;
    }


    /// <summary>
    /// 客戶行動電話物件
    /// </summary>
    public class mobilePhone
    {
        /// <summary>
        /// 0 客戶行動電話
        /// </summary>
        public string number { get; set; } = "";
    }

    /// <summary>
    /// 照會結果
    /// </summary>
    public class calloutResult
    {
        /// <summary>
        /// 照會備註 
        /// </summary>
        public string comment { get; set; } = "";
    }

    /// <summary>
    /// 撥款資料
    /// </summary>
    public class Appropriations
    {
        /// <summary>
        /// 審件編號
        /// </summary>
        public String? examineNo { get; set; }

        /// <summary>
        /// 撥款時間(yyyyMMddhhmm)
        /// </summary>
        public String? appropriationDate { get; set; }

        /// <summary>
        /// 撥款金額 Ex:100000
        /// </summary>
        public String? appropriationAmt { get; set; }

        /// <summary>
        /// 繳款方式 1.支票;2.繳款單;3.ACH 4.電子繳款單;5.信用卡;6.簡訊繳款
        /// 文字格式，經銷商必須於撥款後提醒客戶第㇐期繳款時間及繳款方式
        /// </summary>
        public String? repayKindName { get; set; }

        /// <summary>
        /// 撥款狀態 A001:未申請;A002:申請中;A003:撥款中;A004:已撥款   
        /// </summary>
        public String? status { get; set; }

    }


    public class objInsertReceive
    {
        public TbReceive? _Receive1 { get; set; }
       
    }




    public class tbQCS
    {
        public string? form_no { get; set; }
        public string? qcs_idx { get; set; }
        public string? qcs_time { get; set; }
        public string? explain { get; set; }
        public string? comment { get; set; }
        public string? transactionId { get; set; }
        public string? resulType { get; set; }


    }
    public class tbRE
    {
        public string? form_no { get; set; }
        public string? re_idx { get; set; }
        public string? re_time { get; set; }
        public string? comment { get; set; }
        public string? transactionId { get; set; }
        public string? transactionId_qcs { get; set; }
        public string? explain { get; set; }
        public string? qcs_time { get; set; }
        public string? qcs_comment { get; set; }
    }
    public class tbRS
    {
        public string? form_no { get; set; }
        public string? rs_idx { get; set; }
        public string? rs_time { get; set; }
        public string? comment1 { get; set; }
        public string? comment2 { get; set; }
        public string? comment3 { get; set; }
        public string? comment4 { get; set; }
        public string? comment5 { get; set; }
        public string? transactionId { get; set; }
        public string? transactionId_qcs { get; set; }
        public string? explain { get; set; }
        public string? qcs_time { get; set; }
        public string? qcs_comment { get; set; }
    }

    public class tbRP
    {
        public string? form_no { get; set; }
        public string? rp_idx { get; set; }
        public string? rp_time { get; set; }
        public string? transactionId { get; set; }
        public string? transactionId_qcs { get; set; }
        public string? resq_time { get; set; }

        public string? statusDesc { get; set; }
        public capitalApply[]? capitalApply { get; set; }



    }

    public class tbQA
    {
        public string? form_no { get; set; }
        public string? qa_idx { get; set; }
        public string? qa_time { get; set; }
        public string? transactionId { get; set; }
        public string? transactionId_qa { get; set; }
        public string? transactionId_qcs { get; set; }
        public string? qcs_time { get; set; }
        public string? qcs_comment { get; set; }
    }


    public class TbReceive : Receive
    {
        public string? Action { get; set; } = "";
        public string? Case_Company { get; set; }
        public string? child { get; set; }
        public string? childcount { get; set; }
        public string? marital { get; set; }
        public string? contact_person_company_name_i { get; set; }
        public string? contact_person_company_name_ii { get; set; }
        public string? guarantor_profession_status { get; set; }
        public string? add_user { get; set; }
        public string? add_date { get; set; }
        public string? upd_user { get; set; }
        public string? upd_date { get; set; }

        
        public string? status { get; set; }
        public string? casestatus { get; set; }
        public string? transactionId { get; set; }
        public string? form_no { get; set; }
    }



    public class objReceive : TbReceive
    {
        public attachmentFile[]? rPattachmentFile { get; set; }
        public attachmentFile[]? rEattachmentFile { get; set; }
        public string? re_comment { get; set; }
        public string? forcetryforexam { get; set; }
        public string? re_count { get; set; }
        public string? re_User { get; set; }
        public attachmentFile[]? rSattachmentFile { get; set; }
        public string? casestatusdesc { get; set; }
        public string? rs_count { get; set; }
        public string? rs_user { get; set; }
        public string? rp_count { get; set; }
        public string? rp_user { get; set; }
        public string? rp_date { get; set; }
        public string? rp_amt { get; set; }
        public string? rp_kindname { get; set; }
        public string? rp_status { get; set; }
        public string? rp_statusdesc { get; set; }
        public List<tbQCS> lisQCS { get; set; }
        public List<tbRE> lisRE { get; set; }
        public List<tbRS> lisRS { get; set; }
        public List<tbRP> lisRP { get; set; }
        public List<tbQA> lisQA { get; set; }

    }




}
