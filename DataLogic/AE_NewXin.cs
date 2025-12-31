using KF_WebAPI.BaseClass;
using KF_WebAPI.BaseClass.AE;
using KF_WebAPI.FunctionHandler;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Data;
using System.Diagnostics;
using System.Net;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;

namespace KF_WebAPI.DataLogic
{
    public class NewXinChecklistRequest
    {
        // 新信貸核對清單請求類別
        // 負責接收前端傳入的參數
        public IFormFile? file { get; set; } // 上傳的Excel檔案
        public string selYear_S { get; set; } = ""; // 撥款年月
        public bool isDiff { get; set; } = false; // // 打勾:true -filter 不相等or未對應or難字
    }
    #region 新鑫案件核對表
    public class NewXinChecklistM
    {
        private readonly string _sheetName = "Sheet0"; // Excel工作表名稱

        // 上傳Excel資料
        public List<Dictionary<string, string>> uploadExcelData { get; set; }

        public List<NewXinChecklistD> newXinChecklistDs { get; set; } // 新信貸核對清單資料
        public async Task<ResultClass<string>> UpLoadExcelFile(IFormFile file)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            if (file == null || file.Length == 0)
            {
                resultClass.ResultMsg = "請選擇檔案";
                resultClass.ResultCode = "500";
                return resultClass;
            }

            uploadExcelData = new List<Dictionary<string, string>>();

            using (var stream = new MemoryStream())
            {
                try
                {
                    await file.CopyToAsync(stream);
                }
                catch (Exception ex)
                {
                    resultClass.ResultMsg = "檔案上傳失敗: " + ex.Message;
                    resultClass.ResultCode = "500";
                    return resultClass;
                }

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var package = new ExcelPackage(stream))
                {
                    try
                    {
                        ExcelWorksheet worksheet = package.FindExcelSheet(_sheetName);
                        if (worksheet == null)
                        {
                            resultClass.ResultMsg = "找不到工作表: " + _sheetName;
                            resultClass.ResultCode = "500";
                            return resultClass;
                        }

                        int rowCount = worksheet.Dimension.Rows;
                        int colCount = worksheet.Dimension.Columns;

                        // 讀取第一列標題
                        var headers = new List<string>();
                        for (int col = 1; col <= colCount; col++)
                            headers.Add(worksheet.Cells[1, col].Text);

                        // 從第二列開始讀資料
                        for (int row = 2; row <= rowCount; row++)
                        {
                            var rowDict = new Dictionary<string, string>();
                            for (int col = 1; col <= colCount; col++)
                            {
                                rowDict[headers[col - 1]] = worksheet.Cells[row, col].Text;
                            }
                            if (!string.IsNullOrEmpty(rowDict[headers[0]]))
                            {
                                uploadExcelData.Add(rowDict); // 確保
                            }

                        }
                    }
                    catch (Exception ex)
                    {
                        resultClass.ResultMsg = "找不到工作表: " + _sheetName + " - " + ex.Message;
                        resultClass.ResultCode = "500";
                        return resultClass;
                    }

                }
            }

            resultClass.ResultCode = "000";
            return resultClass;
        }

        // 將上傳的Excel資料與KFNewXinChecklistD進行比對
        public void DiffKF_NewXinChecklistDs(List<KFNewXinChecklistD> KFNewXinChecklistDs)
        {
            FuncHandler _FuncHandler = new FuncHandler();
            newXinChecklistDs = new List<NewXinChecklistD>();
            foreach (var item in uploadExcelData)
            {
                //原規則：刪除 CS_name_xls=總計
                if (item["申購人"] == "總計")
                    continue; // 跳過總計行

                //判斷申請者與House_sender.CS_name是否相同。但, House_sender.CS_name若為難字時會存在NCR字元，需先轉換
                var matchItems = KFNewXinChecklistDs
                                    .Where(x => Regex.IsMatch(item["申購人"], x.GetRegexCS()))
                                    .ToList();

                // 未對應
                if (!matchItems.Any())
                {
                    NewXinChecklistD newXinChecklistD = new NewXinChecklistD
                    {
                        CS_name_xls = item["申購人"],
                        U_BC_xls = item["據點"],
                        Pro_Na_xls = item["適用專案"],
                        Loan_rate_xls = string.IsNullOrEmpty(item["成數"]) ? 0 : Convert.ToDecimal(item["成數"]),
                        interest_rate_pass_xls = string.IsNullOrEmpty(item["合約\r\n利率"]) ? 0 : Convert.ToDecimal(item["合約\r\n利率"])
                    };
                    newXinChecklistDs.Add(newXinChecklistD);
                    continue; // 跳過未對應的項目
                }

                // 將匹配到的項目轉換為列表
                foreach (var matchItem in matchItems)
                {
                    NewXinChecklistD newXinChecklistD = new NewXinChecklistD
                    {
                        CS_name_xls = item["申購人"],
                        U_BC_xls = item["據點"],
                        Pro_Na_xls = item["適用專案"],
                        Loan_rate_xls = string.IsNullOrEmpty(item["成數"]) ? 0 : Convert.ToDecimal(item["成數"]),
                        interest_rate_pass_xls = string.IsNullOrEmpty(item["合約\r\n利率"]) ? 0 : Convert.ToDecimal(item["合約\r\n利率"])
                    };

                    newXinChecklistD.CS_name = _FuncHandler.DeCodeBNWords(matchItem.CS_name);
                    newXinChecklistD.U_BC_name = matchItem.U_BC_name;
                    newXinChecklistD.plan_name = _FuncHandler.DeCodeBNWords(matchItem.plan_name);
                    newXinChecklistD.show_project_title = _FuncHandler.DeCodeBNWords(matchItem.show_project_title);
                    newXinChecklistD.Loan_rate = matchItem.Loan_rate;
                    newXinChecklistD.interest_rate_pass = matchItem.interest_rate_pass;
                    newXinChecklistDs.Add(newXinChecklistD);
                }

            }
        }

        public async Task<ResultClass<string>> GetNewXinChecklistDs(NewXinChecklistRequest req)
        {
            // step1.讀取Excel內容
            ResultClass<string> result = await UpLoadExcelFile(req.file);
            if (result.ResultCode != "000")
                return result;

            // step2.讀取House_sender by 年月份
            KFNewXinChecklistM _KFNewXinChecklistM = new KFNewXinChecklistM();
            result = _KFNewXinChecklistM.GetKFNewXinChecklistDs(req.selYear_S);
            if (result.ResultCode != "000")
                return result;

            // step3.比較Excel與House_sender by 申請人姓名(要注意難字)
            DiffKF_NewXinChecklistDs(_KFNewXinChecklistM.KFNewXinChecklistDs);

            // step4.篩選資料
            if (req.isDiff == true)
            {
                // 如果打勾，則篩選出不相等或未對應的資料
                newXinChecklistDs = newXinChecklistDs.Where(x => x.isDiff || x.isMate || x.isHaveNCR).ToList();
            }

            result.ResultCode = "000";
            return result;
        }

        public ResultClass<string> GetNewXinChecklistDsYearMonth()
        {
            ResultClass<string> result = new ResultClass<string>();
            try
            {
                ADOData _adoData = new ADOData(); // 測試:"Test" / 正式:""
                #region SQL
                var T_SQL = @"select distinct 
                                convert(varchar(4),(convert(varchar(4), get_amount_date, 126)-1911))+'-'+convert(varchar(2),month(get_amount_date)) yyyMM,  
                                convert(varchar(7), get_amount_date, 126)yyyymm, 
                                case 
                                    when day(GETDATE()) <20 and year(DATEADD(month, -1, GETDATE())) = year(get_amount_date) and month(DATEADD(month, -1, GETDATE())) = month(get_amount_date) then 1
                                    when day(GETDATE()) >=20 and year(GETDATE()) = year(get_amount_date) and month(GETDATE()) = month(get_amount_date) then 1
                                    else
                                        0
                                end as op
                                from House_sendcase 
                                where get_amount_date>(DATEADD(month,-4, SYSDATETIME())) 
                                order by convert(varchar(7), get_amount_date, 126) desc";

                #endregion
                var dtResult = _adoData.ExecuteSQuery(T_SQL);

                if (dtResult.Rows.Count > 0)
                {
                    result.ResultCode = "000";
                    result.objResult = JsonConvert.SerializeObject(dtResult);
                }
                else
                {
                    result.ResultCode = "400";
                    result.ResultMsg = "查無資料";
                }
            }
            catch (Exception ex)
            {
                result.ResultCode = "500";
                result.ResultMsg = $" response: {ex.Message}";
            }
            return result;
        }
    }
    public class NewXinChecklistD
    {
        public string CS_name_xls { get; set; } // 檔案名稱
        public string U_BC_xls { get; set; } // 據點
        public string Pro_Na_xls { get; set; } // 適用專案
        public decimal Loan_rate_xls { get; set; } // 成數
        public decimal interest_rate_pass_xls { get; set; } // 合約利率
        public string? CS_name { get; set; } // 申請人
        public string U_BC_name { get; set; } // 區
        public string plan_name { get; set; } // 業務
        public string show_project_title { get; set; } // 專案
        public decimal Loan_rate { get; set; } // 貸款成數
        public decimal interest_rate_pass { get; set; } // 承作利率

        /// <summary>
        /// 判斷是否為難字:難字:true
        /// </summary>
        public bool isHaveNCR
        {
            get
            {
                bool result = false;
                if (string.IsNullOrEmpty(CS_name_xls) == false)
                {
                    result = CS_name_xls.Contains("&#") || CS_name_xls.Contains('?'); // 判斷是否包含NCR字元    
                }

                return result;
            }
        }

        public bool isMate // 未對應
        {
            get
            {
                bool result = string.IsNullOrEmpty(CS_name);
                return result;
            }
        }

        /// <summary>
        /// 判斷成數和利率是否相等
        /// </summary>
        public bool isDiff
        {
            get
            {
                bool result = false;
                if (string.IsNullOrEmpty(CS_name))
                {
                    // 如果CS_name為空，則不進行成數和利率的比較
                    return false;
                }
                result = !(Decimal.Equals(Loan_rate_xls, Loan_rate)
                        && Decimal.Equals(interest_rate_pass_xls, interest_rate_pass));
                //Debug.WriteLine($"{isDiff}--{Loan_rate_xls}-{Loan_rate}-{Decimal.Equals(Loan_rate_xls, Loan_rate)}/{interest_rate_pass_xls}-{interest_rate_pass}-{Decimal.Equals(interest_rate_pass_xls, interest_rate_pass)}");
                return result;
            }
        } // 不相等(成數,利率)

    }

    public class KFNewXinChecklistM
    {

        public List<KFNewXinChecklistD> KFNewXinChecklistDs { get; set; } // KF House_pre
        /// <summary>
        /// 
        /// </summary>
        /// <param name="selYear_S">撥款年月</param>
        /// <returns></returns>
        public ResultClass<string> GetKFNewXinChecklistDs(string selYear_S)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            ADOData _ADO = new();
            var parameters = new List<SqlParameter>();
            var T_SQL = @"
                    SELECT
                        REPLACE(RTRIM(CS_name), '	', '') CS_name, /* 申請人:用來比對Key*/
                        ub.item_D_name U_BC_name, /*  區*/
                        U_name plan_name, /* 業務*/
                        ip.item_D_name show_project_title, /*  專案*/
                        Loan_rate , /*  成數*/
                        interest_rate_pass  /*  合約利率*/
                    FROM House_sendcase as H
                    LEFT JOIN Item_list I ON fund_company = item_D_code and item_M_code = 'fund_company' AND item_D_type = 'Y'
                    LEFT JOIN House_apply A ON A.HA_id = H.HA_id AND A.del_tag = '0'          
                    left join House_pre_project P on P.HP_project_id = H.HP_project_id and P.del_tag = '0'
                    left join Item_list IP on  P.project_title = IP.item_D_code and  IP.item_M_code = 'project_title' AND IP.item_D_type = 'Y'
                    left join User_M u on A.plan_num = u.u_num
                    left join Item_list ub on ub.item_M_code = 'branch_company' and ub.item_D_type = 'Y' AND ub.item_D_code = u.U_BC
                        WHERE (CONVERT(VARCHAR(7), (get_amount_date), 126) = @selYear_S)
                ";
            parameters.Add(new SqlParameter("@selYear_S", selYear_S));
            try
            {
                DataTable dtResult = _ADO.ExecuteQuery(T_SQL, parameters);
                if (dtResult.Rows.Count > 0)
                {
                    KFNewXinChecklistDs = dtResult.AsEnumerable().Select(row => new KFNewXinChecklistD
                    {
                        CS_name = row.Field<string>("CS_name"),
                        U_BC_name = row.Field<string>("U_BC_name"),
                        plan_name = row.Field<string>("plan_name"),
                        show_project_title = row.Field<string>("show_project_title"),
                        Loan_rate = string.IsNullOrEmpty(row.Field<string>("Loan_rate")) ? 0 : Convert.ToDecimal(row.Field<string>("Loan_rate")),
                        interest_rate_pass = string.IsNullOrEmpty(row.Field<string>("Loan_rate")) ? 0 : Convert.ToDecimal(row.Field<string>("interest_rate_pass"))
                    }).ToList();
                }
            }
            catch (Exception ex)
            {
                resultClass.ResultMsg = "查詢失敗: " + ex.Message;
                resultClass.ResultCode = "500";
                return resultClass;
            }

            resultClass.ResultCode = "000";
            return resultClass;
        }

        public string? GetCname(List<int> HA_ids)
        {
            var inList = "(" + string.Join(", ", HA_ids) + ")";
            ResultClass<string> resultClass = new ResultClass<string>();
            ADOData _ADO = new();
            var parameters = new List<SqlParameter>();
            var T_SQL = @"
                    SELECT CS_name
                        FROM House_apply
                        WHERE HA_id IN  " + inList + " ORDER BY HA_id DESC ";

            try
            {
                DataTable dtResult = _ADO.ExecuteQuery(T_SQL, parameters);
                if (dtResult.Rows.Count > 0)
                {
                    string? ltCName = dtResult.AsEnumerable().Select(row => row.Field<string>("CS_name")).FirstOrDefault();
                    return ltCName;
                }
            }
            catch (Exception ex)
            {
                resultClass.ResultMsg = "查詢失敗: " + ex.Message;
            }

            return null;
        }
    }
    public class KFNewXinChecklistD
    {
        public string? CS_name { get; set; } // 申請人
        public string U_BC_name { get; set; } // 區
        public string plan_name { get; set; } // 業務
        public string show_project_title { get; set; } // 專案
        public decimal Loan_rate { get; set; } // 貸款成數
        public decimal interest_rate_pass { get; set; } // 承作利率

        // 取得regex格式的CS_name
        public string GetRegexCS()
        {
            if (string.IsNullOrEmpty(CS_name))
                return string.Empty;

            FuncHandler funcHandler = new FuncHandler();
            string processedStr = CS_name;

            //判定格式
            if (CS_name.Contains("&#"))
            {
                string fixedCSName = funcHandler.FixNCRIfNeeded(CS_name);//補分號
                processedStr = funcHandler.fromNCR(fixedCSName);//NCR字元轉換為正常字元
            }
            else
            {
                processedStr = funcHandler.DeCodeBNWords(CS_name);
            }

            return processedStr;             
        }

        
    }
    #endregion
    #region 新鑫已撥款明細表
    public class NewXinCaseStatusM
    {
        public List<NewXinCaseStatus_res> GetNewXinCaseStatusDs(NewXinCaseStatus_req model)
        {
            FuncHandler _FuncHandler = new FuncHandler();
            List<NewXinCaseStatus_res> newXinCaseStatusList = new List<NewXinCaseStatus_res>();
            var sdate = DateTime.Parse(FuncHandler.ConvertROCToGregorian(model.start_date));
            var edate = DateTime.Parse(FuncHandler.ConvertROCToGregorian(model.end_date));

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"
                        SELECT
                          ISNULL(CS_PID, '') CS_ID
                         ,ISNULL(CS_register_address, '') addr
                         ,ISNULL(Loan_rate, '') Loan_rate
                         ,REPLACE(REPLACE(CONVERT(VARCHAR(10), Send_amount_date, 126), '-', '/'), YEAR(Send_amount_date), YEAR(Send_amount_date) - 1911) Send_amount_date
                         ,REPLACE(REPLACE(CONVERT(VARCHAR(10), get_amount_date, 126), '-', '/'), YEAR(get_amount_date), YEAR(get_amount_date) - 1911) get_amount_date
                         ,H.HS_id
                         ,H.pass_amount
                         ,H.Send_amount
                         ,H.get_amount
                         ,House_pre_project.project_title
                         ,House_apply.CS_name
                         ,House_apply.CS_MTEL1
                         ,House_apply.CS_MTEL2
                         ,House_apply.CS_introducer
                         ,User_M.U_name plan_name
                         ,User_M.U_BC_name
                         ,ISNULL((SELECT
                              item_D_name
                            FROM Item_list
                            WHERE item_M_code = 'Send_result_type'
                            AND item_D_type = 'Y'
                            AND item_D_code = H.Send_result_type
                            AND show_tag = '0'
                            AND del_tag = '0')
                          , '--') AS show_Send_result_type
                         ,ISNULL((SELECT
                              item_D_name
                            FROM Item_list
                            WHERE item_M_code = 'check_amount_type'
                            AND item_D_type = 'Y'
                            AND item_D_code = H.check_amount_type
                            AND show_tag = '0'
                            AND del_tag = '0')
                          , '--') AS show_check_amount_type
                         ,ISNULL((SELECT
                              item_D_name
                            FROM Item_list
                            WHERE item_M_code = 'get_amount_type'
                            AND item_D_type = 'Y'
                            AND item_D_code = H.get_amount_type
                            AND show_tag = '0'
                            AND del_tag = '0')
                          , '--') AS show_get_amount_type
                         ,(SELECT
                              item_D_name
                            FROM Item_list
                            WHERE item_M_code = 'appraise_company'
                            AND item_D_type = 'Y'
                            AND item_D_code = H.appraise_company
                            AND show_tag = '0'
                            AND del_tag = '0')
                          AS show_appraise_company
                         ,(SELECT
                              item_D_name
                            FROM Item_list
                            WHERE item_M_code = 'fund_company'
                            AND item_D_type = 'Y'
                            AND item_D_code = H.fund_company
                            AND show_tag = '0'
                            AND del_tag = '0')
                          AS show_fund_company
                         ,(SELECT
                              item_D_name
                            FROM Item_list
                            WHERE item_M_code = 'project_title'
                            AND item_D_type = 'Y'
                            AND item_D_code = House_pre_project.project_title
                            AND show_tag = '0'
                            AND del_tag = '0')
                          AS show_project_title
                         ,Users.U_name AS fin_name
                         ,CONVERT(VARCHAR, house_pre_project.fin_date, 111) AS fin_date
                        FROM House_sendcase H
                        LEFT JOIN House_apply
                          ON House_apply.HA_id = H.HA_id
                            AND House_apply.del_tag = '0'
                        LEFT JOIN House_pre_project
                          ON House_pre_project.HP_project_id = H.HP_project_id
                            AND House_pre_project.del_tag = '0'
                        LEFT JOIN view_User_sales User_M
                          ON User_M.U_num = House_apply.plan_num
                        LEFT JOIN view_user_sales Users
                          ON house_pre_project.fin_user = Users.U_num
                        WHERE H.del_tag = '0'
                        AND H.sendcase_handle_type = 'Y'
                        AND ISNULL(H.Send_amount, '') <> ''
                        AND (Send_amount_date >= @start_date
                        AND Send_amount_date <= @end_date)
                        AND fund_Company = 'FDCOM001'
                        AND CONVERT(VARCHAR(7), get_amount_date, 126) = @selYear_S
                        AND get_amount_type = 'GTAT002'
                        ORDER BY Send_amount_date, H.HS_id DESC
                    ";
                parameters.Add(new SqlParameter("@start_date", sdate));
                parameters.Add(new SqlParameter("@end_date", edate));
                parameters.Add(new SqlParameter("@selYear_S", model.selYear_S));
                #endregion

                DataTable dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
                if (dtResult.Rows.Count > 0)
                {
                    newXinCaseStatusList = dtResult.AsEnumerable().Select(row => new NewXinCaseStatus_res
                    {
                        HS_id = row.Field<decimal>("HS_id").ToString(),
                        Send_amount_date = row.IsNull("Send_amount_date") ? "" :
                                         row.Field<string>("Send_amount_date"),
                        get_amount_date = row.IsNull("get_amount_date") ? "" :
                                         row.Field<string>("get_amount_date"),
                        CS_name =  row.IsNull("CS_name") ? "" : _FuncHandler.DeCodeBNWords(row.Field<string>("CS_name")),
                        CS_ID = row.Field<string>("CS_ID"),
                        show_project_title = row.IsNull("show_project_title") ? "" : _FuncHandler.DeCodeBNWords(row.Field<string>("show_project_title")),
                        get_amount = row.Field<string>("get_amount")
                    }).ToList();
                }
            }
            catch (Exception ex)
            {
                throw new Exception("查詢失敗: " + ex.Message);
            }

            return newXinCaseStatusList;
        }
        public byte[] GetGetNewXinCaseStatusExcel(NewXinCaseStatus_req model)
        {
            var Excel_Headers = new Dictionary<string, string>
            {
                { "Send_amount_date","出件日期" },
                { "get_amount_date", "撥款日期" },
                { "CS_name", "申請人" },
                { "CS_ID", "申請人ID" },
                { "show_project_title", "出件方案" },
                { "get_amount", "撥款金額(萬)" }
            };
            List<NewXinCaseStatus_Excel> excelList = new List<NewXinCaseStatus_Excel>();
            byte[]? fileBytes = null;
            try
            {
                List<NewXinCaseStatus_res> newXinCaseStatus = GetNewXinCaseStatusDs(model);
                if (newXinCaseStatus.Count > 0)
                {
                    excelList = newXinCaseStatus.Select(row => new NewXinCaseStatus_Excel
                    {
                        Send_amount_date = row.Send_amount_date,
                        get_amount_date = row.get_amount_date,
                        CS_name = row.CS_name,
                        CS_ID = row.CS_ID,
                        show_project_title = row.show_project_title,
                        get_amount = row.get_amount
                    }).ToList();

                    fileBytes = FuncHandler.ExportToExcel(excelList, Excel_Headers);
                    return fileBytes;
                }
            }
            catch (Exception ex)
            {
                throw new Exception("查詢失敗: " + ex.Message);
            }
            return null;
        }
        public ResultClass<string> UpdateNewXinCaseStatus(string HS_ID, string Amount_Date)
        {
            var sdate = DateTime.Parse(FuncHandler.ConvertROCToGregorian(Amount_Date.Replace('/', '-')));
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL
                var T_SQL = @"
                                update House_sendcase
                                    set Send_amount_date = @amount_date
                                    from House_sendcase
                                    where HS_id = @hs_id
                    ";

                var parameters = new List<SqlParameter>
                {
                    new SqlParameter("@amount_date", sdate),
                    new SqlParameter("@hs_id", HS_ID)
                };

                #endregion
                int result = _adoData.ExecuteNonQuery(T_SQL, parameters);
                if (result == 0)
                {
                    resultClass.ResultCode = "400";
                    resultClass.ResultMsg = "異動失敗";
                }
                else
                {
                    resultClass.ResultCode = "000";
                    resultClass.ResultMsg = "異動成功";
                }
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "500";
                resultClass.ResultMsg = $" response: {ex.Message}";
                
            }
            return resultClass;
        }
    }
    #endregion 
    public static class ExcelPackageExtensions
    {
        public static ExcelWorksheet FindExcelSheet(this ExcelPackage package, string sheetName)
        {
            if (package == null || package.Workbook == null || package.Workbook.Worksheets == null)
                return null;
            foreach (var worksheet in package.Workbook.Worksheets)
            {
                if (worksheet.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                    return worksheet;
            }
            return null;
        }
    }
    #region 難字相關處理(已移至FuncHandler.cs)
    /*
    /// <summary>
    /// 檢查中文字是否為難字或不常用字
    /// </summary>
    public static class EncodingChecker
    {
        static EncodingChecker()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }
        public static List<(string Character, bool IsRare)> CheckRareChineseCharacters(string text)
        {
            var result = new List<(string Character, bool IsRare)>();
            for (int i = 0; i < text.Length; i++)
            {
                int codePoint = char.ConvertToUtf32(text, i);

                // 如果是 surrogate pair，要跳過下個 char（它已包含在 codePoint 中）
                if (char.IsHighSurrogate(text[i]) && i + 1 < text.Length && char.IsLowSurrogate(text[i + 1]))
                    i++; // skip next
                bool isRare = IsRareChineseCodePoint(codePoint) || IsDifficultChineseChar(text[i].ToString());
                result.Add((text[i].ToString(), isRare));
            }
            return result;
        }


        public static bool IsRareChineseCodePoint(int codePoint)
        {
            // 常見中文字區：U+4E00 ~ U+9FFF
            // 擴展A區：U+3400 ~ U+4DBF
            // 超出這範圍的中文通常屬於難字

            return !((codePoint >= 0x3400 && codePoint <= 0x4DBF) ||
                      (codePoint >= 0x4E00 && codePoint <= 0x9FFF));
        }

        public static bool IsDifficultChineseChar(string s)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            var big5 = Encoding.GetEncoding(950, new EncoderExceptionFallback(), new DecoderExceptionFallback());

            foreach (var rune in s.EnumerateRunes()) // 支援 UTF-32 字元
            {
                int codePoint = rune.Value;

                // 若超出 CJK 常見範圍
                if (!((codePoint >= 0x3400 && codePoint <= 0x4DBF) ||
                       (codePoint >= 0x4E00 && codePoint <= 0x9FFF)))
                {
                    return true; // 難字
                }

                // 嘗試 Big5 編碼（失敗則為難字）
                try
                {
                    big5.GetBytes(rune.ToString());
                }
                catch
                {
                    return true; // 難字
                }
            }

            return false; // 全部是常用可編碼字
        }

    }
    public class HKSCSMappingM
    {
        /// <summary>
        ///  取得HKSCS對應的Unicode字元
        /// </summary>
        /// <param name="hkscs"></param>
        /// <returns></returns>
        public ResultClass<string> GetUniCode(string hkscs)
        {
            ResultClass<string> result = new ResultClass<string>();
            try
            {
                ADOData _adoData = new ADOData(); // 測試:"Test" / 正式:""
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"select [ISO10646] from [HkscsMapping] where [hkscs2008] = @hkscs";
                parameters.Add(new SqlParameter("@hkscs", hkscs));
                #endregion
                var dtResult = _adoData.ExecuteQuery(T_SQL, parameters);

                if (dtResult.Rows.Count > 0)
                {
                    result.ResultCode = "000";
                    result.objResult = dtResult.AsEnumerable().ToList().Select(row => row.Field<string>("ISO10646")).FirstOrDefault();
                }
                else
                {
                    result.ResultCode = "400";
                    result.ResultMsg = "查無資料";
                }
            }
            catch (Exception ex)
            {
                result.ResultCode = "500";
                result.ResultMsg = $" response: {ex.Message}";
            }
            return result;
        }

        /// <summary>
        /// 將Big5編碼的字串解碼為Unicode字元
        /// </summary>
        /// <param name="strBig5"></param>
        /// <returns></returns>
        public string DeCodeBig5Words(string strBig5)
        {
            FuncHandler _FuncHandler = new FuncHandler();
            HKSCSMappingM hKSCSMappingM = new HKSCSMappingM();
            string decodeWords = string.Empty;
            var words = FuncHandler.CheckRareChineseCharacters(strBig5);
            foreach (var word in words)
            {
                if (word.IsRare)
                {
                    string HtmlEncode = FuncHandler.ConvertDecimalToUrlEncoded(word.Character);
                    ResultClass<string> resUniCode = _FuncHandler.GetUniCode(HtmlEncode.Replace("%", ""));
                    if (resUniCode.ResultCode != "000")
                    {
                        //Debug.WriteLine($"錯誤: {resUniCode.ResultMsg}");
                        continue;
                    }
                    string result = FuncHandler.ConvertHexToUnicodeChar(resUniCode.objResult);
                    //Debug.WriteLine($"解碼後字元：{result}");
                    decodeWords += result; // 將解碼後的字元累加
                }
                else
                {
                    //Debug.WriteLine($"✅ [{word.Character}] 是常見中文字");
                    decodeWords += word.Character; // 將解碼後的字元累加
                }
            }
            return decodeWords;
        }
    }
    public class UnicodeToBig5UrlRecover
    {
        public static string ConvertDecimalToUrlEncoded(string str)
        {
            StringBuilder sb = new StringBuilder();

            for (int i = 0; i < str.Length; i += Char.IsSurrogatePair(str, i) ? 2 : 1)
            {
                int codePoint = Char.ConvertToUtf32(str, i);
                var decimalInput = Char.ConvertFromUtf32(codePoint); //Converts the specified Unicode code point into a UTF-16 encoded string

                Debug.WriteLine("U+{0:X4} {1} at {2}", codePoint, decimalInput, i);

                char c = (char)codePoint;                          // U+E1D0

                Encoding big5 = Encoding.GetEncoding(950);         // Big5 編碼
                byte[] bytes = big5.GetBytes(new[] { c });         // 取得原始 Big5 bytes

                // 轉成 %XX%YY

                foreach (byte b in bytes)
                {
                    sb.Append('%');
                    sb.Append(b.ToString("X2"));
                }
            }
            return sb.ToString();
        }
    }
    public class Big5HkscsDecoder
    {
        public static string ConvertHexToUnicodeChar(string hex)
        {
            // 將十六進位字串轉成整數
            int codePoint = int.Parse(hex, System.Globalization.NumberStyles.HexNumber);

            // 將 codePoint 轉為 Unicode 字元
            string result = char.ConvertFromUtf32(codePoint);

            return result;
        }

    }
    */
    #endregion
}
