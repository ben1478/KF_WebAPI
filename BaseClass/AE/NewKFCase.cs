using KF_WebAPI.DataLogic;
using KF_WebAPI.FunctionHandler;
using Newtonsoft.Json;
using OfficeOpenXml;
using System.Data;
using System.Text.RegularExpressions;

namespace KF_WebAPI.BaseClass.AE
{
    #region 國峯案件核對表
    public class NewKFChecklistRequest
    {
        public IFormFile? file { get; set; } // 上傳的Excel檔案
        public string selYear_S { get; set; } = ""; // 撥款年月
        public bool isDiff { get; set; } = false; // // 打勾:true -filter 不相等or未對應or難字
    }

    public class KFChecklistD
    {
        public string? CS_name { get; set; } // 申請人
        public string U_BC_name { get; set; } // 區
        public string plan_name { get; set; } // 業務
        public string show_project_title { get; set; } // 專案
        public string Loan_rate { get; set; } // 貸款成數
        public string interest_rate_pass { get; set; } // 承作利率
        public string? get_amount { get; set; } // 金額

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

    public class NewKFChecklistD
    {
        public string CS_name_xls { get; set; } // 檔案名稱
        public string U_BC_xls { get; set; } // 據點
        public string Pro_Na_xls { get; set; } // 適用專案
        public string Loan_rate_xls { get; set; } // 成數
        public string interest_rate_pass_xls { get; set; } // 合約利率
        public string get_amount_xls { get; set; } //金額
        public string? CS_name { get; set; } // 申請人
        public string U_BC_name { get; set; } // 區
        public string plan_name { get; set; } // 業務
        public string show_project_title { get; set; } // 專案
        public string Loan_rate { get; set; } // 貸款成數
        public string interest_rate_pass { get; set; } // 承作利率
        public string get_amount { get; set; } //金額
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
        /// 判斷成數、利率、專案、金額是否相等
        /// </summary>
        public bool isDiff
        {
            get
            {
                bool result = false;
                if (string.IsNullOrEmpty(CS_name))
                {
                    return false;
                }
                if (string.IsNullOrEmpty(Loan_rate_xls))
                {
                    result = !(string.Equals(Pro_Na_xls, show_project_title, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(get_amount_xls, get_amount, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(U_BC_xls, U_BC_name, StringComparison.OrdinalIgnoreCase));
                }
                else
                {
                    result = !(string.Equals(Pro_Na_xls, show_project_title, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(get_amount_xls, get_amount, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(Loan_rate_xls, Loan_rate, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(interest_rate_pass_xls, interest_rate_pass, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(U_BC_xls, U_BC_name, StringComparison.OrdinalIgnoreCase));
                }
                
                return result;
            }
        }
    }
    #endregion

}
