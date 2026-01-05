using KF_WebAPI.BaseClass.AE;
using KF_WebAPI.BaseClass;
using OfficeOpenXml;
using KF_WebAPI.DataLogic;
using System.Data;
using Microsoft.Data.SqlClient;
using System.Text.RegularExpressions;
using Newtonsoft.Json;

namespace KF_WebAPI.FunctionHandler
{
    public class KFCaseUpload
    {
        private readonly string _sheetNameMoto = "機車"; // Excel工作表名稱
        private readonly string _sheetNameHouse = "房子"; // Excel工作表名稱
        public List<Dictionary<string, string>> uploadExcelDataMoto;
        public List<Dictionary<string, string>> uploadExcelDataHouse;
        public List<KFChecklistD> KFChecklistDs;
        public List<NewKFChecklistD> NewKFChecklistDs;

        public KFCaseUpload()
        {
            uploadExcelDataMoto = new List<Dictionary<string, string>>();
            uploadExcelDataHouse = new List<Dictionary<string, string>>();
            KFChecklistDs = new List<KFChecklistD>();
            NewKFChecklistDs = new List<NewKFChecklistD>();
        }

        public async Task<ResultClass<string>> GetNewKFChecklistDs(NewKFChecklistRequest req)
        {
            // step1.讀取Excel內容
            ResultClass<string> result = await UpLoadExcelFile(req.file);
            if (result.ResultCode != "000")
                return result;

            // step2.讀取House_sender by 年月份
            result = GetKFNewChecklistDs(req.selYear_S);
            if (result.ResultCode != "000")
                return result;

            // step3.比較Excel與House_sender by 申請人姓名(要注意難字)
            DiffKF_NewChecklistDs();

            // step4.篩選資料
            if (req.isDiff == true)
            {
                // 如果打勾，則篩選出不相等或未對應的資料
                NewKFChecklistDs = NewKFChecklistDs.Where(x => x.isDiff || x.isMate || x.isHaveNCR).ToList();
            }
            result.objResult = JsonConvert.SerializeObject(NewKFChecklistDs);
            result.ResultCode = "000";
            return result;
        }

        public async Task<ResultClass<string>> UpLoadExcelFile(IFormFile file)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            if (file == null || file.Length == 0)
            {
                resultClass.ResultMsg = "請選擇檔案";
                resultClass.ResultCode = "500";
                return resultClass;
            }

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
                        // 第一個 Sheet → 讀 D、E 欄
                        ReadSheet(package, _sheetNameMoto, startRow: 2, columns: new int[] { 3, 4, 5, 7, 8 }, "1");

                        // 第二個 Sheet → 讀 C、D 欄
                        ReadSheet(package, _sheetNameHouse, startRow: 2, columns: new int[] { 3, 4, 6, 7, 9, 10, 11 }, "2");
                    }
                    catch (Exception ex)
                    {
                        resultClass.ResultMsg = "解析 Excel 發生錯誤 - " + ex.Message;
                        resultClass.ResultCode = "500";
                        return resultClass;
                    }
                }
            }

            resultClass.ResultCode = "000";
            return resultClass;
        }

        //style=1:機車,style=2:房子
        private void ReadSheet(ExcelPackage package, string sheetName, int startRow, int[] columns,string style)
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets[sheetName];
            if (worksheet == null)
                throw new Exception("找不到工作表：" + sheetName);

            int rowCount = worksheet.Dimension.Rows;

            for (int row = startRow; row <= rowCount; row++)
            {
                var rowDict = new Dictionary<string, string>();

                foreach (var col in columns)
                {
                    string header = worksheet.Cells[1, col].Text; // 標題
                    string value = worksheet.Cells[row, col].Text; // 值
                    rowDict[header] = value;
                }

                if (style == "1")
                {
                    if (!string.IsNullOrWhiteSpace(rowDict.Values.FirstOrDefault()))
                        uploadExcelDataMoto.Add(rowDict);
                }
                else if(style == "2")
                {
                    if (!string.IsNullOrWhiteSpace(rowDict.Values.FirstOrDefault()))
                        uploadExcelDataHouse.Add(rowDict);
                }
               
            }
        }

        public ResultClass<string> GetKFNewChecklistDs(string selYear_S)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            FuncHandler _FuncHandler = new FuncHandler();
            ADOData _ADO = new();
            var parameters = new List<SqlParameter>();
            var T_SQL = @"
                    SELECT
                        REPLACE(RTRIM(CS_name), '	', '') CS_name, /* 申請人:用來比對Key*/
                        ub.item_D_name U_BC_name, /* 區*/
                        U_name plan_name, /*業務*/
                        ip.item_D_name show_project_title, /* 專案*/
                        Loan_rate , /* 成數*/
                        interest_rate_pass,  /* 合約利率*/
                        H.get_amount /*金額*/
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
                    KFChecklistDs = dtResult.AsEnumerable().Select(row => new KFChecklistD
                    {
                        CS_name = _FuncHandler.DeCodeBNWords(row.Field<string>("CS_name")),
                        U_BC_name = row.Field<string>("U_BC_name"),
                        plan_name = _FuncHandler.DeCodeBNWords(row.Field<string>("plan_name")),
                        show_project_title = _FuncHandler.DeCodeBNWords(row.Field<string>("show_project_title")),
                        Loan_rate = string.IsNullOrEmpty(row.Field<string>("Loan_rate"))? "0%": Convert.ToDecimal(row.Field<string>("Loan_rate")).ToString("0.00") + "%",
                        interest_rate_pass = string.IsNullOrEmpty(row.Field<string>("interest_rate_pass")) ? "0%": row.Field<string>("interest_rate_pass") + "%",
                        get_amount = string.IsNullOrEmpty(row.Field<string>("get_amount")) ? "萬" : row.Field<string>("get_amount") + "萬"
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

        // 將上傳的Excel資料與KFNewXinChecklistD進行比對
        public void DiffKF_NewChecklistDs()
        {
            //1.檢查機車資料
            foreach(var item in uploadExcelDataMoto)
            {
                var matchItems = KFChecklistDs.Where(x => Regex.IsMatch(item["客戶"], x.GetRegexCS())).ToList();

                if (!matchItems.Any())
                {
                    NewKFChecklistD model = new NewKFChecklistD
                    {
                        CS_name_xls = item["客戶"],
                        U_BC_xls = item["區域"],
                        Pro_Na_xls = item["專案"],
                        get_amount_xls = item["金額"]
                    };
                    NewKFChecklistDs.Add(model);
                    continue;
                }

                foreach (var matchItem in matchItems)
                {
                    NewKFChecklistD model = new NewKFChecklistD
                    {
                        CS_name_xls = item["客戶"],
                        U_BC_xls = item["區域"],
                        Pro_Na_xls = item["專案"],
                        get_amount_xls = item["金額"]
                    };
                    model.CS_name = matchItem.CS_name;
                    model.U_BC_name = matchItem.U_BC_name;
                    model.plan_name = matchItem.plan_name;
                    model.show_project_title = matchItem.show_project_title;
                    model.get_amount = matchItem.get_amount;
                    NewKFChecklistDs.Add(model);
                }
            }
            //2.檢查房子資料
            foreach (var item in uploadExcelDataHouse)
            {
                var matchItems = KFChecklistDs.Where(x => Regex.IsMatch(item["客戶"], x.GetRegexCS())).ToList();

                if (!matchItems.Any())
                {
                    NewKFChecklistD model = new NewKFChecklistD
                    {
                        CS_name_xls = item["客戶"],
                        U_BC_xls = item["區域"],
                        Pro_Na_xls = item["專案"],
                        interest_rate_pass_xls = item["利率"],
                        Loan_rate_xls = item["成數"],
                        get_amount_xls = item["金額"]
                    };
                    NewKFChecklistDs.Add(model);
                    continue;
                }

                foreach (var matchItem in matchItems)
                {
                    NewKFChecklistD model = new NewKFChecklistD
                    {
                        CS_name_xls = item["客戶"],
                        U_BC_xls = item["區域"],
                        Pro_Na_xls = item["專案"],
                        interest_rate_pass_xls = item["利率"],
                        Loan_rate_xls = item["成數"],
                        get_amount_xls = item["金額"]
                    };
                    model.CS_name = matchItem.CS_name;
                    model.U_BC_name = matchItem.U_BC_name;
                    model.plan_name = matchItem.plan_name;
                    model.show_project_title = matchItem.show_project_title;
                    model.Loan_rate = matchItem.Loan_rate;
                    model.interest_rate_pass = matchItem.interest_rate_pass;
                    model.get_amount = matchItem.get_amount;
                    NewKFChecklistDs.Add(model);
                }
            }
            //foreach (var item in uploadExcelData)
            //{
            //    var matchItems = KFChecklistDs.Where(x => Regex.IsMatch(item["客戶"], x.GetRegexCS())).ToList();

            //    if (!matchItems.Any())
            //    {
            //        NewKFChecklistD model = new NewKFChecklistD
            //        {
            //            CS_name = item["客戶"]
            //        };
            //        newFChecklistDs.Add(model);
            //        continue;
            //    }

            //    foreach(var matchItem in matchItems) 
            //    {
            //        NewKFChecklistD model = new NewKFChecklistD
            //        {
            //            CS_name = item["客戶"]
            //        };
            //        model.CS_name = _FuncHandler.DeCodeBNWords(matchItem.CS_name);
            //        model.U_BC_name = matchItem.U_BC_name;
            //        model.plan_name = _FuncHandler.DeCodeBNWords(matchItem.plan_name);
            //        model.show_project_title = _FuncHandler.DeCodeBNWords(matchItem.show_project_title);
            //        model.Loan_rate = matchItem.Loan_rate;
            //        model.interest_rate_pass = matchItem.interest_rate_pass;
            //        newFChecklistDs.Add(model);
            //    }
            //}
        }
    }
}
