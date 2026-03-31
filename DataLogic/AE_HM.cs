using Azure;
using KF_WebAPI.BaseClass;
using KF_WebAPI.BaseClass.AE;
using KF_WebAPI.BaseClass.Max104;
using KF_WebAPI.BaseClass.Winton;
using KF_WebAPI.Controllers;
using KF_WebAPI.FunctionHandler;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections;
using System.Data;
using System.Drawing;
using System.Dynamic;
using System.Globalization;
using System.Reflection;
using System.Reflection.PortableExecutable;

namespace KF_WebAPI.DataLogic
{
    public class AE_HM
    {
        ADOData _adoData = new ADOData();
        FuncHandler _Fun = new FuncHandler();

        #region 客訴資料維護
        public ResultClass<string> Complaint_LQuery(Complaint_M_req model)
        {
            var parameters = new List<SqlParameter>();
            ResultClass<string> resultClass = new();
            try
            {
                var T_SQL = @"SELECT Comp_Id,CompDate,ub.item_D_name as CompSou_show,ContractID,CS_name,CS_PID,Um.U_name,
                              Remark,Cs.item_D_name as Risk_show,ResponsWay,ResponsStates,CASE WHEN IsClose='N' THEN '尚未結案' ELSE '結案' END as Close_show,CloseDate 
                              FROM Complaint Ct 
                              LEFT JOIN User_M Um ON Ct.Sales_num = Um.u_num
                              LEFT JOIN Item_list Ub ON ub.item_M_code = 'CompSou'  AND ub.item_D_code = Ct.CompSou
                              LEFT JOIN Item_list Cs ON Cs.item_M_code = 'RiskLevel'  AND Cs.item_D_code = Ct.RiskLevel
                              WHERE 1 = 1 ";

                if (model.IsClose != "A")
                {
                    T_SQL += @" AND IsClose = @IsClose";
                    parameters.Add(new SqlParameter("@IsClose", model.IsClose));
                }
                if (!string.IsNullOrWhiteSpace(model.CheckDateS) && !string.IsNullOrWhiteSpace(model.CheckDateE))
                {
                    T_SQL += @" AND CompDate BETWEEN @CheckDateS and @CheckDateE";
                    parameters.Add(new SqlParameter("@CheckDateS", FuncHandler.ConvertROCToGregorian(model.CheckDateS)));
                    parameters.Add(new SqlParameter("@CheckDateE", FuncHandler.ConvertROCToGregorian(model.CheckDateE)));
                }
                if (!string.IsNullOrWhiteSpace(model.CS_name))
                {
                    T_SQL += @" AND CS_name = @CS_name";
                    parameters.Add(new SqlParameter("@CS_name", model.CS_name));
                }
                if (!string.IsNullOrWhiteSpace(model.CS_PID))
                {
                    T_SQL += @" AND CS_PID = @CS_PID";
                    parameters.Add(new SqlParameter("@CS_PID", model.CS_PID));
                }
                if (!string.IsNullOrWhiteSpace(model.Sales_num))
                {
                    T_SQL += @" AND Sales_num = @Sales_num";
                    parameters.Add(new SqlParameter("@Sales_num", model.Sales_num));
                }
                if (!string.IsNullOrWhiteSpace(model.RiskLevel))
                {
                    T_SQL += @" AND RiskLevel = @RiskLevel";
                    parameters.Add(new SqlParameter("@RiskLevel", model.RiskLevel));
                }
                if (!string.IsNullOrWhiteSpace(model.CompSou))
                {
                    T_SQL += @" AND CompSou = @CompSou";
                    parameters.Add(new SqlParameter("@CompSou", model.CompSou));
                }
                if (model.RoleType == "RoleB")
                {
                    T_SQL += @" AND C.Sales_num = @UserNum";
                    parameters.Add(new SqlParameter("@UserNum", model.User));
                }
                if (model.RoleType == "RoleC")
                {
                    T_SQL += @" AND M1.U_BC = @UserBC";
                    parameters.Add(new SqlParameter("@UserBC", model.U_BC));
                }

                var result = _adoData.ExecuteQuery(T_SQL, parameters).AsEnumerable().Select(row => new
                {
                    Comp_Id = row.Field<decimal>("Comp_Id"),
                    CompDate = row.Field<string>("CompDate"),
                    CompSou_show = row.Field<string>("CompSou_show"),
                    ContractID = row.Field<string>("ContractID"),
                    CS_name = _Fun.DeCodeBNWords(row.Field<string>("CS_name")),
                    CS_PID = row.Field<string>("CS_PID"),
                    U_name = row.Field<string>("U_name"),
                    Remark = row.Field<string>("Remark"),
                    Risk_show = row.Field<string>("Risk_show"),
                    ResponsWay = row.Field<string>("ResponsWay"),
                    ResponsStates = row.Field<string>("ResponsStates"),
                    Close_show = row.Field<string>("Close_show"),
                    CloseDate = row.Field<string>("CloseDate")
                }).ToList();

                resultClass.objResult = JsonConvert.SerializeObject(result);
                return resultClass;
            }
            catch (Exception)
            {
                throw;
            }
        }

        public int Complaint_Ins(Complaint_M model)
        {
            int result = 0;
            try
            {
                var T_SQL = @"Insert into Complaint (CS_name,CS_PID,Sales_num,PassID,ContractID,CompDate,CompSou,RiskLevel,Complaint,ResponsBC,Remark,
                              ResponsWay,ResponsStates,CloseDate,IsClose,add_date,add_num) 
                              Values (@CS_name,@CS_PID,@Sales_num,@PassID,@ContractID,@CompDate,
                              @CompSou,@RiskLevel,@Complaint,@ResponsBC,@Remark,@ResponsWay,@ResponsStates,@CloseDate,@IsClose,getdate(),@add_num)";
                var parameters = new List<SqlParameter>
                {
                    new SqlParameter("@CS_name",model.CS_name),
                    new SqlParameter("@CS_PID",model.CS_PID),
                    new SqlParameter("@Sales_num",string.IsNullOrWhiteSpace(model.Sales_num)? (object)DBNull.Value:model.Sales_num),
                    new SqlParameter("@PassID",string.IsNullOrWhiteSpace(model.PassID) ?(object) DBNull.Value : model.PassID),
                    new SqlParameter("@ContractID",string.IsNullOrWhiteSpace(model.ContractID) ?(object) DBNull.Value : model.ContractID),
                    new SqlParameter("@CompDate",FuncHandler.ConvertROCToGregorian(model.CompDate)),
                    new SqlParameter("@CompSou",model.CompSou),
                    new SqlParameter("@RiskLevel",model.RiskLevel),
                    new SqlParameter("@Complaint",model.Complaint),
                    new SqlParameter("@ResponsBC",string.IsNullOrWhiteSpace(model.ResponsBC)?(object)DBNull.Value : model.ResponsBC),
                    new SqlParameter("@Remark",model.Remark),
                    new SqlParameter("@ResponsWay",string.IsNullOrWhiteSpace(model.ResponsWay) ?(object) DBNull.Value : model.ResponsWay),
                    new SqlParameter("@ResponsStates",string.IsNullOrWhiteSpace(model.ResponsStates) ?(object) DBNull.Value : model.ResponsStates),
                    new SqlParameter("@CloseDate",string.IsNullOrWhiteSpace(model.CloseDate)? (object)DBNull.Value: FuncHandler.ConvertROCToGregorian(model.CloseDate)),
                    new SqlParameter("@IsClose",model.IsClose),
                    new SqlParameter("@add_num",model.tbInfo.add_num)
                };
                result = _adoData.ExecuteNonQuery(T_SQL, parameters);
                return result;
            }
            catch (Exception)
            {
                throw;
            }
        }

        public ResultClass<string> Complaint_SQuery(string Comp_Id)
        {
            ResultClass<string> resultClass = new();

            try
            {
                var T_SQL = @"SELECT Comp_Id,CompDate,CompSou,ContractID,CS_name,CS_PID,Um.U_name,Remark,RiskLevel,PassID,
                              ResponsWay,ResponsBC,ResponsStates,IsClose,Um.U_name,Sales_num,CloseDate,Complaint
                              FROM Complaint Ct 
                              LEFT JOIN User_M Um ON Ct.Sales_num = Um.u_num
                              WHERE Comp_Id = @Comp_Id";
                var parameters = new List<SqlParameter>()
                {
                    new SqlParameter("@Comp_Id",Comp_Id)
                };
                var result = _adoData.ExecuteQuery(T_SQL, parameters).AsEnumerable().Select(row => new
                {
                    Comp_Id = row.Field<decimal>("Comp_Id"),
                    CompDate = FuncHandler.ConvertGregorianToROC(row.Field<string>("CompDate")),
                    Complaint = row.Field<string>("Complaint"),
                    CompSou = row.Field<string>("CompSou"),
                    ContractID = row.Field<string>("ContractID"),
                    CS_name = _Fun.DeCodeBNWords(row.Field<string>("CS_name")),
                    CS_PID = row.Field<string>("CS_PID"),
                    U_num = row.Field<string>("Sales_num"),
                    U_name = row.Field<string>("U_name"),
                    Remark = row.Field<string>("Remark"),
                    PassID = row.Field<string>("PassID"),
                    RiskLevel = row.Field<string>("RiskLevel"),
                    ResponsWay = row.Field<string>("ResponsWay"),
                    ResponsBC = row.Field<string>("ResponsBC"),
                    ResponsStates = row.Field<string>("ResponsStates"),
                    IsClose = row.Field<string>("IsClose"),
                    CloseDate = row.Field<string>("CloseDate") == null ? "" : FuncHandler.ConvertGregorianToROC(row.Field<string>("CloseDate"))
                }).FirstOrDefault();

                resultClass.objResult = JsonConvert.SerializeObject(result);
                return resultClass;
            }
            catch (Exception)
            {
                throw;
            }
        }

        public int Complaint_Upd(Complaint_M model)
        {
            int result = 0;
            var parameters = new List<SqlParameter>();
            try
            {
                var T_SQL = @"Update Complaint Set CS_name = @CS_name ,CS_PID = @CS_PID ,Sales_num = @Sales_num ,PassID = @PassID ,ContractID = @ContractID ,
                              CompDate = @CompDate ,CompSou = @CompSou ,RiskLevel = @RiskLevel ,Complaint = @Complaint ,ResponsBC = @ResponsBC ,Remark = @Remark ,
                              ResponsWay = @ResponsWay ,ResponsStates = @ResponsStates ,CloseDate = @CloseDate ,IsClose = @IsClose, edit_date = getdate(), edit_num = @add_num";

                //如果結案 需計算處理時效(天) DealDay
                if (model.IsClose == "Y")
                {
                    DateTime d1 = DateTime.Parse(FuncHandler.ConvertROCToGregorian(model.CompDate));
                    DateTime d2 = DateTime.Parse(FuncHandler.ConvertROCToGregorian(model.CloseDate));
                    int days = (d2 - d1).Days;

                    T_SQL += @" ,DealDay = @DealDay";
                    parameters.Add(new SqlParameter("@DealDay", days));
                }

                T_SQL += @" Where Comp_Id = @Comp_Id";

                parameters.Add(new SqlParameter("@Comp_Id", model.Comp_Id));
                parameters.Add(new SqlParameter("@CS_name", model.CS_name));
                parameters.Add(new SqlParameter("@CS_PID", model.CS_PID));
                parameters.Add(new SqlParameter("@Sales_num", string.IsNullOrWhiteSpace(model.Sales_num) ? (object)DBNull.Value : model.Sales_num));
                parameters.Add(new SqlParameter("@PassID", string.IsNullOrWhiteSpace(model.PassID) ? (object)DBNull.Value : model.PassID));
                parameters.Add(new SqlParameter("@ContractID", string.IsNullOrWhiteSpace(model.ContractID) ? (object)DBNull.Value : model.ContractID));
                parameters.Add(new SqlParameter("@CompDate", FuncHandler.ConvertROCToGregorian(model.CompDate)));
                parameters.Add(new SqlParameter("@CompSou", model.CompSou));
                parameters.Add(new SqlParameter("@RiskLevel", model.RiskLevel));
                parameters.Add(new SqlParameter("@Complaint", model.Complaint));
                parameters.Add(new SqlParameter("@ResponsBC", string.IsNullOrWhiteSpace(model.ResponsBC) ? (object)DBNull.Value : model.ResponsBC));
                parameters.Add(new SqlParameter("@Remark", model.Remark));
                parameters.Add(new SqlParameter("@ResponsWay", string.IsNullOrWhiteSpace(model.ResponsWay) ? (object)DBNull.Value : model.ResponsWay));
                parameters.Add(new SqlParameter("@ResponsStates", string.IsNullOrWhiteSpace(model.ResponsStates) ? (object)DBNull.Value : model.ResponsStates));
                parameters.Add(new SqlParameter("@CloseDate", string.IsNullOrWhiteSpace(model.CloseDate) ? (object)DBNull.Value : FuncHandler.ConvertROCToGregorian(model.CloseDate)));
                parameters.Add(new SqlParameter("@IsClose", model.IsClose));
                parameters.Add(new SqlParameter("@add_num", model.tbInfo.add_num));

                result = _adoData.ExecuteNonQuery(T_SQL, parameters);
                return result;
            }
            catch (Exception)
            {

                throw;
            }
        }

        public byte[] Complaint_Excel(Complaint_M_req model)
        {
            var parameters = new List<SqlParameter>();
            try
            {
                var T_SQL = @"SELECT Ct.*,ub.item_D_name as CompSou_show,Um.U_name,Cs.item_D_name as Risk_show,Bc.item_D_name as ResponsBC_Show,
                              CASE WHEN IsClose='N' THEN '尚未結案' ELSE '結案' END as Close_show
                              FROM Complaint Ct 
                              LEFT JOIN User_M Um ON Ct.Sales_num = Um.u_num
                              LEFT JOIN Item_list Ub ON ub.item_M_code = 'CompSou'  AND ub.item_D_code = Ct.CompSou
                              LEFT JOIN Item_list Cs ON Cs.item_M_code = 'RiskLevel'  AND Cs.item_D_code = Ct.RiskLevel
                              LEFT JOIN Item_list Bc ON Bc.item_M_code = 'branch_company' AND Bc.item_D_code = Ct.ResponsBC
                              WHERE 1 = 1 ";

                if (model.IsClose != "A")
                {
                    T_SQL += @" AND IsClose = @IsClose";
                    parameters.Add(new SqlParameter("@IsClose", model.IsClose));
                }
                if (!string.IsNullOrWhiteSpace(model.CheckDateS) && !string.IsNullOrWhiteSpace(model.CheckDateE))
                {
                    T_SQL += @" AND CompDate BETWEEN @CheckDateS and @CheckDateE";
                    parameters.Add(new SqlParameter("@CheckDateS", FuncHandler.ConvertROCToGregorian(model.CheckDateS)));
                    parameters.Add(new SqlParameter("@CheckDateE", FuncHandler.ConvertROCToGregorian(model.CheckDateE)));
                }
                if (!string.IsNullOrWhiteSpace(model.CS_name))
                {
                    T_SQL += @" AND CS_name = @CS_name";
                    parameters.Add(new SqlParameter("@CS_name", model.CS_name));
                }
                if (!string.IsNullOrWhiteSpace(model.CS_PID))
                {
                    T_SQL += @" AND CS_PID = @CS_PID";
                    parameters.Add(new SqlParameter("@CS_PID", model.CS_PID));
                }
                if (!string.IsNullOrWhiteSpace(model.Sales_num))
                {
                    T_SQL += @" AND Sales_num = @Sales_num";
                    parameters.Add(new SqlParameter("@Sales_num", model.Sales_num));
                }
                if (!string.IsNullOrWhiteSpace(model.RiskLevel))
                {
                    T_SQL += @" AND RiskLevel = @RiskLevel";
                    parameters.Add(new SqlParameter("@RiskLevel", model.RiskLevel));
                }
                if (!string.IsNullOrWhiteSpace(model.CompSou))
                {
                    T_SQL += @" AND CompSou = @CompSou";
                    parameters.Add(new SqlParameter("@CompSou", model.CompSou));
                }
                if (model.RoleType == "RoleB")
                {
                    T_SQL += @" AND C.Sales_num = @UserNum";
                    parameters.Add(new SqlParameter("@UserNum", model.User));
                }
                if (model.RoleType == "RoleC")
                {
                    T_SQL += @" AND M1.U_BC = @UserBC";
                    parameters.Add(new SqlParameter("@UserBC", model.U_BC));
                }

                var excelList = _adoData.ExecuteQuery(T_SQL, parameters).AsEnumerable().Select(row => new Complaint_report_excel
                {
                    CompDate = row.Field<string>("CompDate"),
                    CompSou_show = row.Field<string>("CompSou_show"),
                    ContractID = row.Field<string>("ContractID"),
                    PassID = row.Field<string>("PassID"),
                    CS_name = _Fun.DeCodeBNWords(row.Field<string>("CS_name")),
                    CS_PID = row.Field<string>("CS_PID"),
                    U_name = row.Field<string>("U_name"),
                    ResponsBC_Show = row.Field<string>("ResponsBC_Show"),
                    Complaint = row.Field<string>("Complaint"),
                    Remark = row.Field<string>("Remark"),
                    Risk_show = row.Field<string>("Risk_show"),
                    ResponsWay = row.Field<string>("ResponsWay"),
                    ResponsStates = row.Field<string>("ResponsStates"),
                    Close_show = row.Field<string>("Close_show"),
                    CloseDate = row.Field<string>("CloseDate"),
                    DealDay = row.Field<int?>("DealDay")
                }).ToList();

                var Excel_Headers = new Dictionary<string, string>
                {
                    { "CompDate","客訴日期" },
                    { "CompSou_show", "客訴來源" },
                    { "ContractID", "合約編號" },
                    { "CS_name", "姓名" },
                    { "CS_PID", "ID" },
                    { "PassID", "通路商" },
                    { "U_name", "業務員" },
                    { "Complaint", "客訴原因" },
                    { "Remark", "客訴說明" },
                    { "Risk_show", "風險級別" },
                    { "ResponsBC_Show", "客訴處理權責單位" },
                    { "ResponsWay", "處理方式" },
                    { "ResponsStates", "處理狀況" },
                    { "Close_show", "是否完成" },
                    { "CloseDate", "結案日期" },
                    { "DealDay", "處理時效(天)" }
                };
                var fileBytes = FuncHandler.ExportToExcel(excelList, Excel_Headers);
                return fileBytes;
            }
            catch (Exception)
            {

                throw;
            }
        }

        public byte[] Complaint_Close_Excel(string DeadlineDate)
        {
            var parameters = new List<SqlParameter>();
            var DeadlineDateGre = FuncHandler.ConvertROCToGregorian(DeadlineDate);
            try
            {
                var T_SQL = @"SELECT Ct.*,ub.item_D_name as CompSou_show,Um.U_name,Cs.item_D_name as Risk_show,Bc.item_D_name as ResponsBC_Show,
                              CASE WHEN IsClose='N' THEN '尚未結案' ELSE '結案' END as Close_show
                              FROM Complaint Ct 
                              LEFT JOIN User_M Um ON Ct.Sales_num = Um.u_num
                              LEFT JOIN Item_list Ub ON ub.item_M_code = 'CompSou'  AND ub.item_D_code = Ct.CompSou
                              LEFT JOIN Item_list Cs ON Cs.item_M_code = 'RiskLevel'  AND Cs.item_D_code = Ct.RiskLevel
                              LEFT JOIN Item_list Bc ON Bc.item_M_code = 'branch_company' AND Bc.item_D_code = Ct.ResponsBC
                              WHERE 1 = 1 AND CompDate <= @Date";

                parameters.Add(new SqlParameter("@Date", DeadlineDateGre));

                var resultList = _adoData.ExecuteQuery(T_SQL, parameters).AsEnumerable().Select(row => new ComplaintCloseDto
                {
                    CompSou = row.Field<string>("CompSou"),
                    CompDate = row.Field<string>("CompDate"),
                    CloseDate = row.Field<string>("CloseDate"),
                    IsClose = row.Field<string>("IsClose"),
                    DealDay = row.Field<int?>("DealDay")
                }).ToList();


                DateTime monthEnd = DateTime.ParseExact(DeadlineDateGre, "yyyy/MM/dd", CultureInfo.InvariantCulture);
                DateTime monthStart = new DateTime(monthEnd.Year, monthEnd.Month, 1);

                //本月結案目前是採用結案日(CloseDate)判定
                var monthList = resultList.Where(x => DateTime.TryParseExact(x.CloseDate, "yyyy/MM/dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime closeDate)
                && closeDate >= monthStart && closeDate <= monthEnd && x.IsClose == "Y").ToList();

                //未結案
                var closeNoneList = resultList.Where(x=>x.IsClose == "N").ToList();

                #region EXCEL
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("結案檢核表");

                    #region 添加表頭
                    SetHeader(worksheet.Cells[1, 1, 2, 2], "案件來源");
                    SetHeader(worksheet.Cells[1, 3, 2, 3], "總件數(累計)");
                    SetHeader(worksheet.Cells[1, 4, 2, 4], "本月結案件數");
                    SetHeader(worksheet.Cells[1, 5, 2, 5], "處理時效(天)");
                    SetHeader(worksheet.Cells[1, 6, 2, 6], "本日結案");
                    SetHeader(worksheet.Cells[1, 7, 1, 9], "未結案");
                    SetHeader(worksheet.Cells[2, 7], "本日新增");
                    SetHeader(worksheet.Cells[2, 8], "處理中");
                    SetHeader(worksheet.Cells[2, 9], "小計");
                    #endregion

                    #region 添加表身
                    SetBody(worksheet.Cells[3, 1, 5, 1], "公家機關");

                    SetBody(worksheet.Cells[3, 2], "立法委員/議員");
                    SetBody(worksheet.Cells[3, 3], (resultList?.Count(x => x.CompSou == "1") ?? 0).ToString());
                    SetBody(worksheet.Cells[3, 4], (monthList?.Count(x => x.CompSou == "1") ?? 0).ToString());
                    SetBody(worksheet.Cells[3, 5], (monthList?.Where(x => x.CompSou == "1").Sum(x=>x.DealDay) ?? 0).ToString());
                    SetBody(worksheet.Cells[3, 6], (monthList?.Where(x => x.CompSou == "1" && x.CloseDate == DeadlineDateGre).Count() ?? 0).ToString());
                    SetBody(worksheet.Cells[3, 7], (closeNoneList?.Where(x => x.CompSou == "1" && x.CompDate == DeadlineDateGre).Count() ?? 0).ToString());
                    SetBody(worksheet.Cells[3, 8], (closeNoneList?.Where(x => x.CompSou == "1" && x.CompDate != DeadlineDateGre).Count() ?? 0).ToString());
                    SetBody(worksheet.Cells[3, 9], (closeNoneList?.Where(x => x.CompSou == "1").Count() ?? 0).ToString());

                    SetBody(worksheet.Cells[4, 2], "警察局/其他司法機關");
                    SetBody(worksheet.Cells[4, 3], (resultList?.Count(x => x.CompSou == "2") ?? 0).ToString());
                    SetBody(worksheet.Cells[4, 4], (monthList?.Count(x => x.CompSou == "2") ?? 0).ToString());
                    SetBody(worksheet.Cells[4, 5], (monthList?.Where(x => x.CompSou == "2").Sum(x => x.DealDay) ?? 0).ToString());
                    SetBody(worksheet.Cells[4, 6], (monthList?.Where(x => x.CompSou == "2" && x.CloseDate == DeadlineDateGre).Count() ?? 0).ToString());
                    SetBody(worksheet.Cells[4, 7], (closeNoneList?.Where(x => x.CompSou == "2" && x.CompDate == DeadlineDateGre).Count() ?? 0).ToString());
                    SetBody(worksheet.Cells[4, 8], (closeNoneList?.Where(x => x.CompSou == "2" && x.CompDate != DeadlineDateGre).Count() ?? 0).ToString());
                    SetBody(worksheet.Cells[4, 9], (closeNoneList?.Where(x => x.CompSou == "2").Count() ?? 0).ToString());

                    SetBody(worksheet.Cells[5, 2], "其他機關");
                    SetBody(worksheet.Cells[5, 3], (resultList?.Count(x => x.CompSou == "3") ?? 0).ToString());
                    SetBody(worksheet.Cells[5, 4], (monthList?.Count(x => x.CompSou == "3") ?? 0).ToString());
                    SetBody(worksheet.Cells[5, 5], (monthList?.Where(x => x.CompSou == "3").Sum(x => x.DealDay) ?? 0).ToString());
                    SetBody(worksheet.Cells[5, 6], (monthList?.Where(x => x.CompSou == "3" && x.CloseDate == DeadlineDateGre).Count() ?? 0).ToString());
                    SetBody(worksheet.Cells[5, 7], (closeNoneList?.Where(x => x.CompSou == "3" && x.CompDate == DeadlineDateGre).Count() ?? 0).ToString());
                    SetBody(worksheet.Cells[5, 8], (closeNoneList?.Where(x => x.CompSou == "3" && x.CompDate != DeadlineDateGre).Count() ?? 0).ToString());
                    SetBody(worksheet.Cells[5, 9], (closeNoneList?.Where(x => x.CompSou == "3").Count() ?? 0).ToString());

                    SetBody(worksheet.Cells[6, 1, 9, 1], "非公家機關");

                    SetBody(worksheet.Cells[6, 2], "電話/傳真");
                    SetBody(worksheet.Cells[6, 3], (resultList?.Count(x => x.CompSou == "4") ?? 0).ToString());
                    SetBody(worksheet.Cells[6, 4], (monthList?.Count(x => x.CompSou == "4") ?? 0).ToString());
                    SetBody(worksheet.Cells[6, 5], (monthList?.Where(x => x.CompSou == "4").Sum(x => x.DealDay) ?? 0).ToString());
                    SetBody(worksheet.Cells[6, 6], (monthList?.Where(x => x.CompSou == "4" && x.CloseDate == DeadlineDateGre).Count() ?? 0).ToString());
                    SetBody(worksheet.Cells[6, 7], (closeNoneList?.Where(x => x.CompSou == "4" && x.CompDate == DeadlineDateGre).Count() ?? 0).ToString());
                    SetBody(worksheet.Cells[6, 8], (closeNoneList?.Where(x => x.CompSou == "4" && x.CompDate != DeadlineDateGre).Count() ?? 0).ToString());
                    SetBody(worksheet.Cells[6, 9], (closeNoneList?.Where(x => x.CompSou == "4").Count() ?? 0).ToString());

                    SetBody(worksheet.Cells[7, 2], "律師函/存證信函");
                    SetBody(worksheet.Cells[7, 3], (resultList?.Count(x => x.CompSou == "5") ?? 0).ToString());
                    SetBody(worksheet.Cells[7, 4], (monthList?.Count(x => x.CompSou == "5") ?? 0).ToString());
                    SetBody(worksheet.Cells[7, 5], (monthList?.Where(x => x.CompSou == "5").Sum(x => x.DealDay) ?? 0).ToString());
                    SetBody(worksheet.Cells[7, 6], (monthList?.Where(x => x.CompSou == "5" && x.CloseDate == DeadlineDateGre).Count() ?? 0).ToString());
                    SetBody(worksheet.Cells[7, 7], (closeNoneList?.Where(x => x.CompSou == "5" && x.CompDate == DeadlineDateGre).Count() ?? 0).ToString());
                    SetBody(worksheet.Cells[7, 8], (closeNoneList?.Where(x => x.CompSou == "5" && x.CompDate != DeadlineDateGre).Count() ?? 0).ToString());
                    SetBody(worksheet.Cells[7, 9], (closeNoneList?.Where(x => x.CompSou == "5").Count() ?? 0).ToString());

                    SetBody(worksheet.Cells[8, 2], "客服信箱");
                    SetBody(worksheet.Cells[8, 3], (resultList?.Count(x => x.CompSou == "6") ?? 0).ToString());
                    SetBody(worksheet.Cells[8, 4], (monthList?.Count(x => x.CompSou == "6") ?? 0).ToString());
                    SetBody(worksheet.Cells[8, 5], (monthList?.Where(x => x.CompSou == "6").Sum(x => x.DealDay) ?? 0).ToString());
                    SetBody(worksheet.Cells[8, 6], (monthList?.Where(x => x.CompSou == "6" && x.CloseDate == DeadlineDateGre).Count() ?? 0).ToString());
                    SetBody(worksheet.Cells[8, 7], (closeNoneList?.Where(x => x.CompSou == "6" && x.CompDate == DeadlineDateGre).Count() ?? 0).ToString());
                    SetBody(worksheet.Cells[8, 8], (closeNoneList?.Where(x => x.CompSou == "6" && x.CompDate != DeadlineDateGre).Count() ?? 0).ToString());
                    SetBody(worksheet.Cells[8, 9], (closeNoneList?.Where(x => x.CompSou == "6").Count() ?? 0).ToString());

                    SetBody(worksheet.Cells[9, 2], "其他方式");
                    SetBody(worksheet.Cells[9, 3], (resultList?.Count(x => x.CompSou == "7") ?? 0).ToString());
                    SetBody(worksheet.Cells[9, 4], (monthList?.Count(x => x.CompSou == "7") ?? 0).ToString());
                    SetBody(worksheet.Cells[9, 5], (monthList?.Where(x => x.CompSou == "7").Sum(x => x.DealDay) ?? 0).ToString());
                    SetBody(worksheet.Cells[9, 6], (monthList?.Where(x => x.CompSou == "7" && x.CloseDate == DeadlineDateGre).Count() ?? 0).ToString());
                    SetBody(worksheet.Cells[9, 7], (closeNoneList?.Where(x => x.CompSou == "7" && x.CompDate == DeadlineDateGre).Count() ?? 0).ToString());
                    SetBody(worksheet.Cells[9, 8], (closeNoneList?.Where(x => x.CompSou == "7" && x.CompDate != DeadlineDateGre).Count() ?? 0).ToString());
                    SetBody(worksheet.Cells[9, 9], (closeNoneList?.Where(x => x.CompSou == "7").Count() ?? 0).ToString());
                    #endregion

                    #region 總計
                    SetBody(worksheet.Cells[10, 1], "總計");
                    SetBody(worksheet.Cells[10, 2], "");
                    SetBody(worksheet.Cells[10, 3], (resultList?.Count() ?? 0).ToString());
                    SetBody(worksheet.Cells[10, 4], (monthList?.Count() ?? 0).ToString());
                    SetBody(worksheet.Cells[10, 5], (monthList?.Sum(x => x.DealDay) ?? 0).ToString());
                    SetBody(worksheet.Cells[10, 6], (monthList?.Where(x=> x.CloseDate == DeadlineDateGre).Count() ?? 0).ToString());
                    SetBody(worksheet.Cells[10, 7], (closeNoneList?.Where(x => x.CompDate == DeadlineDateGre).Count() ?? 0).ToString());
                    SetBody(worksheet.Cells[10, 8], (closeNoneList?.Where(x => x.CompDate != DeadlineDateGre).Count() ?? 0).ToString());
                    SetBody(worksheet.Cells[10, 9], (closeNoneList?.Count() ?? 0).ToString());

                    var range = worksheet.Cells[10, 1, 10, 9];
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.LightYellow);
                    #endregion

                    //調整長度
                    worksheet.Column(1).Width = 13;
                    worksheet.Column(2).Width = 20;
                    worksheet.Column(3).Width = 15;
                    worksheet.Column(4).Width = 15;
                    worksheet.Column(5).Width = 15;
                    worksheet.Column(6).Width = 15;

                    return package.GetAsByteArray();
                }
                #endregion
            }
            catch (Exception)
            {

                throw;
            }

        }
        #endregion


        void SetHeader(ExcelRange range, string text)
        {
            range.Merge = true;
            range.Value = text;

            var style = range.Style;

            // 背景
            style.Fill.PatternType = ExcelFillStyle.Solid;
            style.Fill.BackgroundColor.SetColor(Color.LightBlue);

            // 文字
            style.WrapText = true;
            style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            style.VerticalAlignment = ExcelVerticalAlignment.Center;

            // 邊框（四邊）
            style.Border.Top.Style = ExcelBorderStyle.Thin;
            style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            style.Border.Left.Style = ExcelBorderStyle.Thin;
            style.Border.Right.Style = ExcelBorderStyle.Thin;
        }

        void SetBody(ExcelRange range, string text)
        {
            range.Merge = true;
            range.Value = text;

            var style = range.Style;

            // 文字
            style.WrapText = true;
            style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            style.VerticalAlignment = ExcelVerticalAlignment.Center;

            // 邊框（四邊）
            style.Border.Top.Style = ExcelBorderStyle.Thin;
            style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            style.Border.Left.Style = ExcelBorderStyle.Thin;
            style.Border.Right.Style = ExcelBorderStyle.Thin;
        }
    }
}
