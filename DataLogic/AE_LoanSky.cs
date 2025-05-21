using Azure;
using KF_WebAPI.BaseClass;
using KF_WebAPI.BaseClass.AE;
using KF_WebAPI.BaseClass.LoanSky;
using KF_WebAPI.FunctionHandler;
using LoanSky.Model;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using System.Collections.Generic;
using System.Data;
using System.ServiceModel.Security;
using System.Text;
using System.Text.Json;

namespace KF_WebAPI.DataLogic
{
    public class AE_LoanSky
    {
        private string account = "testCGU"; // 測試用帳號:testCGU
        private readonly string _storagePath = @"C:\UploadedFiles";
        ADOData _ADO = new();

        /// <summary>
        /// 取得AE要拋轉LoanSky案件資料
        /// </summary>
        /// <param name="req"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public runOrderRealEstateRequest GetOrderRealEstateRequest(LoanSky_Req req)
        {
            runOrderRealEstateRequest runReq = new runOrderRealEstateRequest();

            try
            {
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"
                     select *
                         ,(select CS_PID from House_apply where HA_cknum = @HA_cknum) as CS_PID
                         ,(select  case when U_num in (SELECT item_D_code FROM dbo.Item_list where item_M_code ='newloan' and item_M_type='N') then 'BC0800' else U_BC  end U_BC FROM User_M where U_num = House_pre.add_num AND del_tag='0') as U_BC
                         ,(select U_name FROM User_M where U_num = House_pre.add_num AND del_tag='0') as add_name
                         ,(select U_num  FROM User_M where U_num = House_pre.add_num AND del_tag='0') as user_num
                         ,(select item_D_name from Item_list where item_M_code = 'building_kind' AND item_D_type='Y' AND item_D_code = House_pre.pre_building_kind  AND show_tag='0' AND del_tag='0') as show_pre_building_kind
                         ,(select item_D_name from Item_list where item_M_code = 'parking_kind' AND item_D_type='Y' AND item_D_code = House_pre.pre_parking_kind  AND show_tag='0' AND del_tag='0') as show_pre_parking_kind
                         ,(select item_D_name from Item_list where item_M_code = 'building_material' AND item_D_type='Y' AND item_D_code = House_pre.pre_building_material  AND show_tag='0' AND del_tag='0') as show_pre_building_material
                         ,isnull((select U_name FROM User_M where U_num = House_pre.pre_process_num AND del_tag='0'),'') as pre_process_name
                         from House_pre 
                         where del_tag = '0' AND HP_id = @HP_id  AND HP_cknum = @HP_cknum
                    ";
                parameters.Add(new SqlParameter("@HA_cknum", req.HA_cknum));
                parameters.Add(new SqlParameter("@HP_id", req.HP_id));
                parameters.Add(new SqlParameter("@HP_cknum", req.HP_cknum));

                #endregion
                DataTable dtResult = _ADO.ExecuteQuery(T_SQL, parameters);
                if (dtResult.Rows.Count > 0)
                {
                    // 來源資料
                    runReq.housePre_res = dtResult.AsEnumerable().Select(row => new HousePre_res
                    {
                        HP_id = row.IsNull("HP_id") ? 0 : row.Field<decimal>("HP_id"), // 房屋預估資料ID
                        HP_cknum = row.IsNull("HP_cknum") ? "" : row.Field<string>("HP_cknum"), // 房屋預估資料流水號
                        add_name = row.IsNull("add_name") ? "" : row.Field<string>("add_name"),  //經辦人名稱
                        pre_apply_name = row.IsNull("pre_apply_name") ? "" : row.Field<string>("pre_apply_name"),    // 申請人
                        show_pre_building_kind = row.IsNull("show_pre_building_kind") ? "" : row.Field<string>("show_pre_building_kind"),  // 建物類型(請參照對照表)
                        pre_city = row.IsNull("pre_city") ? "" : row.Field<string>("pre_city"),// 縣市代碼(請參照對照表)
                        pre_area = row.IsNull("pre_area") ? "" : row.Field<string>("pre_area"),// 鄉鎮市區代碼(請參照對照表)
                        pre_road = row.IsNull("pre_road") ? "" : row.Field<string>("pre_road"),// 段(請參照對照表)
                        pre_road_s = row.IsNull("pre_road_s") ? "" : row.Field<string>("pre_road_s")// 小段(請參照對照表)
                    }).FirstOrDefault();


                    #region 取得段代碼(狀態:current)：縣市代碼+區代碼+段名稱
                    SectionCode reqCode = dtResult.AsEnumerable().Select(row => new SectionCode
                    {
                        city_name = row.IsNull("pre_city") ? string.Empty : row.Field<string>("pre_city"),
                        area_name = row.IsNull("pre_area") ? string.Empty : row.Field<string>("pre_area"),
                        road_name = row.IsNull("pre_road") ? string.Empty : row.Field<string>("pre_road"),
                        road_name_s = (row.IsNull("pre_road_s") ? string.Empty : row.Field<string>("pre_road_s"))
                    }).SingleOrDefault();

                    if (reqCode.isRight() == false)
                    {
                        runReq.message = reqCode.message;
                        runReq.isNeedPopupWindow = true;
                        return runReq;
                    }
                    reqCode = GetMoiSectionCode(reqCode); // 回傳值nulll:段代碼不存在

                    // 判斷段代碼(狀態:before) 與 段代碼(狀態:current) 不同時，預設為:段代碼(狀態:current)
                    // 段代碼為空值時
                    if ((reqCode == null) || (reqCode?.road_num == null))
                    {
                        runReq.message = "找不到段代碼，檢查對照表";
                        runReq.isNeedPopupWindow = true;
                        return runReq;
                    }
                    #endregion
                    #region 案件資料
                    runReq.oreRequest = dtResult.AsEnumerable().Select(row => new OrderRealEstateRequest
                    {
                        Account = account, // 承辦人員帳號:測試用帳號:testCGU
                        BusinessUserName = row.IsNull("add_name") ? "" : row.Field<string>("add_name"),  //經辦人名稱
                        //BusinessTel = "82096100",  // 經辦人電話
                        //BusinessFax = "123",   // 經辦人傳真
                        Applicant = row.IsNull("pre_apply_name") ? "" : row.Field<string>("pre_apply_name"),    // 申請人
                        IdNo = row.IsNull("CS_PID") ? "" : row.Field<string>("CS_PID"),         // 公司統編/身分證號
                        Condominium = row.IsNull("pre_community_name") ? "" : row.Field<string>("pre_community_name"),   // 社區(大樓)名稱
                        BuildingState = row.IsNull("show_pre_building_kind") ? "" : AE2BuildingState(row.Field<string>("show_pre_building_kind")),  // 建物類型(請參照對照表)
                        Situation = row.IsNull("pre_use_kind") ? "" : row.Field<string>("pre_use_kind"),   //使用現況
                        ParkCategory = (row.IsNull("show_pre_parking_kind") || string.IsNullOrEmpty(row.Field<string>("show_pre_parking_kind")) ) ? "" : AE2ParkCategory(row.Field<string>("show_pre_parking_kind")),  // 車位型態(請參照對照表)
                        Note =  ((row.IsNull("pre_note") || string.IsNullOrEmpty(row.Field<string>("pre_note")))?"": $"[建物類型備註]{row.Field<string>("pre_note")}{Environment.NewLine}") +  
                                ((row.IsNull("pre_building_material_note") || string.IsNullOrEmpty(row.Field<string>("pre_building_material_note"))) ?"":$"[主要建築材料備註]{row.Field<string>("pre_building_material_note")}{Environment.NewLine}") +
                                ((row.IsNull("pre_principal_note") || string.IsNullOrEmpty(row.Field<string>("pre_principal_note"))) ? "" : $"[委託人備註]{row.Field<string>("pre_principal_note")}{Environment.NewLine}") +
                                ((row.IsNull("pre_other_note") || string.IsNullOrEmpty(row.Field<string>("pre_other_note"))) ? "" : $"[其他備註]{row.Field<string>("pre_other_note")}"), // 備註
                        MoiCityCode = reqCode.city_num, // 縣市代碼(請參照對照表)
                        MoiTownCode = reqCode.area_num // 鄉鎮市區代碼(請參照對照表)s
                    }).FirstOrDefault();

                    var no = dtResult.AsEnumerable().Select(row => new OrderRealEstateNoRequest
                    {
                        MoiSectionCode = reqCode.road_num,  // 段代碼:查詢條件：縣市代碼+區代碼+段名稱
                        BuildNos = row.IsNull("pre_build_num") ? "" : row.Field<string>("pre_build_num").Replace('－', '-').Replace('、', ',').Replace('；', ','),   // 建號 pre_build_num 多筆用逗號分隔
                        LandNos = row.IsNull("pre_land_num") ? "" : row.Field<string>("pre_land_num").Replace('－', '-').Replace('、', ',').Replace('；', ',') // 地號 pre_land_num 多筆用逗號分隔
                    }).SingleOrDefault();
                    runReq.oreRequest.Nos.Add(no);
                    #endregion
                    #region 取得AE要拋轉LoanSky案件資料-pdf附件
                    List<OrderRealEstateAttachmentRequest> attachments = GetOrderRealEstateNoRequest(req.HA_cknum);
                    if (attachments != null)
                    {
                        foreach (var item in attachments)
                        {
                            runReq.oreRequest.Attachments.Add(item);
                        }
                    }
                    #endregion
                }
            }
            catch (Exception ex)
            {
                throw new Exception("AE_LoanSky.GetOrderRealEstateRequest", ex);
            }
            return runReq;
        }
        /// <summary>
        /// 取得AE要拋轉LoanSky案件資料-附件
        /// </summary>
        /// <param name="HA_cknum"></param>
        /// <returns></returns>
        public List<OrderRealEstateAttachmentRequest> GetOrderRealEstateNoRequest(string HA_cknum)
        {
            List<OrderRealEstateAttachmentRequest> ReqClass = new List<OrderRealEstateAttachmentRequest>();
            var parameters = new List<SqlParameter>();
            var t_sqlAttachment = @"
                         select * 
                            from ASP_UpLoad
                            where del_tag = '0' AND cknum = @HA_cknum
                                and upload_name_code like '%.pdf'
                                order by add_date desc
                    ";
            parameters.Add(new SqlParameter("@HA_cknum", HA_cknum));
            DataTable dtResult = _ADO.ExecuteQuery(t_sqlAttachment, parameters);
            if (dtResult.Rows.Count > 0)
            {
                foreach (DataRow row in dtResult.Rows)
                {
                    OrderRealEstateAttachmentRequest attachmentRequest = new OrderRealEstateAttachmentRequest();
                    attachmentRequest.OrginalFileName = row.IsNull("upload_name_show") ? "" : row.Field<string>("upload_name_show"); // 原始檔案名稱(需包含副檔名且是PDF檔) 

                    var upload_name_code = (row["upload_name_code"]).ToString();
                    var upload_name_show = (row["upload_name_show"]).ToString();

                    string _filePath = Path.Combine(_storagePath, upload_name_code.Substring(0, 6), upload_name_code.Substring(0, 8), upload_name_code);
                    if (!System.IO.File.Exists(_filePath))
                    {
                        //throw new Exception("AE_LoanSky.GetOrderRealEstateNoRequest:file is not fund");
                        continue; // 檔案不存在時返回 null
                    }
                    attachmentRequest.File = System.IO.File.ReadAllBytes(_filePath); // 附件檔案內容

                    ReqClass.Add(attachmentRequest);
                }
            }
            return ReqClass;
        }
        public string GetJsonFileContent(string filePath)
        {
            string jsonData = string.Empty;

            try
            {
                if (!System.IO.File.Exists(filePath))
                {
                    throw new FileNotFoundException("SectionCode.json file not found.");
                }
                jsonData = System.IO.File.ReadAllText(filePath);
            }
            catch (Exception ex)
            {
                throw new Exception("AE_LoanSky.GetJsonFileContent", ex);
            }
            return jsonData;
        }
        public List<CityCode> AE2MoiCityCode(string pre_city)
        {
            string filePath = Path.Combine(Directory.GetCurrentDirectory(), "BaseClass", "LoanSky", "CityCode.json");
            string jsonData = GetJsonFileContent(filePath);
            var jsonObject = JsonSerializer.Deserialize<List<CityCode>>(jsonData);

            var resultcode = jsonObject.Where(x => string.IsNullOrEmpty(pre_city) || x.city_name.Contains(pre_city.Replace("台", "臺"))).ToList();
            return resultcode;
        }
        public List<AreaCode> AE2MoiTownCode(string pre_city, string pre_area)
        {
            string filePath = Path.Combine(Directory.GetCurrentDirectory(), "BaseClass", "LoanSky", "AreaCode.json");
            string jsonData = GetJsonFileContent(filePath);
            var jsonObject = JsonSerializer.Deserialize<List<AreaCode>>(jsonData);

            var resultcode = jsonObject.Where(x => x.city_num.Equals(pre_city) && (string.IsNullOrEmpty(pre_area) || x.area_name.Contains(pre_area))).ToList();
            return resultcode;
        }
        public List<SectionCode> GetMoiSectionCode(string city_num, string area_num, string road_name)
        {
            FuncHandler _fun = new FuncHandler();
            string filePath = Path.Combine(Directory.GetCurrentDirectory(), "BaseClass", "LoanSky", "SectionCode.json");
            string jsonData = GetJsonFileContent(filePath);
            var jsonObject = JsonSerializer.Deserialize<List<SectionCode>>(jsonData);

            var resultcode = jsonObject.Where(x => x.city_num.Equals(city_num)
                && x.area_num.Equals(area_num.TrimStart(new char[] { '0' }))
                && (string.IsNullOrEmpty(road_name) || x.road_name.Contains(road_name))).ToList();
            return resultcode;
        }
        /// <summary>
        /// KF2LoanSky:縣市代碼+區代碼+段名稱+段代碼
        /// </summary>
        /// <param name="sectionCode"></param>
        /// <returns></returns>
        public SectionCode GetMoiSectionCode(SectionCode sectionCode)
        {
            #region 驗證傳入參數:縣市代碼+鄉鎮市區代碼+段代碼必需要有值才能判斷
            if (sectionCode.isRight() == false)
            {
                return null;
            }
            #endregion
            string filePath = Path.Combine(Directory.GetCurrentDirectory(), "BaseClass", "LoanSky", "SectionCode.json");
            string jsonData = GetJsonFileContent(filePath);
            var jsonObject = JsonSerializer.Deserialize<List<SectionCode>>(jsonData);


            bool isHaveEndWord = sectionCode.road_name.Substring(sectionCode.road_name.Length - 1).Equals("段");
            if (!isHaveEndWord)
            {
                sectionCode.road_name = sectionCode.road_name + "段";
            }
            string roadNameS = sectionCode.road_name_s?.Replace("0", ""); // 0代表沒有“小段”
            sectionCode.road_name = sectionCode.road_name + roadNameS;
            var resultcode = jsonObject.Where(x =>
                        x.city_name.Contains(sectionCode.city_name.Replace("台", "臺"))
                    && x.area_name.Contains(sectionCode.area_name)
                    && x.road_name.Contains(sectionCode.road_name))
                .Select(x => new SectionCode
                {
                    city_num = x.city_num,
                    city_name = x.city_name,
                    area_num = x.area_num.PadLeft(2, '0'),
                    area_name = x.area_name,
                    road_num = x.road_num.PadLeft(4, '0'),
                    road_name = x.road_name
                });


            return (resultcode == null || resultcode.Count() > 1) ? null : resultcode.SingleOrDefault();
        }
        
        /// <summary>
        /// KF2LoanSky:建物類型
        /// </summary>
        /// <param name="sectionCode"></param>
        /// <returns></returns>
        public string AE2BuildingState(string show_pre_building_kind)
        {

            List<(string, string)> lsBuildingState = new List<(string, string)>{
                ("公寓", "公寓(5樓含以下無電梯)"),
                ("華廈", "華廈(10層含以下有電梯)"),
                ("電梯大樓", "住宅大樓(11層含以上有電梯)"),
                ("集合住宅", "住宅大樓(11層含以上有電梯)"),
                ("透天", "透天厝"),
                ("假透天", "透天厝"),
                ("商業辦公大樓", "辦公商業大樓"),
                ("店面", "店面(店鋪)"),
                ("套房", "套房(1房1廳1衛)"),
                ("工廠", "工廠"),
                ("套房", "套房(1房1廳1衛)"),
                ("廠辦", "廠辦"),
                ("農舍", "農舍"),
                ("倉庫", "倉庫")
                //("其他", "");
                //("", "土地")
            };

            string BuildingState = lsBuildingState.Where(i => i.Item1.Equals(show_pre_building_kind)).Select(i => i.Item2).FirstOrDefault();

            return BuildingState;
        }
        /// <summary>
        /// KF2LoanSky:車位型態
        /// </summary>
        /// <param name="show_pre_building_kind"></param>
        /// <returns></returns>
        public string AE2ParkCategory(string show_pre_parking_kind)
        {

            List<(string, string)> lsParkCategory = new List<(string, string)>{
                ("坡道平面", "坡道平面"),
                ("坡道機械", "坡道機械"),
                ("機械平面", "機械機械")
                //("", "一樓平面"),
                //("", "塔式車位"),
                //("", "其他"),
            };

            string ParkCategory = lsParkCategory.Where(i => i.Item1.Equals(show_pre_parking_kind)).Select(i => i.Item2).FirstOrDefault();

            return ParkCategory;
        }
        public void InsertExternal_API_Log(External_API_Log p_External_API_Log)
        {

            try
            {
                string TableName = "External_API_Log";
                External_API_Log[] arrExternal_API_Log = new External_API_Log[1];
                arrExternal_API_Log[0] = p_External_API_Log;

                _ADO.DataTableToSQL(TableName, arrExternal_API_Log, _ADO.ConnStr);
            }
            catch
            {
                throw;
            }

        }
        public runOrderRealEstateRequest IsLoanSkyFieldsNull(LoanSky_Req req)
        {
            runOrderRealEstateRequest runReq = new runOrderRealEstateRequest();

            try
            {
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"
                     select
                         HP_id, HP_cknum
                         ,[pre_city],[pre_area],[pre_road],[pre_road_s],pre_apply_name
                         ,(select U_name FROM User_M where U_num = House_pre.add_num AND del_tag='0') as add_name
                         ,(select item_D_name from Item_list where item_M_code = 'building_kind' AND item_D_type='Y' AND item_D_code = House_pre.pre_building_kind  AND show_tag='0' AND del_tag='0') as show_pre_building_kind
                         ,(select item_D_name from Item_list where item_M_code = 'parking_kind' AND item_D_type='Y' AND item_D_code = House_pre.pre_parking_kind  AND show_tag='0' AND del_tag='0') as show_pre_parking_kind
                         ,[BuildingState],[ParkCategory],[MoiCityCode],[MoiTownCode],[MoiSectionCode]
                     from House_pre 
                     where del_tag = '0' AND HP_id = @HP_id  AND HP_cknum = @HP_cknum
                    ";
                parameters.Add(new SqlParameter("@HP_id", req.HP_id));
                parameters.Add(new SqlParameter("@HP_cknum", req.HP_cknum));

                #endregion
                DataTable dtResult = _ADO.ExecuteQuery(T_SQL, parameters);
                if (dtResult.Rows.Count > 0)
                {
                    // 來源資料
                    runReq.housePre_res = dtResult.AsEnumerable().Select(row => new HousePre_res
                    {
                        HP_id = row.IsNull("HP_id") ? 0 : row.Field<decimal>("HP_id"), // 房屋預估資料ID
                        HP_cknum = row.IsNull("HP_cknum") ? "" : row.Field<string>("HP_cknum"), // 房屋預估資料流水號
                        add_name = row.IsNull("add_name") ? "" : row.Field<string>("add_name"),  //經辦人名稱
                        pre_apply_name = row.IsNull("pre_apply_name") ? "" : row.Field<string>("pre_apply_name"),    // 申請人
                        show_pre_building_kind = row.IsNull("show_pre_building_kind") ? "" : row.Field<string>("show_pre_building_kind"),  // 建物類型(請參照對照表)
                        pre_city = row.IsNull("pre_city") ? "" : row.Field<string>("pre_city"),// 縣市代碼(請參照對照表)
                        pre_area = row.IsNull("pre_area") ? "" : row.Field<string>("pre_area"),// 鄉鎮市區代碼(請參照對照表)
                        pre_road = row.IsNull("pre_road") ? "" : row.Field<string>("pre_road"),// 段(請參照對照表)
                        pre_road_s = row.IsNull("pre_road_s") ? "" : row.Field<string>("pre_road_s")// 小段(請參照對照表)
                    }).FirstOrDefault();
                    if (string.IsNullOrEmpty(runReq.housePre_res.show_pre_building_kind))
                    {
                        runReq.message = "建物類型為必填欄位";
                        runReq.isNeedPopupWindow = true;
                        return runReq;
                    }

                    if (string.IsNullOrEmpty(runReq.housePre_res.show_pre_building_kind))
                    {
                        runReq.message = "建物類型為其他，請選擇合適";
                        runReq.isNeedPopupWindow = true;
                        return runReq;
                    }

                    // 取得段代碼(狀態:current)：縣市代碼+區代碼+段名稱
                    SectionCode reqCode = dtResult.AsEnumerable().Select(row => new SectionCode
                    {
                        city_name = row.IsNull("pre_city") ? string.Empty : row.Field<string>("pre_city"),
                        area_name = row.IsNull("pre_area") ? string.Empty : row.Field<string>("pre_area"),
                        road_name = row.IsNull("pre_road") ? string.Empty : row.Field<string>("pre_road"),
                        road_name_s = (row.IsNull("pre_road_s") ? string.Empty : row.Field<string>("pre_road_s"))
                    }).SingleOrDefault();
                    // 判斷縣市代碼+區代碼+段名稱是否有缺
                    if (reqCode.isRight() == false)
                    {
                        runReq.message = reqCode.message;
                        runReq.isNeedPopupWindow = true;
                        return runReq;
                    }
                    reqCode = GetMoiSectionCode(reqCode);
                    // 判斷段代碼(狀態:before) 與 段代碼(狀態:current) 不同時，預設為:段代碼(狀態:current)
                    // 段代碼為空值時
                    if ((reqCode == null) || (reqCode?.road_num == null))
                    {
                        runReq.message = "找不到段代碼，檢查對照表";
                        runReq.isNeedPopupWindow = true;
                        return runReq;
                    }

                    // 串接LoanSky
                    runReq.oreRequest = dtResult.AsEnumerable().Select(row => new OrderRealEstateRequest
                    {
                        BusinessUserName = row.IsNull("add_name") ? "" : row.Field<string>("add_name"),  //經辦人名稱
                        Applicant = row.IsNull("pre_apply_name") ? "" : row.Field<string>("pre_apply_name"),    // 申請人
                        BuildingState = row.IsNull("BuildingState") ? AE2BuildingState(row.Field<string>("show_pre_building_kind")) : row.Field<string>("BuildingState"),  // 建物類型(請參照對照表)
                        ParkCategory = row.IsNull("ParkCategory") ? AE2ParkCategory(row.Field<string>("show_pre_parking_kind")) : row.Field<string>("ParkCategory"),  // 車位型態(請參照對照表)
                        MoiCityCode = row.IsNull("MoiCityCode") ? reqCode.city_num : row.Field<string>("MoiCityCode"), // 縣市代碼(請參照對照表)
                        MoiTownCode = row.IsNull("MoiTownCode") ? reqCode.area_num : row.Field<string>("MoiTownCode") // 鄉鎮市區代碼(請參照對照表)
                    }).FirstOrDefault();
                    var no = dtResult.AsEnumerable().Select(row => new OrderRealEstateNoRequest
                    {
                        MoiSectionCode = row.IsNull("MoiSectionCode") ? reqCode.road_num : row.Field<string>("MoiSectionCode")  // 段代碼
                    }).SingleOrDefault();
                    runReq.oreRequest.Nos.Add(no);
                    runReq.isNeedPopupWindow = runReq.oreRequest.IsLoanSkyFieldsNull;
                    if (runReq.isNeedPopupWindow)
                    {
                        runReq.message = "LoanSky欄位為空值";
                        return runReq;
                    }
                    return runReq;
                }
            }
            catch (Exception ex)
            {
                throw new Exception("AE_LoanSky.GetOrderRealEstateRequest", ex);
            }
            return runReq;
        }
        public ResultClass<string> House_pre_Update(runOrderRealEstateRequest model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL

                var T_SQL = @"
                        update  House_pre
                          set BuildingState = @BuildingState,
                            ParkCategory = @ParkCategory,
                            MoiCityCode = @MoiCityCode,
                            MoiTownCode = @MoiTownCode,
                            MoiSectionCode = @MoiSectionCode
                         FROM House_pre
                         where HP_cknum = @hp_cknum
                ";
                var parameters = new List<SqlParameter>
                  {
                      new SqlParameter("@BuildingState", string.IsNullOrEmpty(model.oreRequest.BuildingState)? DBNull.Value :model.oreRequest.BuildingState ),
                      new SqlParameter("@ParkCategory", string.IsNullOrEmpty(model.oreRequest.ParkCategory)? DBNull.Value : model.oreRequest.ParkCategory),
                      new SqlParameter("@MoiCityCode", string.IsNullOrEmpty(model.oreRequest.MoiCityCode) ? DBNull.Value : model.oreRequest.MoiCityCode),
                      new SqlParameter("@MoiTownCode", string.IsNullOrEmpty(model.oreRequest.MoiTownCode) ? DBNull.Value : model.oreRequest.MoiTownCode),
                      new SqlParameter("@MoiSectionCode", string.IsNullOrEmpty(model.oreRequest.Nos.FirstOrDefault()?.MoiSectionCode)? DBNull.Value : model.oreRequest.Nos.FirstOrDefault().MoiSectionCode),
                      new SqlParameter("@hp_cknum", model.housePre_res.HP_cknum)
                  };
                #endregion
                int result = _adoData.ExecuteNonQuery(T_SQL, parameters);

                if (result == 0)
                {
                    resultClass.ResultCode = "400";
                    resultClass.ResultMsg = "更新失敗";
                }
                else
                {
                    resultClass.ResultCode = "000";
                    resultClass.ResultMsg = "更新成功";
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



    public class runOrderRealEstateRequest
    {
        public OrderRealEstateRequest oreRequest;   // LoanSky串接資料
        public HousePre_res housePre_res;           // 來源資料
        public string message { get; set; }
        /// <summary>
        /// 是否需要彈跳視窗讓使用者確認LoanSky參數
        /// </summary>
        public bool isNeedPopupWindow { get; set; }
    }
}
