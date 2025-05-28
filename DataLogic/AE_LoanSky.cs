using KF_WebAPI.BaseClass;
using KF_WebAPI.BaseClass.AE;
using KF_WebAPI.BaseClass.LoanSky;
using KF_WebAPI.FunctionHandler;
using LoanSky.Model;
using Microsoft.Data.SqlClient;
using System.Data;
using System.Text;
using System.Text.Json;

namespace KF_WebAPI.DataLogic
{
    public class AE_LoanSky
    {
        private string account = "testCGU"; // [測試區]使用測試用帳號:testCGU;當空值時，為[正式區]用u_num對應分公司的帳號
        
        private readonly string _storagePath = @"C:\UploadedFiles";
        ADOData _ADO = new();

        /// <summary>
        /// 取得LoanSky帳號
        /// </summary>
        /// <param name="U_num"></param>
        /// <returns></returns>
        public LoanSkyAccount GetLoanSkyAccountByUser(string U_num)
        {             // 取得LoanSky帳號
            User_M user = GetUser(U_num: U_num);
            LoanSkyAccount result = GetAllLoanSkyAccount(user.U_BC);
            return  result;
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
        /// 取得各分公司LoanSky Account
        /// </summary>
        /// <param name="pre_city"></param>
        /// <param name="pre_area"></param>
        /// <returns></returns>
        public LoanSkyAccount GetAllLoanSkyAccount(string item_D_code)
        {
            string filePath = Path.Combine(Directory.GetCurrentDirectory(), "BaseClass", "LoanSky", "LoanSkyAccount.json");
            string jsonData = GetJsonFileContent(filePath);
            var jsonObject = JsonSerializer.Deserialize<List<LoanSkyAccount>>(jsonData);

            var resultcode = jsonObject.Where(x => x.item_D_code.Equals(item_D_code)).SingleOrDefault();
            return resultcode;
        }

        public User_M GetUser(string? U_num="", string? U_name="")
        {
            string U_BC = string.Empty;
            User_M user = new User_M();
            try
            {
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"
                    select *  
                       FROM User_M 
                       where 
                           (@U_num ='' or U_num=@U_num)
                           AND (@U_name ='' or U_name=@U_name)
                    ";
                parameters.Add(new SqlParameter("@U_num", U_num));
                parameters.Add(new SqlParameter("@U_name", U_name));
                #endregion
                DataTable dtResult = _ADO.ExecuteQuery(T_SQL, parameters);
                
                if (dtResult.Rows.Count > 0)
                {
                    user = dtResult.AsEnumerable().Select(row => new User_M {
                        U_num = row.IsNull("U_num") ? "" : row.Field<string>("U_num"), // 
                        U_BC = row.IsNull("U_BC") ? "" : row.Field<string>("U_BC"),     // 分公司
                        U_name = row.IsNull("U_name") ? "" : row.Field<string>("U_name") // 分公司
                    }).SingleOrDefault();
                    
                }
            }
            catch (Exception ex)
            {
                throw new Exception("AE_LoanSky.GetU_BC", ex);
            }
            return user;
        }

        public List<(string, string)> GetPre_building_kind()
        {
            List<(string, string)> lsPBK = new List<(string, string)> ();
            try
            {
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"
                     select item_D_code, item_D_name 
                        from Item_list 
                        where item_M_code = 'building_kind' AND item_D_type='Y' 
                            AND show_tag='0' AND del_tag='0'
                    ";
                #endregion
                DataTable dtResult = _ADO.ExecuteQuery(T_SQL, parameters);
                if (dtResult.Rows.Count > 0)
                {
                    // 來源資料
                    lsPBK = dtResult.AsEnumerable().Select(row => 
                    (
                        row.IsNull("item_D_code") ? "" : row.Field<string>("item_D_code"), // 房屋預估資料ID
                        row.IsNull("item_D_name") ? "" : row.Field<string>("item_D_name") // 房屋預估資料流水號
                    )).ToList();
                }
                lsPBK.Insert(0, ("0", "請選擇"));
            }
            catch (Exception ex)
            {
                throw new Exception("AE_LoanSky.GetOrderRealEstateRequest", ex);
            }
            return lsPBK;
        }

        public List<(int, string)> GetBuildingState()
        {
            List<(int, string)> lsBuildingState = new List<(int, string)>{
                (1,"公寓(5樓含以下無電梯)"),
                (2,"透天厝"),
                (3,"店面(店鋪)"),
                (4,"辦公商業大樓"),
                (5,"住宅大樓(11層含以上有電梯)"),
                (6,"華廈(10層含以下有電梯)"),
                (7,"套房(1房1廳1衛)"),
                (8,"工廠"),
                (9,"農舍"),
                (10,"廠辦"),
                (11,"倉庫"),
                (12,"土地")
            };
            return lsBuildingState;
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
            runReq.housePre_res = runReq.GetHouse_Pre(req);
            SectionCode sectionCode; // 段代碼
            var errors = isRight(runReq.housePre_res, out sectionCode);
            if (errors.Count > 0) 
            {
                runReq.isNeedPopupWindow = true;
                runReq.message = string.Join(Environment.NewLine, errors);
                return runReq;
            }
            // 準備給LoanSky的資料
            runReq.housePre_res.MoiCityCode = sectionCode.city_num; // 縣市代碼(請參照對照表)
            runReq.housePre_res.MoiTownCode = sectionCode.area_num; // 鄉鎮市區代碼(請參照對照表)
            runReq.housePre_res.MoiSectionCode = sectionCode.road_num; // 段代碼
            runReq.housePre_res.Account = string.IsNullOrEmpty(account)? GetAllLoanSkyAccount(runReq.housePre_res.U_BC).Account : account; // 承辦人員帳號:測試用帳號:testCGU
            runReq.housePre_res.BusinessUserName = GetAllLoanSkyAccount(runReq.housePre_res.U_BC).branch_company;  //經辦人名稱:各分公司名稱
            runReq.housePre_res.BuildingState = AE2BuildingState(runReq.housePre_res.show_pre_building_kind);  // 建物類型(請參照對照表)
            runReq.housePre_res.ParkCategory = AE2ParkCategory(runReq.housePre_res.show_pre_parking_kind);  // 車位型態(請參照對照表)
            runReq.housePre_res.HA_cknum = req.HA_cknum; // 房屋預估資料流水號
            // KF2LoanSky
            runReq.KF2LoanSky();
            return runReq;
        }
        public ResultClass<string> House_pre_Update(HousePre_res model)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                ADOData _adoData = new ADOData();
                #region SQL

                var T_SQL = @"
                        update  House_pre
                          set 
                              pre_building_kind = ISNULL(@pre_building_kind, pre_building_kind),
                              BuildingState = isnull(@BuildingState, BuildingState),
                              ParkCategory = isnull(@ParkCategory, ParkCategory),
                              MoiCityCode = isnull(@MoiCityCode,MoiCityCode),
                              MoiTownCode = isnull(@MoiTownCode,MoiTownCode),
                              MoiSectionCode = isnull(@MoiSectionCode,MoiSectionCode),
                              pre_city = isnull(@pre_city,pre_city),
                              pre_area = isnull(@pre_area,pre_area),
                              pre_road = isnull(@pre_road,pre_road),
                              pre_road_s = isnull(@pre_road_s,pre_road_s),
                              edit_date = @edit_date,
                              edit_num = @edit_num
                         FROM House_pre
                         where HP_cknum = @hp_cknum
                ";
                var parameters = new List<SqlParameter>
                  {
                      new SqlParameter("@pre_building_kind", string.IsNullOrEmpty(model.pre_building_kind)? DBNull.Value :model.pre_building_kind ),
                      new SqlParameter("@BuildingState", string.IsNullOrEmpty(model.BuildingState)? DBNull.Value :model.BuildingState ),
                      new SqlParameter("@ParkCategory", string.IsNullOrEmpty(model.ParkCategory)? DBNull.Value : model.ParkCategory),
                      new SqlParameter("@MoiCityCode", string.IsNullOrEmpty(model.MoiCityCode) ? DBNull.Value : model.MoiCityCode),
                      new SqlParameter("@MoiTownCode", string.IsNullOrEmpty(model.MoiTownCode) ? DBNull.Value : model.MoiTownCode),
                      new SqlParameter("@MoiSectionCode", string.IsNullOrEmpty(model.MoiSectionCode)? DBNull.Value : model.MoiSectionCode),
                      new SqlParameter("@pre_city", string.IsNullOrEmpty(model.pre_city)? DBNull.Value : model.pre_city),
                      new SqlParameter("@pre_area", string.IsNullOrEmpty(model.pre_area)? DBNull.Value : model.pre_area),
                      new SqlParameter("@pre_road", string.IsNullOrEmpty(model.pre_road)? DBNull.Value : model.pre_road),
                      new SqlParameter("@pre_road_s", string.IsNullOrEmpty(model.pre_road_s)? DBNull.Value : model.pre_road_s),
                      new SqlParameter("@hp_cknum", model.HP_cknum),
                      new SqlParameter("@edit_date", DateTime.Today),
                      new SqlParameter("@edit_num", model.edit_num)
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
        public List<string> isRight(HousePre_res housePre_res, out SectionCode sectionCode)
        {
            List<string> errors = new List<string>();
            sectionCode = new SectionCode
            {
                city_name = housePre_res.pre_city,
                area_name = housePre_res.pre_area,
                road_name = housePre_res.pre_road,
                road_name_s = housePre_res.pre_road_s
            };
            errors = housePre_res.isRight();
            if (errors.Count > 0)
            {
                return errors;
            }
            sectionCode = GetMoiSectionCode(sectionCode); //取得段代碼
            if(sectionCode == null || string.IsNullOrEmpty(sectionCode.road_num))
            {
                errors.Add("找不到段代碼，請檢查對照表");
                return errors;
            }
            errors.AddRange(sectionCode.isRight()); 
            return errors;
        }
    }

    public class runOrderRealEstateRequest
    {
        ADOData _ADO = new();
        public OrderRealEstateRequest oreRequest { get; set; }   // LoanSky串接資料
        public HousePre_res housePre_res { get; set; }           // 來源資料
        public string message { get; set; }
        /// <summary>
        /// 是否需要彈跳視窗讓使用者確認LoanSky參數
        /// </summary>
        public bool isNeedPopupWindow { get; set; }

        public HousePre_res GetHouse_Pre(LoanSky_Req req)
        {
            try
            {
                #region SQL
                var parameters = new List<SqlParameter>();
                var T_SQL = @"
                    select
                         HP_id, HP_cknum
                        ,[pre_city],[pre_area],[pre_road],[pre_road_s],pre_apply_name
                        ,pre_building_kind, pre_build_num,pre_land_num
                        ,pre_note,pre_building_material_note,pre_principal_note,pre_other_note
                        ,(select CS_PID from House_apply where HA_cknum =@HA_cknum) as CS_PID
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
                parameters.Add(new SqlParameter("@HP_id", req.HP_id));
                parameters.Add(new SqlParameter("@HP_cknum", req.HP_cknum));
                parameters.Add(new SqlParameter("@HA_cknum", req.HA_cknum));
                #endregion
                DataTable dtResult = _ADO.ExecuteQuery(T_SQL, parameters);
                if (dtResult.Rows.Count > 0)
                {
                    // 來源資料
                    var housePre_res = dtResult.AsEnumerable().Select(row => new HousePre_res
                    {
                        HP_id = row.IsNull("HP_id") ? 0 : row.Field<decimal>("HP_id"), // 房屋預估資料ID
                        HP_cknum = row.IsNull("HP_cknum") ? "" : row.Field<string>("HP_cknum"), // 房屋預估資料流水號
                        add_name = row.IsNull("add_name") ? "" : row.Field<string>("add_name"),  //經辦人名稱
                        pre_apply_name = row.IsNull("pre_apply_name") ? "" : row.Field<string>("pre_apply_name"),    // 申請人
                        pre_building_kind = row.IsNull("pre_building_kind") ? "" : row.Field<string>("pre_building_kind"), // 建物類型
                        show_pre_building_kind = row.IsNull("show_pre_building_kind") ? "" : row.Field<string>("show_pre_building_kind"),  // 建物類型(請參照對照表)
                        pre_city = row.IsNull("pre_city") ? "" : row.Field<string>("pre_city"),// 縣市代碼(請參照對照表)
                        pre_area = row.IsNull("pre_area") ? "" : row.Field<string>("pre_area"),// 鄉鎮市區代碼(請參照對照表)
                        pre_road = row.IsNull("pre_road") ? "" : row.Field<string>("pre_road"),// 段(請參照對照表)
                        pre_road_s = row.IsNull("pre_road_s") ? "" : row.Field<string>("pre_road_s"),// 小段(請參照對照表)
                        U_BC = row.IsNull("U_BC") ? "" : row.Field<string>("U_BC"), // 分公司
                        pre_land_num = row.IsNull("pre_land_num") ? "" : row.Field<string>("pre_land_num"), // 地號
                        pre_build_num = row.IsNull("pre_build_num") ? "" : row.Field<string>("pre_build_num"), // 建號
                        pre_note = row.IsNull("pre_note") ? "" : row.Field<string>("pre_note"), // 備註
                        pre_building_material_note = row.IsNull("pre_building_material_note") ? "" : row.Field<string>("pre_building_material_note"), // 主要建築材料備註
                        pre_principal_note = row.IsNull("pre_principal_note") ? "" : row.Field<string>("pre_principal_note"), // 委託人備註
                        pre_other_note = row.IsNull("pre_other_note") ? "" : row.Field<string>("pre_other_note"), // 其他備註
                        CS_PID = row.IsNull("CS_PID") ? "" : row.Field<string>("CS_PID") // 公司統編/身分證號
                    }).FirstOrDefault();
                    return housePre_res;
                }
            }
            catch (Exception ex)
            {
                throw new Exception("AE_LoanSky.GetOrderRealEstateRequest", ex);
            }
            return null;
        }
        public void KF2LoanSky()
        {
            #region 案件資料
            oreRequest = new OrderRealEstateRequest
            {
                Account = housePre_res.Account, // 承辦人員帳號:測試用帳號:testCGU
                BusinessUserName = housePre_res.BusinessUserName ,  //經辦人名稱:各分公司名稱
                Applicant = housePre_res.pre_apply_name,    // 申請人
                IdNo = housePre_res.CS_PID,         // 公司統編/身分證號
                Condominium = housePre_res.pre_community_name,   // 社區(大樓)名稱
                BuildingState = housePre_res.BuildingState,  // 建物類型(請參照對照表)
                Situation = housePre_res.pre_use_kind,   //使用現況
                ParkCategory = housePre_res.ParkCategory,  // 車位型態(請參照對照表)
                Note = (string.IsNullOrEmpty(housePre_res.pre_note) ? "" : $"[建物類型備註]{housePre_res.pre_note}{Environment.NewLine}") +
                        (string.IsNullOrEmpty(housePre_res.pre_building_material_note) ? "" : $"[主要建築材料備註]{housePre_res.pre_building_material_note}{Environment.NewLine}") +
                        (string.IsNullOrEmpty(housePre_res.pre_principal_note) ? "" : $"[委託人備註]{housePre_res.pre_principal_note}{Environment.NewLine}") +
                        (string.IsNullOrEmpty(housePre_res.pre_other_note) ? "" : $"[其他備註]{housePre_res.pre_other_note}"), // 備註
                MoiCityCode = housePre_res.MoiCityCode, // 縣市代碼(請參照對照表)
                MoiTownCode = housePre_res.MoiTownCode // 鄉鎮市區代碼(請參照對照表)s
            };

            var no = new OrderRealEstateNoRequest
            {
                MoiSectionCode = housePre_res.MoiSectionCode,  // 段代碼:查詢條件：縣市代碼+區代碼+段名稱
                BuildNos = housePre_res.pre_build_num.Replace('－', '-').Replace('、', ',').Replace('；', ','),   // 建號 pre_build_num 多筆用逗號分隔
                LandNos = housePre_res.pre_land_num.Replace('－', '-').Replace('、', ',').Replace('；', ',') // 地號 pre_land_num 多筆用逗號分隔
            };
            oreRequest.Nos.Add(no);
            #endregion
            #region 取得AE要拋轉LoanSky案件資料-pdf附件
            List<OrderRealEstateAttachmentRequest> attachments = GetOrderRealEstateNoRequest(housePre_res.HA_cknum);
            if (attachments != null)
            {
                foreach (var item in attachments)
                {
                    oreRequest.Attachments.Add(item);
                }
            }
            #endregion
        }
        public List<OrderRealEstateAttachmentRequest> GetOrderRealEstateNoRequest(string HA_cknum)
        {
            string _storagePath = @"C:\UploadedFiles";
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
    }
}
