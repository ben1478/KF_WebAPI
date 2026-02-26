using KF_WebAPI.BaseClass;
using KF_WebAPI.BaseClass.AE;
using KF_WebAPI.FunctionHandler;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Newtonsoft.Json;
using System.Data;

namespace KF_WebAPI.DataLogic
{
    public class AE_FT
    {
        ADOData _adoData = new ADOData();
        FuncHandler _Fun = new FuncHandler();

        #region 業績折扣標準設定
        /// <summary>
        /// 取得業績折扣標準設定資料
        /// </summary>
        public ResultClass<string> Feat_M_LQuery(string bcType)
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            try
            {
                #region SQL
                var T_SQL_M = @"select * from Feat_rule where del_tag = '0' AND FR_M_type = 'Y' and U_BC = @U_BC order by FR_sort,FR_id";
                var T_SQL_D_All = @"select * from Feat_rule where del_tag = '0' AND FR_D_type = 'Y' and U_BC = @U_BC order by FR_D_ratio_A desc,FR_D_rate desc,FR_D_discount";
                var parameters_m = new List<SqlParameter>()
                {
                     new SqlParameter("@U_BC",bcType)
                };
                var parameters_d = new List<SqlParameter>()
                {
                    new SqlParameter("@U_BC",bcType)
                };
                #endregion
                var resultMList = _adoData.ExecuteQuery(T_SQL_M, parameters_m).AsEnumerable().Select(row => new Feat_M_res
                {
                    FR_id = row.Field<decimal>("FR_id"),
                    FR_M_code = row.Field<string>("FR_M_code"),
                    FR_M_name = _Fun.DeCodeBNWords(row.Field<string>("FR_M_name"))
                }).ToList();

                var allDetails = _adoData.ExecuteQuery(T_SQL_D_All, parameters_d).AsEnumerable().Select(row => new Feat_D
                {
                    FR_M_code = row.Field<string>("FR_M_code"),
                    FR_D_ratio_A = row.Field<decimal>("FR_D_ratio_A"),
                    FR_D_ratio_B = row.Field<decimal>("FR_D_ratio_B"),
                    FR_D_rate = row.Field<string>("FR_D_rate"),
                    FR_D_discount = row.Field<string>("FR_D_discount"),
                    FR_D_replace = row.Field<string>("FR_D_replace")
                }).ToList();

                foreach (var item in resultMList)
                {
                    item.feat_Ds = allDetails.Where(d => d.FR_M_code == item.FR_M_code).ToList();
                }

                resultClass.objResult = JsonConvert.SerializeObject(resultMList);
            }
            catch (Exception)
            {
                throw;
            }
            return resultClass;
        }

        /// <summary>
        /// 新增申貸方案
        /// </summary>
        public ResultClass<string> Feat_M_Ins(Feat_M_req model)
        {
            var resultClass = new ResultClass<string>();
            try
            {
                #region SQL
                var T_SQL = @"Insert into Feat_rule (FR_cknum,FR_M_type,FR_M_code,FR_M_name,FR_D_type,FR_D_code,FR_D_name,FR_D_rate,FR_D_discount,FR_D_replace,
                              FR_sort,del_tag,show_tag,add_date,add_num,add_ip,U_BC)
                              Values (@FR_cknum,'Y',@FR_M_code,@FR_M_name,'N','','','','','',99999,0,0,Getdate(),@add_num,@add_ip,@U_BC)";
                var parameters = new List<SqlParameter>()
                {
                    new SqlParameter("@FR_cknum",FuncHandler.GetCheckNum()),
                    new SqlParameter("@FR_M_code",model.FR_M_code),
                    new SqlParameter("@FR_M_name",model.FR_M_name),
                    new SqlParameter("@add_num",model.tbInfo.add_num),
                    new SqlParameter("@add_ip",model.tbInfo.add_ip),
                    new SqlParameter("@U_BC",model.U_BC)
                };
                #endregion
                int result = _adoData.ExecuteNonQuery(T_SQL, parameters);
                if (result == 0)
                {
                    resultClass.ResultCode = "400";
                    resultClass.ResultMsg = "新增失敗";
                }
                else
                {
                    resultClass.ResultCode = "000";
                    resultClass.ResultMsg = "新增成功";
                }
            }
            catch (Exception)
            {
                throw;
            }

            return resultClass;
        }

        /// <summary>
        /// 刪除申貸方案
        /// </summary>
        public ResultClass<string> Feat_M_Del(string FR_M_code, string user, string bcType,string clientIp)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                #region SQL
                var T_SQL = @"Update Feat_rule Set del_tag = '1',del_date=Getdate(),del_num=@User,del_ip=@IP  where FR_M_code=@FR_M_code and del_tag='0' and U_BC=@u_bc";
                var parameters = new List<SqlParameter>()
                {
                    new SqlParameter("@User",user),
                    new SqlParameter("@IP",clientIp),
                    new SqlParameter("@FR_M_code",FR_M_code),
                    new SqlParameter("@u_bc",bcType)
                };
                #endregion
                int result = _adoData.ExecuteNonQuery(T_SQL, parameters);
                if (result == 0) 
                {
                    resultClass.ResultCode = "400";
                    resultClass.ResultMsg = "刪除失敗";
                }
                else
                {
                    resultClass.ResultCode = "000";
                    resultClass.ResultMsg = "刪除成功";
                }
            }
            catch (Exception)
            {
                throw;
            }
            return resultClass;
        }

        /// <summary>
        /// 取得方案折扣設定
        /// </summary>
        public ResultClass<string> Feat_D_LQuery(string FR_M_code, string bcType)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                #region SQL
                var T_SQL = @"select * from Feat_rule where del_tag = '0' AND FR_D_type = 'Y' AND FR_M_code = @FR_M_code and U_BC=@U_BC
                                order by FR_D_ratio_A desc,FR_D_rate desc,FR_D_discount";
                var parameters = new List<SqlParameter>()
                {
                    new SqlParameter("@FR_M_code",FR_M_code),
                    new SqlParameter("@U_BC",bcType)
                };
                #endregion
                var resultList = _adoData.ExecuteQuery(T_SQL, parameters).AsEnumerable().Select(row => new
                {
                    FR_id = row.Field<decimal>("FR_id"),
                    FR_D_ratio_A = row.Field<decimal>("FR_D_ratio_A"),
                    FR_D_ratio_B = row.Field<decimal>("FR_D_ratio_B"),
                    FR_D_rate = row.Field<string>("FR_D_rate"),
                    FR_D_discount = row.Field<string>("FR_D_discount"),
                    FR_D_replace = row.Field<string>("FR_D_replace")
                }).ToList();

                resultClass.objResult = JsonConvert.SerializeObject(resultList);
            }
            catch (Exception)
            {
                throw;
            }
            return resultClass;
        }

        /// <summary>
        /// 刪除單筆折扣
        /// </summary>
        public ResultClass<string> Feat_D_Del(string FR_id, string user, string clientIp)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                #region SQL
                var T_SQL = @"Update Feat_rule Set del_tag = '1',del_date=Getdate(),del_num=@User,del_ip=@IP  where FR_id=@FR_id";
                var parameters = new List<SqlParameter>()
                {
                    new SqlParameter("@User",user),
                    new SqlParameter("@IP",clientIp),
                    new SqlParameter("@FR_id",FR_id)
                };
                #endregion
                int result = _adoData.ExecuteNonQuery(T_SQL, parameters);
                if (result == 0)
                {
                    resultClass.ResultCode = "400";
                    resultClass.ResultMsg = "刪除失敗";
                }
                else
                {
                    resultClass.ResultCode = "000";
                    resultClass.ResultMsg = "刪除成功";
                }

            }
            catch (Exception)
            {
                throw;
            }
            return resultClass;
        }

        /// <summary>
        /// 變更方案折扣設定
        /// </summary>
        public ResultClass<string> Feat_D_Upd(List<Feat_D> modelList)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                #region SQL
                var T_SQL_OLD = @"select * from Feat_rule where del_tag = '0' AND FR_D_type = 'Y' AND FR_M_code = @FR_M_code and U_BC=@U_BC
                                order by FR_D_ratio_A desc,FR_D_rate desc,FR_D_discount";
                var parameters_old = new List<SqlParameter>()
                {
                    new SqlParameter("@FR_M_code",modelList[0].FR_M_code),
                    new SqlParameter("@U_BC",modelList[0].U_BC)
                };
                #endregion
                var dbData = _adoData.ExecuteQuery(T_SQL_OLD, parameters_old).AsEnumerable().Select(row => new Feat_D
                {
                    FR_id = row.Field<decimal>("FR_id"),
                    FR_D_ratio_A = row.Field<decimal>("FR_D_ratio_A"),
                    FR_D_ratio_B = row.Field<decimal>("FR_D_ratio_B"),
                    FR_D_rate = row.Field<string>("FR_D_rate"),
                    FR_D_discount = row.Field<string>("FR_D_discount"),
                    FR_D_replace = row.Field<string>("FR_D_replace")
                }).ToList();

                var dbDict = dbData.ToDictionary(x => x.FR_id);

                foreach (var item in modelList) 
                {
                    if (!item.FR_id.HasValue || item.FR_id == 0)
                    {
                        InsertFeatRule(item);
                    }
                    else
                    {
                        if (dbDict.TryGetValue(item.FR_id, out var existing))
                        {
                            bool isChanged =
                                     item.FR_D_ratio_A != existing.FR_D_ratio_A ||
                                     item.FR_D_ratio_B != existing.FR_D_ratio_B ||
                                     item.FR_D_rate != existing.FR_D_rate ||
                                     item.FR_D_discount != existing.FR_D_discount ||
                                     item.FR_D_replace != existing.FR_D_replace;

                            if (isChanged)
                            {
                                UpdateFeatRule(item);
                            }
                        }
                    }
                }

                resultClass.ResultCode = "000";
                resultClass.ResultMsg = "儲存成功";
            }
            catch (Exception)
            {
                throw;
            }
            return resultClass;
        }
        #endregion

        #region 佣金標準設定
        /// <summary>
        /// 取得新鑫佣金標準設定資料
        /// </summary>
        public ResultClass<string> Feat_NM_LQuery()
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            try
            {
                #region SQL_M
                var T_SQL_M = @"select * from Feat_rule_comm where del_tag = '0' AND FR_M_type = 'Y' and U_BC='general' order by FR_sort,FR_id";
                var T_SQL_D_All = @"select * from Feat_rule_comm where del_tag = '0' AND FR_D_type = 'Y' and U_BC= 'general' order by FR_D_ratio_A desc,FR_D_rate desc,FR_D_discount ";
                #endregion
                var resultMList = _adoData.ExecuteSQuery(T_SQL_M).AsEnumerable().Select(row => new Feat_M_res
                {
                    FR_id = row.Field<decimal>("FR_id"),
                    FR_M_code = row.Field<string>("FR_M_code"),
                    FR_M_name = _Fun.DeCodeBNWords(row.Field<string>("FR_M_name"))
                }).ToList();

                var allDetails = _adoData.ExecuteSQuery(T_SQL_D_All).AsEnumerable().Select(row => new Feat_D
                {
                    FR_M_code = row.Field<string>("FR_M_code"),
                    FR_D_ratio_A = row.Field<decimal>("FR_D_ratio_A"),
                    FR_D_ratio_B = row.Field<decimal>("FR_D_ratio_B"),
                    FR_D_rate = row.Field<string>("FR_D_rate"),
                    FR_D_discount = row.Field<string>("FR_D_discount"),
                    FR_D_replace = row.Field<string>("FR_D_replace")
                }).ToList();

                foreach (var item in resultMList)
                {
                    item.feat_Ds = allDetails.Where(d => d.FR_M_code == item.FR_M_code).ToList();
                }

                resultClass.objResult = JsonConvert.SerializeObject(resultMList);
            }
            catch (Exception)
            {
                throw;
            }
            return resultClass;
        }

        /// <summary>
        /// 新增新鑫申貸方案
        /// </summary>
        public ResultClass<string> Feat_NM_Ins(Feat_M_req model)
        {
            var resultClass = new ResultClass<string>();
            try
            {
                #region SQL
                var T_SQL = @"Insert into Feat_rule_comm (FR_cknum,FR_M_type,FR_M_code,FR_M_name,FR_D_type,FR_D_code,FR_D_name,FR_D_rate,FR_D_discount,FR_D_replace,
                              FR_sort,del_tag,show_tag,add_date,add_num,add_ip,U_BC)
                              Values (@FR_cknum,'Y',@FR_M_code,@FR_M_name,'N','','','','','',99999,0,0,Getdate(),@add_num,@add_ip,'general')";
                var parameters = new List<SqlParameter>()
                {
                    new SqlParameter("@FR_cknum",FuncHandler.GetCheckNum()),
                    new SqlParameter("@FR_M_code",model.FR_M_code),
                    new SqlParameter("@FR_M_name",model.FR_M_name),
                    new SqlParameter("@add_num",model.tbInfo.add_num),
                    new SqlParameter("@add_ip",model.tbInfo.add_ip)
                };
                #endregion
                int result = _adoData.ExecuteNonQuery(T_SQL, parameters);
                if (result == 0)
                {
                    resultClass.ResultCode = "400";
                    resultClass.ResultMsg = "新增失敗";
                }
                else
                {
                    resultClass.ResultCode = "000";
                    resultClass.ResultMsg = "新增成功";
                }
            }
            catch (Exception)
            {
                throw;
            }

            return resultClass;
        }

        /// <summary>
        /// 刪除新鑫申貸方案
        /// </summary>
        public ResultClass<string> Feat_NM_Del(string FR_M_code, string user, string bcType, string clientIp)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                #region SQL
                var T_SQL = @"Update Feat_rule_comm Set del_tag = '1',del_date=Getdate(),del_num=@User,del_ip=@IP  where FR_M_code=@FR_M_code and del_tag='0' and U_BC= 'general'";
                var parameters = new List<SqlParameter>()
                {
                    new SqlParameter("@User",user),
                    new SqlParameter("@IP",clientIp),
                    new SqlParameter("@FR_M_code",FR_M_code),
                    new SqlParameter("@u_bc",bcType)
                };
                #endregion
                int result = _adoData.ExecuteNonQuery(T_SQL, parameters);
                if (result == 0)
                {
                    resultClass.ResultCode = "400";
                    resultClass.ResultMsg = "刪除失敗";
                }
                else
                {
                    resultClass.ResultCode = "000";
                    resultClass.ResultMsg = "刪除成功";
                }
            }
            catch (Exception)
            {
                throw;
            }
            return resultClass;
        }

        /// <summary>
        /// 取得新鑫方案折扣設定
        /// </summary>
        public ResultClass<string> Feat_ND_LQuery(string FR_M_code, string bcType)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                #region SQL
                var T_SQL = @"select * from Feat_rule_comm where del_tag = '0' AND FR_D_type = 'Y' AND FR_M_code = @FR_M_code and U_BC = 'general'
                            order by FR_D_ratio_A desc,FR_D_rate desc,FR_D_discount";
                var parameters = new List<SqlParameter>()
                {
                    new SqlParameter("@FR_M_code",FR_M_code)
                };
                #endregion
                var resultList = _adoData.ExecuteQuery(T_SQL, parameters).AsEnumerable().Select(row => new
                {
                    FR_id = row.Field<decimal>("FR_id"),
                    FR_D_ratio_A = row.Field<decimal>("FR_D_ratio_A"),
                    FR_D_ratio_B = row.Field<decimal>("FR_D_ratio_B"),
                    FR_D_rate = row.Field<string>("FR_D_rate"),
                    FR_D_discount = row.Field<string>("FR_D_discount"),
                    FR_D_replace = row.Field<string>("FR_D_replace")
                }).ToList();

                resultClass.objResult = JsonConvert.SerializeObject(resultList);
            }
            catch (Exception)
            {
                throw;
            }
            return resultClass;
        }

        /// <summary>
        /// 刪除新鑫單筆折扣
        /// </summary>
        public ResultClass<string> Feat_ND_Del(string FR_id, string user, string clientIp)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                #region SQL
                var T_SQL = @"Update Feat_rule_comm Set del_tag = '1',del_date=Getdate(),del_num=@User,del_ip=@IP  where FR_id=@FR_id";
                var parameters = new List<SqlParameter>()
                {
                    new SqlParameter("@User",user),
                    new SqlParameter("@IP",clientIp),
                    new SqlParameter("@FR_id",FR_id)
                };
                #endregion
                int result = _adoData.ExecuteNonQuery(T_SQL, parameters);
                if (result == 0)
                {
                    resultClass.ResultCode = "400";
                    resultClass.ResultMsg = "刪除失敗";
                }
                else
                {
                    resultClass.ResultCode = "000";
                    resultClass.ResultMsg = "刪除成功";
                }

            }
            catch (Exception)
            {
                throw;
            }
            return resultClass;
        }

        /// <summary>
        /// 變更新鑫方案折扣設定
        /// </summary>
        public ResultClass<string> Feat_ND_Upd(List<Feat_D> modelList)
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {
                #region SQL
                var T_SQL_OLD = @"select * from Feat_rule_comm where del_tag = '0' AND FR_D_type = 'Y' AND FR_M_code = @FR_M_code and U_BC=@U_BC
                                order by FR_D_ratio_A desc,FR_D_rate desc,FR_D_discount";
                var parameters_old = new List<SqlParameter>()
                {
                    new SqlParameter("@FR_M_code",modelList[0].FR_M_code),
                    new SqlParameter("@U_BC",modelList[0].U_BC)
                };
                #endregion
                var dbData = _adoData.ExecuteQuery(T_SQL_OLD, parameters_old).AsEnumerable().Select(row => new Feat_D
                {
                    FR_id = row.Field<decimal>("FR_id"),
                    FR_D_ratio_A = row.Field<decimal>("FR_D_ratio_A"),
                    FR_D_ratio_B = row.Field<decimal>("FR_D_ratio_B"),
                    FR_D_rate = row.Field<string>("FR_D_rate"),
                    FR_D_discount = row.Field<string>("FR_D_discount"),
                    FR_D_replace = row.Field<string>("FR_D_replace")
                }).ToList();

                var dbDict = dbData.ToDictionary(x => x.FR_id);

                foreach (var item in modelList)
                {
                    if (!item.FR_id.HasValue || item.FR_id == 0)
                    {
                        InsertFeatRuleN(item);
                    }
                    else
                    {
                        if (dbDict.TryGetValue(item.FR_id, out var existing))
                        {
                            bool isChanged =
                                     item.FR_D_ratio_A != existing.FR_D_ratio_A ||
                                     item.FR_D_ratio_B != existing.FR_D_ratio_B ||
                                     item.FR_D_rate != existing.FR_D_rate ||
                                     item.FR_D_discount != existing.FR_D_discount ||
                                     item.FR_D_replace != existing.FR_D_replace;

                            if (isChanged)
                            {
                                UpdateFeatRuleN(item);
                            }
                        }
                    }
                }

                resultClass.ResultCode = "000";
                resultClass.ResultMsg = "儲存成功";
            }
            catch (Exception)
            {
                throw;
            }
            return resultClass;
        }

        /// <summary>
        /// 取得國峯佣金標準設定資料
        /// </summary>
        public ResultClass<string> Feat_KF_LQuery()
        {
            ResultClass<string> resultClass = new ResultClass<string>();
            try
            {
                #region SQL_M
                var T_SQL_M = @"select * from Item_list where del_tag = '0' AND item_M_code='Return' and item_D_name like '%國%'  order by item_sort,item_id";
                #endregion
                var resultMList = _adoData.ExecuteSQuery(T_SQL_M).AsEnumerable().Select(row => new
                {
                    item_id = row.Field<decimal>("item_id"),
                    item_show = row.Field<string>("item_D_code") + '-' + _Fun.DeCodeBNWords(row.Field<string>("item_D_name")),
                    item_int = row.Field<int>("item_D_int_A"),
                    show_tag = row.Field<string>("show_tag")
                }).ToList();

                resultClass.objResult = JsonConvert.SerializeObject(resultMList);
            }
            catch (Exception)
            {
                throw;
            }
            return resultClass;
        }

        /// <summary>
        /// 修改國峯佣金標準設定資料
        /// </summary>
        public ResultClass<string> Feat_KF_Upd()
        {
            ResultClass<string> resultClass = new ResultClass<string>();

            try
            {

            }
            catch (Exception)
            {
                throw;
            }
            return resultClass;
        }
        #endregion

        void InsertFeatRule(Feat_D model)
        {
            try
            {
                #region SQL
                var T_SQL_IN = @"Insert into Feat_rule (FR_cknum,FR_M_type,FR_M_code,FR_M_name,FR_D_type,FR_D_code,FR_D_name,FR_D_ratio_A,FR_D_ratio_B,FR_D_rate,
                           FR_D_discount,FR_D_replace,FR_sort,del_tag,show_tag,add_date,add_num,add_ip,U_BC) 
                           Values (@FR_cknum,'N',@FR_M_code,@FR_M_name,'Y','','',@FR_D_ratio_A,@FR_D_ratio_B,@FR_D_rate,@FR_D_discount,
                           @FR_D_replace,99999,0,0,Getdate(),@add_num,@add_ip,@U_BC)";
                #endregion
                var parameters_in = new List<SqlParameter>()
                {
                    new SqlParameter("@FR_cknum",FuncHandler.GetCheckNum()),
                    new SqlParameter("@FR_M_code",model.FR_M_code),
                    new SqlParameter("@FR_M_name",model.FR_M_name),
                    new SqlParameter("@FR_D_ratio_A",model.FR_D_ratio_A),
                    new SqlParameter("@FR_D_ratio_B",model.FR_D_ratio_B),
                    new SqlParameter("@FR_D_rate",model.FR_D_rate),
                    new SqlParameter("@FR_D_discount",model.FR_D_discount),
                    new SqlParameter("@FR_D_replace",model.FR_D_replace),
                    new SqlParameter("@add_num",model.tbInfo.add_num),
                    new SqlParameter("@add_ip",model.tbInfo.add_ip),
                    new SqlParameter("@U_BC",model.U_BC)
                };
                _adoData.ExecuteNonQuery(T_SQL_IN, parameters_in);
            }
            catch (Exception)
            {
                throw new Exception("DB 寫入失敗");
            }
        }

        void UpdateFeatRule(Feat_D model)
        {
            try
            {
                #region SQL_Upd
                var T_SQL = @"Update Feat_rule Set FR_D_ratio_A = @FR_D_ratio_A,FR_D_ratio_B = @FR_D_ratio_B,FR_D_rate = @FR_D_rate,
                                              FR_D_discount = @FR_D_discount,FR_D_replace = @FR_D_replace,edit_date = Getdate(),
                                              edit_num = @edit_num,edit_ip = @IP Where FR_id = @FR_id";
                var parameters = new List<SqlParameter>()
                                {
                                    new SqlParameter("@FR_D_ratio_A",model.FR_D_ratio_A),
                                    new SqlParameter("@FR_D_ratio_B",model.FR_D_ratio_B),
                                    new SqlParameter("@FR_D_rate",model.FR_D_rate),
                                    new SqlParameter("@FR_D_discount",model.FR_D_discount),
                                    new SqlParameter("@FR_D_replace",model.FR_D_replace),
                                    new SqlParameter("@edit_num",model.tbInfo.edit_num),
                                    new SqlParameter("@IP",model.tbInfo.add_ip),
                                    new SqlParameter("@FR_id",model.FR_id)
                                };
                #endregion
                _adoData.ExecuteNonQuery(T_SQL, parameters);
            }
            catch (Exception)
            {
                throw new Exception("DB 異動失敗");
            }
        }

        void InsertFeatRuleN(Feat_D model)
        {
            try
            {
                #region SQL
                var T_SQL_IN = @"Insert into Feat_rule_comm (FR_cknum,FR_M_type,FR_M_code,FR_M_name,FR_D_type,FR_D_code,FR_D_name,FR_D_ratio_A,FR_D_ratio_B,FR_D_rate,
                           FR_D_discount,FR_sort,del_tag,show_tag,add_date,add_num,add_ip,U_BC) 
                           Values (@FR_cknum,'N',@FR_M_code,@FR_M_name,'Y','','',@FR_D_ratio_A,@FR_D_ratio_B,@FR_D_rate,@FR_D_discount,99999,0,0,Getdate(),@add_num,@add_ip,@U_BC)";
                #endregion
                var parameters_in = new List<SqlParameter>()
                {
                    new SqlParameter("@FR_cknum",FuncHandler.GetCheckNum()),
                    new SqlParameter("@FR_M_code",model.FR_M_code),
                    new SqlParameter("@FR_M_name",model.FR_M_name),
                    new SqlParameter("@FR_D_ratio_A",model.FR_D_ratio_A),
                    new SqlParameter("@FR_D_ratio_B",model.FR_D_ratio_B),
                    new SqlParameter("@FR_D_rate",model.FR_D_rate),
                    new SqlParameter("@FR_D_discount",model.FR_D_discount),
                    new SqlParameter("@add_num",model.tbInfo.add_num),
                    new SqlParameter("@add_ip",model.tbInfo.add_ip),
                    new SqlParameter("@U_BC",model.U_BC)
                };
                _adoData.ExecuteNonQuery(T_SQL_IN, parameters_in);
            }
            catch (Exception)
            {
                throw new Exception("DB 寫入失敗");
            }
        }

        void UpdateFeatRuleN(Feat_D model)
        {
            try
            {
                #region SQL_Upd
                var T_SQL = @"Update Feat_rule_comm Set FR_D_ratio_A = @FR_D_ratio_A,FR_D_ratio_B = @FR_D_ratio_B,FR_D_rate = @FR_D_rate,
                                              FR_D_discount = @FR_D_discount,edit_date = Getdate(),
                                              edit_num = @edit_num,edit_ip = @IP Where FR_id = @FR_id";
                var parameters = new List<SqlParameter>()
                                {
                                    new SqlParameter("@FR_D_ratio_A",model.FR_D_ratio_A),
                                    new SqlParameter("@FR_D_ratio_B",model.FR_D_ratio_B),
                                    new SqlParameter("@FR_D_rate",model.FR_D_rate),
                                    new SqlParameter("@FR_D_discount",model.FR_D_discount),
                                    new SqlParameter("@edit_num",model.tbInfo.edit_num),
                                    new SqlParameter("@IP",model.tbInfo.add_ip),
                                    new SqlParameter("@FR_id",model.FR_id)
                                };
                #endregion
                _adoData.ExecuteNonQuery(T_SQL, parameters);
            }
            catch (Exception)
            {
                throw new Exception("DB 異動失敗");
            }
        }
    }
}
