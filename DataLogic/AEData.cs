
using KF_WebAPI.BaseClass;
using KF_WebAPI.BaseClass.AE;
using KF_WebAPI.BaseClass.Max104;
using KF_WebAPI.FunctionHandler;
using Microsoft.AspNetCore.DataProtection.KeyManagement;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Newtonsoft.Json;
using System;
using System.Data;
namespace KF_WebAPI.DataLogic
{
    public class AEData
    {
        ADOData _ADO = new();
        Common _Common = new();


        private Int32 GetFilesCountByKeyID(string p_Key, string p_Key_Type)
        {
            Common _Comm = new();
            Int32 m_FileCount =0;
            try
            {
                ADOData _adoData = new ADOData();
                var parameters = new List<SqlParameter>();
                #region SQL
                var T_SQL = @"SELECT count(*) FileCount FROM AE_Files WHERE KeyID = @KeyID  and Key_Type=@Key_Type";
                #endregion


                parameters.Add(new SqlParameter("@KeyID", p_Key));
                parameters.Add(new SqlParameter("@Key_Type", p_Key_Type));

                var dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
                if (dtResult.Rows.Count > 0)
                {
                    m_FileCount = Convert.ToInt32(dtResult.Rows[0]["FileCount"]);
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
            return m_FileCount;
        }


        private Int32 GetFilesMaxIndexByKeyID(string p_Key, string p_Key_Type)
        {
            Common _Comm = new();
            Int32 m_MaxIdx = 0;
            try
            {
                ADOData _adoData = new ADOData();
                var parameters = new List<SqlParameter>();
                #region SQL
                var T_SQL = @"SELECT max(cast(file_index as decimal)) MaxIdx FROM AE_Files WHERE KeyID = @KeyID  and Key_Type=@Key_Type";
                #endregion


                parameters.Add(new SqlParameter("@KeyID", p_Key));
                parameters.Add(new SqlParameter("@Key_Type", p_Key_Type));

                var dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
                if (dtResult.Rows.Count > 0)
                {
                    m_MaxIdx = Convert.ToInt32(dtResult.Rows[0]["MaxIdx"]);
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
            return m_MaxIdx;
        }


        public ResultClass<AE_Files[]> GetFilesByKeyID(string p_Key, string p_Key_Type)
        {

            ResultClass<AE_Files[]> resultClass = new();
            Common _Comm = new();
            Int32 _FileCount = GetFilesCountByKeyID(p_Key, p_Key_Type);   
            AE_Files[] m_AE_Files = new AE_Files[_FileCount];
            try
            {
                using (SqlConnection connection = new SqlConnection(_ADO.GetConnStr()))
                {
                    connection.Open();

                    using (SqlCommand command = connection.CreateCommand())
                    {
                        command.CommandText = "SELECT *,format(add_date,'yyyy/MM/dd')Add_YYMMDD FROM AE_Files WHERE KeyID = @KeyID  and Key_Type=@Key_Type";
                        command.Parameters.AddWithValue("@KeyID", p_Key);
                        command.Parameters.AddWithValue("@Key_Type", p_Key_Type);

                        Int32 m_Count = 0;
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            if (reader.HasRows)
                            {
                                while (reader.Read())
                                {
                                    byte[] blob = (byte[])reader["file_body_encode"];
                                    string base64String = Convert.ToBase64String(blob);
                                    AE_Files m_AE_File = new();
                                    m_AE_File.file_body_encode = _Comm.DecompressFile(base64String);
                                    m_AE_File.file_size = reader["file_size"].ToString();
                                    m_AE_File.file_index = reader["file_index"].ToString();
                                    m_AE_File.file_name = reader["file_name"].ToString();
                                    m_AE_File.content_type = reader["content_type"].ToString();
                                    m_AE_File.add_date = reader["Add_YYMMDD"].ToString();
                                    
                                    m_AE_Files[m_Count] = (m_AE_File);
                                    m_Count++;
                                }
                            }
                        }
                    }
                }
                resultClass.ResultCode = "000";
                resultClass.ResultMsg = "";
                resultClass.objResult = m_AE_Files;

            }
            catch (Exception ex)
            {
                throw ex;
            }
            return resultClass;
        }

        public ResultClass<AE_Files> GetFile(string Key, string Key_Type, string file_index = "")
        {
            Common _Comm = new();
            ResultClass<AE_Files> resultClass = new();
            AE_Files m_attachmentFile = new();
            try
            {
                string m_TypeName = "";

                using (SqlConnection connection = new SqlConnection(_ADO.GetConnStr()))
                {
                    connection.Open();

                    using (SqlCommand command = connection.CreateCommand())
                    {
                        command.CommandText = "SELECT * FROM AE_Files WHERE KeyID = @KeyID and  Key_Type = @Key_Type and file_index=@file_index ";
                        command.Parameters.AddWithValue("@KeyID", Key);
                        command.Parameters.AddWithValue("@Key_Type", m_TypeName);

                        if (file_index != "")
                        {
                            command.Parameters.AddWithValue("@file_index", file_index);
                        }

                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                byte[] blob = (byte[])reader["file_body_encode"];
                                string base64String = Convert.ToBase64String(blob);
                                m_attachmentFile.file_body_encode = _Comm.DecompressFile(base64String);
                                m_attachmentFile.file_size = reader["file_size"].ToString();
                                m_attachmentFile.file_index = reader["file_index"].ToString();
                                m_attachmentFile.file_name = reader["file_name"].ToString();
                                m_attachmentFile.content_type = reader["content_type"].ToString();
                            }
                        }
                    }
                }
                resultClass.ResultCode = "000";
                resultClass.ResultMsg = "";
                resultClass.objResult = m_attachmentFile;
            }
            catch (Exception ex)
            {
                resultClass.ResultCode = "999";
                resultClass.ResultMsg = ex.Message;
                resultClass.objResult = null;
            }
            return resultClass;
        }

        public int InsertFile( AE_Files[] p_attachmentFiles, string p_Key, string p_Key_Type, string p_User )
        {
            Common _Comm = new();
          
            int m_Execut = 0;
            Int32 _MaxIndex = GetFilesMaxIndexByKeyID(p_Key, p_Key_Type);

            try
            {
                using SqlConnection conn = new SqlConnection(_ADO.GetConnStr());
                // 開啟資料庫連線
                conn.Open();
                SqlTransaction transaction = conn.BeginTransaction();
                try
                {
                    foreach (AE_Files file in p_attachmentFiles)
                    {
                        // 建立 SQL 命令
                        string m_SQL = " INSERT INTO AE_Files (KeyID,Key_Type,file_index,file_body_encode,file_size,content_type,file_name,add_num)  ";
                        m_SQL += "  VALUES (@KeyID,@Key_Type,@file_index,@file_body_encode,@file_size,@content_type,@file_name,@add_num)  ";
                        using SqlCommand command = new SqlCommand(m_SQL, conn, transaction);
                        // 設定參數
                        string base64String = _Comm.CompressFile(file.file_body_encode);
                        byte[] imageBytes = Convert.FromBase64String(base64String);
                        command.Parameters.AddWithValue("@KeyID", p_Key);
                        command.Parameters.AddWithValue("@Key_Type", p_Key_Type);

                        command.Parameters.AddWithValue("@file_index",Convert.ToInt32(file.file_index)+ _MaxIndex);
                        command.Parameters.AddWithValue("@file_body_encode", imageBytes);
                        command.Parameters.AddWithValue("@file_size", file.file_size);
                        command.Parameters.AddWithValue("@content_type", file.content_type);
                        command.Parameters.AddWithValue("@file_name", file.file_name);
                        command.Parameters.AddWithValue("@add_num", p_User);
                     
                        // 執行 SQL 命令
                        m_Execut += command.ExecuteNonQuery();
                    }
                    transaction.Commit();
                    // 關閉資料庫連線
                    conn.Close();
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    throw new Exception("檔案上傳失敗!!"+ ex.Message);
                }
            }
            catch (Exception ex)
            {

                throw new Exception("檔案上傳失敗!!" + ex.Message);
            }

            return m_Execut;
        }

        public int DeleteFile( string p_KeyID, string p_Key_Type, string p_file_index)
        {
            Common _Comm = new();
            int m_Execut = 0;
            try
            {
               
                List<SqlParameter> Params = new List<SqlParameter>()
                {
                    new SqlParameter() {ParameterName = "@KeyID", SqlDbType = SqlDbType.VarChar, Value= p_KeyID},
                    new SqlParameter() {ParameterName = "@Key_Type", SqlDbType = SqlDbType.VarChar, Value= p_Key_Type},
                    new SqlParameter() {ParameterName = "@file_index", SqlDbType = SqlDbType.VarChar, Value= p_file_index}
                };
                m_Execut= _ADO.ExecuteNonQuery("Delete FROM dbo.AE_Files where KeyID=@KeyID and  Key_Type=@Key_Type and file_index=@file_index ", Params);
            }
            catch (Exception ex)
            {
                throw new Exception("刪除檔案失敗!!");
            }

            return m_Execut;
        }


        public string GetWebsiteURL()
        {
            Common _Comm = new();
            string m_URL = "";
             AE_Files m_attachmentFile = new();
            try
            {
                string m_Key = "websiteURL";

                ADOData _adoData = new ADOData();
                var parameters = new List<SqlParameter>();
                #region SQL
                var T_SQL = @"select item_D_name websiteURL  from Item_list where item_D_code=@Key";
                #endregion

                parameters.Add(new SqlParameter("@Key", m_Key));
               

                var dtResult = _adoData.ExecuteQuery(T_SQL, parameters);
                if (dtResult.Rows.Count > 0)
                {
                    m_URL = dtResult.Rows[0]["websiteURL"].ToString();
                }

            }
            catch (Exception ex)
            {
                m_URL = "";
            }
            return m_URL;
        }


    }
}
