

using System.Data;
using Microsoft.Data.SqlClient;
using System.Data.Common;
using KF_WebAPI.BaseClass.Max104;
using System.Reflection;

namespace KF_WebAPI.FunctionHandler
{
    public class ADOData
    {
        public  string ConnStr = "Data Source=ERP;Initial Catalog=AE_DB_TEST;User ID=sa;Password=juestcho;";

        /// <summary>
        /// 根據物件屬性產生Datatable的欄位
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="objects"></param>
        /// <returns></returns>
        public static DataTable CreateDataTableFromClass<T>(T[] objects)
        {
            DataTable dataTable = new DataTable();

            // 1. 動態建立 DataTable 結構
            PropertyInfo[] properties = typeof(T).GetProperties();
            foreach (var prop in properties)
            {
                dataTable.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
            }

            // 2. 填充資料
            foreach (var obj in objects)
            {
                DataRow row = dataTable.NewRow();
                foreach (var prop in properties)
                {
                    row[prop.Name] = prop.GetValue(obj) ?? DBNull.Value;
                }
                dataTable.Rows.Add(row);
            }

            return dataTable;
        }

        public void DataTableToSQL<T>(string TableName, T[] objects, string p_ConnStr)
        {

            DataTable dataTable = CreateDataTableFromClass(objects);

           
            using (SqlConnection connection = new SqlConnection(p_ConnStr))
            {
                connection.Open();
                using (SqlBulkCopy bulkCopy = new SqlBulkCopy(connection))
                {
                    bulkCopy.DestinationTableName = TableName; // 替換成你的資料表名稱

                    // 動態設定 ColumnMappings
                    foreach (DataColumn column in dataTable.Columns)
                    {
                        bulkCopy.ColumnMappings.Add(column.ColumnName, column.ColumnName);
                    }

                    // 寫入資料
                    bulkCopy.WriteToServer(dataTable);
                }
            }
        }


        public ADOData(string p_DB = "AE")
        {
            switch (p_DB)
            {
                case "AE":
                    ConnStr = "Data Source=ERP;Initial Catalog=AE_DB;User ID=sa;Password=juestcho;TrustServerCertificate=True;";
                    break;
                case "Other":
                    ConnStr = "Data Source=ERP;Initial Catalog=Other;User ID=sa;Password=juestcho;";
                    break;
            }
        }

        public string GetConnStr()
        {
            return ConnStr;
        }

        public Int32 ExecuteNonQuery(string strSQL, List<SqlParameter> Param)
        {
            Int32 m_Execut = 0;
            try
            {
                using SqlConnection conn = new(ConnStr);
                conn.Open();
                using (SqlCommand cmd = new SqlCommand(strSQL, conn))
                {
                    if (Param != null)
                    {
                        foreach (DbParameter param in Param)
                        {
                            cmd.Parameters.Add(param);
                        }
                    }
                    m_Execut = cmd.ExecuteNonQuery();
                    conn.Close();
                }
            }
            catch
            {
                throw;
            }
            return m_Execut;
        }

        public DataTable ExecuteQuery(string strSQL, List<SqlParameter> Param, Boolean isTrim = false)
        {
            DataTable dtResult = new DataTable();
            try
            {
                using SqlConnection conn = new(ConnStr);
                if (isTrim)
                {
                    strSQL = strSQL.Replace("\r\n", " ").Replace("\r", " ").Replace("\n", " ");
                }
                using (SqlCommand cmd = new SqlCommand(strSQL, conn))
                {
                    if (Param != null)
                    {
                        foreach (DbParameter param in Param)
                        {
                            cmd.Parameters.Add(param);
                        }
                    }
                    using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                    {
                        adapter.Fill(dtResult);
                    }
                }
            }
            catch 
            {
                throw;
            }
            return dtResult;
        }

        public DataTable ExecuteSQuery(string strSQL)
        {
            DataTable dtResult = new DataTable();
            try
            {
                using SqlConnection conn = new(ConnStr);
                using SqlCommand cmd = new SqlCommand(strSQL, conn);
                using SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                adapter.Fill(dtResult);
            }
            catch (Exception ex)
            {
                throw;
            }
            return dtResult;
        }
        public object ExecuteScalar(string strSQL, List<SqlParameter> Param)
        {
            object result = null;
            try
            {
                using SqlConnection conn = new(ConnStr);
                using SqlCommand cmd = new SqlCommand(strSQL, conn);
                if (Param != null)
                {
                    foreach (DbParameter param in Param)
                    {
                        cmd.Parameters.Add(param);
                    }
                    conn.Open();
                    result = cmd.ExecuteScalar();
                }
            }
            catch (Exception)
            {

                throw;
            }
            return result;
        }

    }
}
