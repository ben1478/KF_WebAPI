

using System.Data;
using Microsoft.Data.SqlClient;
using System.Data.Common;


namespace KF_WebAPI.FunctionHandler
{
    public class ADOData
    {
        public  string ConnStr = "Data Source=ERP;Initial Catalog=AE_DB_TEST;User ID=sa;Password=juestcho;";


        public ADOData(string p_DB = "AE")
        {
            switch (p_DB)
            {
                case "AE":
                    ConnStr = "Data Source=ERP;Initial Catalog=AE_DB_TEST;User ID=sa;Password=juestcho;";
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

        public DataTable ExecuteQuery(string strSQL, List<SqlParameter> Param)
        {
            DataTable dtResult = new DataTable();
            try
            {
                using SqlConnection conn = new(ConnStr);
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

    }
}
