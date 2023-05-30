
using KF_WebAPI.FunctionHandler;
using System.Data;
using System.Data.SqlClient;
using KF_WebAPI.BaseClass;
using Microsoft.AspNetCore.DataProtection.KeyManagement;
using Newtonsoft.Json;
using Microsoft.AspNetCore.Mvc;
using System.Reflection;
using static KF_WebAPI.Controllers.YuRichAPIController;
using System.Data.Common;
using Microsoft.AspNetCore.Http;
using System.Xml.Linq;
using System.Transactions;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;

namespace KF_WebAPI.FunctionHandler
{
    public class ADOData
    {
        public string _ConnStr = "Data Source=ERP;Initial Catalog=AE_DB;User ID=sa;Password=juestcho;";
        public string _ConnStr1 = "Data Source=tcp:192.168.1.27\\SQLEXPRESS,50093;Initial Catalog=EIPDATA;User ID=KFWeb;Password=Kf52611690;";

        public DataTable ExecuteQuery_EIP(string strSQL, List<SqlParameter> Param)
        {
            DataTable dtResult = new DataTable();
            try
            {
                using SqlConnection conn = new(_ConnStr1);
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
            catch (Exception ex)
            {
                throw ex;
            }
            return dtResult;
        }



        public Int32 ExecuteNonQuery(string strSQL, List<SqlParameter> Param)
        {
            Int32 m_Execut = 0;
            try
            {
                using SqlConnection conn = new(_ConnStr);
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
                using SqlConnection conn = new(_ConnStr);
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
            catch (Exception ex)
            {
                throw ex;
            }
            return dtResult;
        }



    }
}
