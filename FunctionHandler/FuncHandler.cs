using System.Data;

namespace KF_WebAPI.FunctionHandler
{
    public class FuncHandler
    {
        public static DataTable GetPage(DataTable dt, int page, int pageSize)
        {
            if (dt == null || dt.Rows.Count == 0 || page <= 0 || pageSize <= 0)
            {
                return new DataTable();
            }

            int startIndex = (page - 1) * pageSize;
            int endIndex = startIndex + pageSize;

            if (startIndex >= dt.Rows.Count)
            {
                return new DataTable();
            }

            if (endIndex > dt.Rows.Count)
            {
                endIndex = dt.Rows.Count; 
            }

            DataTable pageTable = dt.Clone();

            for (int i = startIndex; i < endIndex; i++)
            {
                pageTable.ImportRow(dt.Rows[i]);
            }

            return pageTable;
        }
    }
}
