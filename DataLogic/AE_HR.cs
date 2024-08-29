using KF_WebAPI.FunctionHandler;
using System.Data;
using System.Data.SqlClient;
using KF_WebAPI.BaseClass;
using Newtonsoft.Json;
using System.Data.Common;
using System.Transactions;


namespace KF_WebAPI.DataLogic
{
    public class AE_HR
    {
        ADOData _ADO = new();
        Common _Common = new();
    }
}
