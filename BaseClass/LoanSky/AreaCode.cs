using KF_WebAPI.FunctionHandler;
using Microsoft.Extensions.Primitives;

namespace KF_WebAPI.BaseClass.LoanSky
{
    public class AreaCode
    {
        FuncHandler _fun = new FuncHandler();

        /// <summary>
        /// 縣市代碼
        /// </summary>
        public string city_num { get; set; }

        /// <summary>
        /// 區代碼
        /// </summary>
        public string area_num { get; set; }

        private string _area_name = string.Empty;
        /// <summary>
        /// 區名稱
        /// </summary>
        public string area_name
        {
            get
            {
                return _fun.fromNCR(_area_name);
            }
            set
            {
                _area_name = value;
            }
        }

    }
}
