using KF_WebAPI.FunctionHandler;
using Microsoft.Extensions.Primitives;

namespace KF_WebAPI.BaseClass.LoanSky
{
    public class SectionCode
    {
        FuncHandler _fun = new FuncHandler();
        /// <summary>
        /// 縣市代碼
        /// </summary>
        public string city_num { get; set; }
        /// <summary>
        /// 縣市名稱
        /// </summary>
        public string city_name { get; set; }

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



        /// <summary>
        /// 段代碼
        /// </summary>
        public string road_num { get; set; }
        
        private string _road_name=string.Empty;
        /// <summary>
        /// 段名稱
        /// </summary>
        public string road_name
        {
            get
            {
                return _fun.fromNCR(_road_name);
            }
            set
            {
                _road_name = value;
            }
        }

        private string _road_name_s=string.Empty;
        /// <summary>
        /// 段.小段名稱
        /// </summary>
        public string road_name_s 
        { 
            get
            {
                return _fun.fromNCR(_road_name_s);
            }
            set
            {
                _road_name_s = value;
            }
        }

        private string _message;
        public string message {
            
            get { return _message; }
        }
        /// <summary>
        /// 判斷參數是否正確。true:正確; false:錯誤
        /// </summary>
        public List<string> isRight()
        {
            List<string> errors = new List<string>();
            if (string.IsNullOrEmpty(city_num))
            {
                errors.Add("縣市代碼為Null");
            }
            if (string.IsNullOrEmpty(area_num))
            {
                errors.Add("鄉鎮市區為Null");
            }
            if (string.IsNullOrEmpty(road_num))
            {
                errors.Add("段代碼為Null");
            }
            return errors;
        }
    }
}
