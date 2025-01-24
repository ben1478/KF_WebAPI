using System.Collections.Generic;
namespace KF_WebAPI.BaseClass.Max104
{
    /// <summary>
    /// 單一data
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class ResultClass_104<T>
    {
        public string? code { get; set; }
        public string? msg { get; set; }
        public T? data { get; set; }

    }
    /// <summary>
    /// 多筆data
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class arrResultClass_104<T>
    {
        public string? code { get; set; }
        public string? msg { get; set; }
        public T[]? data { get; set; }
    }
}
