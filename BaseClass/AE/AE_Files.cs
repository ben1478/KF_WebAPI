namespace KF_WebAPI.BaseClass.AE
{

    public class AE_Files
    {
        /// <summary>
        /// 檔案編碼索引
        /// </summary>
        public String? file_index { get; set; }

        /// <summary>
        /// 檔案主體
        /// </summary>
        public String? file_body_encode { get; set; }

        /// <summary>
        /// 檔案大小
        /// </summary>
        public String? file_size { get; set; }

        /// <summary>
        /// 檔案格式
        /// </summary>
        public String? content_type { get; set; }
        /// <summary>
        /// 檔案格式
        /// </summary>
        public String? file_name { get; set; }

    }


}
