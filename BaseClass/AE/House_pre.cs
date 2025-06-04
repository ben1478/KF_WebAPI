using Microsoft.Extensions.Primitives;
using System.Text;

namespace KF_WebAPI.BaseClass.AE
{
    public class HousePre_res
    {
        public decimal HP_id { get; set; }
        public string? HP_cknum { get; set; }
        public decimal? HA_id { get; set; }
        public string? pre_apply_name { get; set; }
        public DateTime? pre_apply_date { get; set; }
        public string? pre_zip_code { get; set; }
        public string? pre_city { get; set; }
        public string? pre_area { get; set; }
        public string? pre_road { get; set; }
        public string? pre_road_s { get; set; }
        public string? pre_land_num { get; set; }
        public string? pre_other_note { get; set; }
        public string? pre_present_value { get; set; }
        public string? pre_transfer { get; set; }
        public string? pre_transfer_date { get; set; }
        public string? pre_share_area_m { get; set; }
        public string? pre_share_area_p { get; set; }
        public string? pre_address { get; set; }
        public string? pre_build_finish_date { get; set; }
        public string? pre_building_material_note { get; set; }
        public string? pre_building_kind { get; set; }
        public string? pre_build_num { get; set; }
        public string? pre_build_area_m { get; set; }
        public string? pre_build_area_p { get; set; }
        public string? pre_public_area_m { get; set; }
        public string? pre_public_area_p { get; set; }
        public string? pre_use_kind { get; set; }
        public string? pre_build_area_total_p { get; set; }
        public string? pre_parking_kind { get; set; }
        public string? pre_parking_area_total_p { get; set; }
        public string? pre_principal_note { get; set; }
        public string? pre_pdf_path { get; set; }
        public string? pre_building_material { get; set; }
        public string? pre_storey_total { get; set; }
        public string? pre_community_name { get; set; }
        public string? pre_note { get; set; }
        public string? pre_process_type { get; set; }
        public string? pre_process_num { get; set; }
        public DateTime? pre_process_date { get; set; }
        public string? del_tag { get; set; }
        public DateTime? add_date { get; set; }
        public string? add_num { get; set; }
        public string? add_ip { get; set; }
        public DateTime? edit_date { get; set; }
        public string? edit_num { get; set; }
        public string? edit_ip { get; set; }
        public DateTime? del_date { get; set; }
        public string? del_num { get; set; }
        public string? del_ip { get; set; }
        public string? private_set { get; set; }
        public string? BuildingState { get; set; }
        public string? ParkCategory { get; set; }
        public string? MoiCityCode { get; set; }
        public string? MoiTownCode { get; set; }
        public string? MoiSectionCode { get; set; }
        public string? add_name { get; set; }
        public string? show_pre_building_kind { get; set; }
        public string? Account { get; set; }
        public string? BusinessUserName { get; set; }

        public string? HA_cknum { get; set; }
        public string? show_pre_parking_kind { get; set; }
        public string? CS_PID { get; set; }
        public string? U_BC { get; set; }
        /// <summary>
        /// 檢查資料是否正確
        /// </summary>
        /// <returns>null 正常;錯誤訊息</returns>
        public List<string> isRight()
        {
            List<string> errors = new List<string>();
            if(string.IsNullOrEmpty(pre_building_kind))
                errors.Add("建物類型不能為空");
            if (string.IsNullOrEmpty(pre_city))
                errors.Add("縣市代碼不能為空");
            if (string.IsNullOrEmpty(pre_area))
                errors.Add("鄉鎮市區不能為空");
            if (string.IsNullOrEmpty(pre_road))
                errors.Add("段不能為空");
            if(string.IsNullOrEmpty(pre_apply_name))
                errors.Add("申請人不能為空");
            if (string.IsNullOrEmpty(pre_build_num) && string.IsNullOrEmpty(pre_land_num))
                errors.Add("需建號/地號其中之一");
            return errors;
        }
    }
}
