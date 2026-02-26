using System;
using System.Drawing.Printing;

namespace KF_WebAPI.BaseClass.AE
{
    public class ItemList
    {
        public tbInfo? tbInfo { get; set; }
        public decimal item_id { get; set; }
        public string? item_cknum { get; set;}
        public string? item_M_type { get; set; }
        public string? tem_M_code { get; set; }
        public string? tem_M_name { get; set; }
        public string? tem_D_type { get; set; }
        public string? tem_D_code { get; set; }
        public string? tem_D_name { get; set; }
        public string? tem_D_color { get; set; }
        public string? tem_D_txt_A { get; set; }
        public string? item_D_txt_B { get; set; }
        public int? item_D_int_A { get; set; }
        public int? tem_D_int_B { get; set; }
        public int? item_sort { get; set; }
        public string? del_tag { get; set; }
        public string? show_tag { get; set; }
        public DateTime? add_date { get; set; }
        public string? add_num { get; set; }
        public string? add_ip { get; set; }
        public DateTime? edit_date { get; set; }
        public string? edit_num { get; set; }
        public string? edit_ip { get; set; }
        public DateTime? del_date { get; set; }
        public string? del_num { get; set; }
        public string? del_ip { get; set; }
        public string? project_condition_enabled { get; set; }
    }
}
