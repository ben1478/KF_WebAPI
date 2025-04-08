using System;

namespace KF_WebAPI.BaseClass.AE
{
    public class House_agency
    {
        public int? AG_id { get; set; }
        public string? AG_cknum { get; set; }
        public int? HS_id { get; set; }
        public string? case_com { get; set; }
        public string? agency_com { get; set; }
        public string? case_text { get; set; }
        public string? CS_text { get; set; }
        public string? check_date { get; set; }
        public string? check_address { get; set; }
        public string? pass_amount { get; set; }
        public string? set_amount { get; set; }
        public string? print_data { get; set; }
        public string? get_data { get; set; }
        public string? process_charge { get; set; }
        public string? AG_note { get; set; }
        public string? check_leader_num { get; set; }
        public string? check_process_num { get; set; }
        public DateTime? set_process_date { get; set; }
        public string? check_process_type { get; set; }
        public DateTime? check_process_date { get; set; }
        public string? check_process_note { get; set; }
        public string? close_type { get; set; }
        public DateTime? close_type_date { get; set; }
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



    }

    public class House_agency_Ins
    {
        public int? AG_id { get; set; }
       
        public string? AG_cknum { get; set; }
        public int? HS_id { get; set; }
        public string? case_com { get; set; }
        public string? agency_com { get; set; }
        public string? case_text { get; set; }
        public string? CS_text { get; set; }
        public string? check_date { get; set; }
        public string? check_address { get; set; }
        public string? pass_amount { get; set; }
        public string? set_amount { get; set; }
        public string? print_data { get; set; }
        public string? get_data { get; set; }
        public string? process_charge { get; set; }
        public string? AG_note { get; set; }
        public string? check_leader_num { get; set; }
        public string? check_process_num { get; set; }
        public string? set_process_date { get; set; }
        public string? check_process_type { get; set; }
        public string? check_process_date { get; set; }
        public string? check_process_note { get; set; }
        public string? close_type { get; set; }
        public string? close_type_date { get; set; }
        public string? del_tag { get; set; }
        public string? add_num { get; set; }
        public string? edit_num { get; set; }
        public string? del_num { get; set; }
     

    }

    public class House_agency_Req
    {
        public int? AG_id { get; set; }
        public string Date_S { get; set; }
        public string Date_E { get; set; }
    }
    
}
