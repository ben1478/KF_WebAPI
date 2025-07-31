namespace KF_WebAPI.BaseClass.AE
{
    public class User_M
    {
        public tbInfo? tbInfo { get; set; }
        public string? U_id { get; set; }
        public string? U_cknum { get; set; }
        public string? U_type { get; set; }
        public string? U_num { get; set; }
        public string? U_name { get; set; }
        public string? U_psw { get; set; }
        public string? U_sex { get; set; }
        public string? U_BC { get; set; }
        public string? U_PFT { get; set; }
        public string? U_agent_num { get; set; }
        public string? U_leader_1_num { get; set; }
        public string? U_leader_2_num { get; set; }
        public string? U_leader_3_num { get; set; }
        public string? U_address_live { get; set; }
        public DateTime? U_arrive_date { get; set; }
        public DateTime? U_leave_date { get; set; }
        public string? U_Check_BC { get; set; }
        public string? job { get; set; }
        public string? Tel { get; set; }
        public string? Role_num { get; set; }
        public string? U_Ename { get; set; }
        public DateTime? U_Birthday { get; set; }
        public string? Marriage { get; set; }
        public int? Children { get; set; }
        public string? U_PID { get; set; }
        public string? Military { get; set; }
        public string? Military_SDate { get; set; }
        public string? Military_EDate { get; set; }
        public string? Military_Exemption { get; set; }
        public string? License_Car { get; set; }
        public string? Self_Car { get; set; }
        public string? License_Motorcycle { get; set; }
        public string? Self_Motorcycle { get; set; }
        public string? U_Tel { get; set; }
        public string? U_MTel { get; set; }
        public string? U_Email { get; set; }
        public string? Emergency_contact { get; set; }
        public string? Emergency_Tel { get; set; }
        public string? Emergency_MTel { get; set; }
        public string? School_Level { get; set; }
        public string? School_Name { get; set; }
        public string? School_SDate { get; set; }
        public string? School_EDate { get; set; }
        public string? School_Graduated { get; set; }
        public string? School_D_N { get; set; }
        public string? School_Major { get; set; }

    }

    public class User_M_Ins
    {
        public tbInfo? tbInfo { get; set; }
        public string? U_num { get; set; }
        public string? U_name { get; set; }
        public string? U_Ename { get; set; }
        public string? str_U_Birthday { get; set; }
        public string? U_sex { get; set; }
        public string? Marriage { get; set; }
        public int? Children { get; set; }
        public string? U_PID { get; set; }
        public string? Military { get; set; }
        public string? Military_SDate { get; set; }
        public string? Military_EDate { get; set; }
        public string? Military_Exemption { get; set; }
        public string? License_Car { get; set; }
        public string? Self_Car { get; set; }
        public string? License_Motorcycle { get; set; }
        public string? Self_Motorcycle { get; set; }
        public string? U_Tel { get; set; }
        public string? U_MTel { get; set; }
        public string? U_Email { get; set; }
        public string? Emergency_contact { get; set; }
        public string? Emergency_Tel { get; set; }
        public string? Emergency_MTel { get; set; }
        public string? School_Level { get; set; }
        public string? School_Name { get; set; }
        public string? School_SDate { get; set; }
        public string? School_EDate { get; set; }
        public string? School_Graduated { get; set; }
        public string? School_D_N { get; set; }
        public string? School_Major { get; set; }
        public string? U_BC { get; set; }
        public string? U_PFT { get; set; }
        public string? Role_num { get; set; }
        public string? U_agent_num { get; set; }
        public string? U_leader_1_num { get; set; }
        public string? U_leader_2_num { get; set; }
        public string? U_address_live { get; set; }
        public string? str_U_arrive_date { get; set; }
        public string? str_U_leave_date { get; set; }
        public string User { get; set; }
    }

    public class User_M_Upd: User_M_Ins
    {
        public string U_id { get; set; }
        public string? U_Check_BC { get; set; }
    }
    public class Uesr_M_req
    {
        public string? U_Num_Name { get; set; }
        public string? U_BC { get;set; }
        public string? Job_Status { get;set; }
        public string? U_Role { get;set; }
        public string? Marriage { get; set; }
        public int? Children { get; set; }
        public string? U_PID { get; set; }
        public string? ID_104 { get; set; }
    }

    public class User_M_res
    {
        public decimal? ID_104 { get; set; }
        public string? U_BC_name { get; set; }
        public string? U_PFT_name { get; set; }
        public string? R_name { get; set; }
        public string? U_num { get; set; }
        public string? U_name { get; set; }
        public string? U_agent_name { get; set; }
        public string? U_leader_1_name { get; set; }
        public string? U_leader_2_name { get; set; }
        public string? del_tag { get; set; }
        public string? cknum { get; set; }
        public string? U_Check_BC { get; set; }
    }
}
