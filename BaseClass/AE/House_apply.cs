namespace KF_WebAPI.BaseClass.AE
{
    public class House_apply
    {
        public tbInfo? tbInfo { get; set; }
        public string CS_name { get; set; }
        public string? CS_sex { get; set; }
        public string CS_PID { get; set; }
        public string? CS_birth_year { get; set; }
        public string? CS_MTEL1 { get; set; }
        public string? CS_MTEL2 { get; set; }
        public string? CS_introducer { get; set; }
        public string? CS_introducer_PID { get; set; }
        public string? CS_note { get; set; }
        public string? plan_num { get; set; }
        public string? plan_date { get; set; }
        public string? CS_birthday { get; set; }
        public string? CS_register_address { get; set; }
        public string? CS_car { get; set; }
        public string? CS_carBrand { get; set; }
        public string? CS_carNumber { get; set; }
        public string? CS_carModel { get; set; }
        public string? CS_carManufacture { get; set; }
        public string? CS_carDisplacement { get; set; }
        public string? CS_EngineNo { get; set; }
        public string? CS_company_name { get; set; }
        public string? CS_company_address { get; set; }
        public string? CS_company_tel { get; set; }
        public string? CS_job_kind { get; set; }
        public string? CS_job_title { get; set; }
        public string? CS_job_years { get; set; }
        public string? CS_income_way { get; set; }
        public string? CS_rental { get; set; }
        public string? CS_income_everymonth { get; set; }
        public string? CS_license { get; set; }
        public string? CS_EMAIL { get; set; }
    }

    public class HouseApplyCar_req : House_apply
    {
        public string? project_title { get; set; }
        public string? pre_address { get; set; }
        public int project_apply_amount { get; set; }
        public string? HouseApplyChk { get; set; }  
    }
}
