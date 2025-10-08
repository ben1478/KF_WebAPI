using System;

namespace KF_WebAPI.BaseClass.AE
{
    public class Feat_M
    {
        public tbInfo? tbInfo { get; set; }
        public decimal? FR_id { get; set; }
        public string? FR_M_code { get; set; }
        public string? FR_M_name { get; set; }
        public string? U_BC { get; set; }
    }

    public class Feat_M_res: Feat_M
    {
        public List<Feat_D> feat_Ds { get; set; } = new List<Feat_D>();
    }

    public class Feat_D : Feat_M
    {
        public decimal FR_D_ratio_A { get; set; }
        public decimal FR_D_ratio_B { get; set; }
        public string FR_D_rate { get; set; }
        public string FR_D_discount { get; set; } 
        public string FR_D_replace { get; set; }    
    }

    public class Feat_M_req
    {
        public tbInfo? tbInfo { get; set; }
        public string FR_M_code { get; set; }
        public string FR_M_name { get; set; }
        public string U_BC { get; set; }
    }
}
