﻿namespace KF_WebAPI.BaseClass.LoanSky
{
    public class LoanSky_Req
    {
        public string HA_cknum { get; set; } // 
        public string HP_id { get; set; } // 
        public string HP_cknum { get; set; } //
        public string seq { get; set; } //

        public string p_USER { get; set; } //client 用戶員編
        public string LoanSkyApiKey { get; set; }
        public string LoanSkyUrl { get; set; } // LoanSky ApiKey
        public string LoanSkyAccount { get; set; } // LoanSky 帳號

    }
}
