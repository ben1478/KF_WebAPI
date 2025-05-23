﻿using OfficeOpenXml.Packaging.Ionic.Zlib;

namespace KF_WebAPI.BaseClass.AE
{
    public class Procurement_Ins
    {
        public tbInfo? tbInfo { get; set; }
        public List<Procurement_D_Ins>? PD_Ins_List { get; set; }
        public string? PM_ID { get; set; }
        public string? PM_BC { get; set; }
        public string? PM_Pay_Type { get; set; }
        public string? PM_AppDate { get; set; }
        public string? PM_U_num { get; set; }
        public string? PM_Caption { get; set; }
        public decimal? PM_Amt { get; set; }
        public decimal? PM_Busin_Tax { get; set; }
        public decimal? PM_Tax_Amt { get; set; }
        public string? PM_Other { get; set; }
        public string? PM_Cancel { get; set; }
    }

    public class Procurement_D_Ins
    {
        public tbInfo? tbInfo { get; set; }
        public string? PM_ID { get; set; }
        public string? PD_Pro_name { get; set; }
        public string? PD_Unit { get; set; }
        public string? PD_Count { get; set; }
        public string? PD_Date { get; set; }
        public decimal? PD_Univalent { get; set; }
        public decimal? PD_Amt { get; set; }
        public string? PD_Company_name { get; set; }
        public decimal? PD_Est_Cost { get; set; }
    }

    public class Procurement_Res
    {
        public List<Procurement_D_Res> Procurement_D { get;set; }
        public List<Procurement_D_Diff> Procurement_Diff { get; set; }
        public string? PM_BC { get; set; }
        public string? PM_Pay_Type { get; set; }
        public string? PM_Caption { get; set; }
        public decimal? PM_Amt { get; set; }
        public decimal? PM_Busin_Tax { get; set; }
        public decimal? PM_Tax_Amt { get; set; }
        public string? PM_Other { get; set; }
        public string? PM_ID { get; set; }
    }

    public class Procurement_D_Res 
    {
        public int PD_ID { get; set; }
        public string? PD_Pro_name { get; set; }
        public string? PD_Unit { get; set; }
        public string? PD_Count { get; set; }
        public string? PD_Date { get; set; }
        public decimal? PD_Univalent { get; set; }
        public decimal? PD_Amt { get; set; }
        public string? PD_Company_name { get; set; }
        public decimal? PD_Est_Cost { get; set; }
    }

    public class Procurement_D_Diff
    {
        public string? PM_Step { get; set; }
        public string? PD_Pro_name { get; set; }
        public string? PD_Count { get; set; }
        public decimal? PD_Univalent { get; set; }
        public decimal? PD_Amt { get; set; }
        public decimal? PD_Est_Cost { get; set; }
    }

    public class ProcForm_M
    {
        public List<ProcForm_D> ProcFormDList { get; set; }
        public string? U_BC_Name { get; set; }
        public string? PM_ID { get; set; }
        public string? PM_Pay_Name { get; set; }
        public string? AppDate { get; set; }
        public string? PM_Caption { get; set;}
        public string? PM_Amt { get; set; }
        public string? PM_Busin_Tax { get; set; }
        public string? PM_Tax_Amt { get; set; }
        public string? bank_name { get; set; }
        public string? branch_name { get; set; }
        public string? bank_account { get; set; }
        public string? payee_name { get; set; }
        public string? PM_Other { get;set; }
        public string? PM_Step { get; set; }
    }

    public class ProcForm_D
    {
        public string? PD_Pro_name { get; set; }
        public string? PD_Unit { get; set; }
        public string? PD_Count { get;set; }
        public string? PD_Date { get;set; }
        public string? PD_Univalent { get;set; }
        public string? PD_Amt { get; set; }
        public string? PD_Company_name { get;set; }
        public string? PD_Est_Cost { get; set; }
    }

    public class Proc_Print 
    {
        public List<Proc_Print_Deatil>? PT_Deatil_List { get; set; }
        public List<Proc_Print_File>? PT_File_List { get; set; }
        public List<Proc_Print_Flow>? PT_Flow_List { get; set; }
        public string BC_Name { get; set; }
        public string U_name { get; set; }
        public string PM_AppDate { get; set; } 
        public string PM_ID { get; set; }
        public string PM_Caption { get; set;}
        public string? PM_Amt { get; set; }
        public string? PM_Busin_Tax { get; set; }
        public string? PM_Tax_Amt { get; set; }
        public string PM_cknum { get; set; }
        public string PM_Pay_Type { get; set; }
        public string? PM_Other { get; set; }
        public string? add_date { get; set; }
    }

    public class Proc_Print_Deatil
    {
        public string PD_Pro_name { get; set; }
        public string? PD_Unit { get; set; }
        public string? PD_Count { get; set; }
        public string? PD_Date { get; set; }
        public string? PD_Univalent { get; set; }
        public string? PD_Amt { get; set; }
        public string? PD_Company_name { get; set; }
        public string? PD_Est_Cost { get; set; } 
    }

    public class Proc_Print_File
    {
        public string upload_name_show { get; set; }
    }

    public class Proc_Print_Flow
    {
        public string FD_Sign_Countersign { get; set; }
        public string FD_Step { get; set; }
        public string U_name { get; set; }
        public string FD_Step_date { get; set; }
        public string? FD_Step_desc { get; set; }   
    }

    public class PT_Rpt_req
    {
        public string? Type { get; set; }
        public string Date_S { get; set; }
        public string Date_E { get; set; }
    }

    public class PT_Excel
    {
        public string Form_ID { get; set; }
        public string strType { get; set;}
        public string BC_Name { get; set; }
        public string U_Name { get; set; }
        public string? str_Name { get; set; }
        public string? str_Amt { get; set; }

    }
}
