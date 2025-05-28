using ProtoBuf;
using System.ComponentModel;
using System.Text;

namespace LoanSky.Model
{
    /// <summary>
    /// 案件資料
    /// </summary>
    [ProtoContract()]
    public partial class OrderRealEstateRequest : IExtensible
    {
        private IExtension extensionObject;
        IExtension IExtensible.GetExtensionObject(bool createIfMissing) => Extensible.GetExtensionObject(ref extensionObject, createIfMissing);

        [ProtoMember(1, Name = @"account")]
        [DefaultValue("")]
        public string Account { get; set; } = "";

        [ProtoMember(2,Name = "businessUserName")]
        [DefaultValue("")]
        public string BusinessUserName { get; set; } = "";

        [ProtoMember(3,Name = "businessTel")]
        [DefaultValue("")]
        public string BusinessTel { get; set; } = "";

        [ProtoMember(4,Name = "businessFax")]
        [DefaultValue("")]
        public string BusinessFax { get; set; } = "";

        [ProtoMember(5, Name = "applicant")]
        [DefaultValue("")]
        public string Applicant { get; set; } = "";

        [ProtoMember(6,Name = "idNo")]
        [DefaultValue("")]
        public string IdNo { get; set; } = "";

        [ProtoMember(7, Name = @"condominium")]
        [DefaultValue("")]
        public string Condominium { get; set; } = "";

        [ProtoMember(8,Name = "buildingState")]
        [DefaultValue("")]
        public string BuildingState { get; set; } = "";

        [ProtoMember(9, Name = @"situation")]
        [DefaultValue("")]
        public string Situation { get; set; } = "";

        [ProtoMember(10,Name = "parkCategory")]
        [DefaultValue("")]
        public string ParkCategory { get; set; } = "";

        [ProtoMember(11, Name = @"note")]
        [DefaultValue("")]
        public string Note { get; set; } = "";

        [ProtoMember(12,Name = "moiCityCode")]
        [DefaultValue("")]
        public string MoiCityCode { get; set; } = "";

        [ProtoMember(13,Name = "moiTownCode")]
        [DefaultValue("")]
        public string MoiTownCode { get; set; } = "";

        [ProtoMember(14, Name = @"attachments")]
        public System.Collections.Generic.List<OrderRealEstateAttachmentRequest> Attachments { get; } = new System.Collections.Generic.List<OrderRealEstateAttachmentRequest>();

        [ProtoMember(15, Name = @"nos")]
        public System.Collections.Generic.List<OrderRealEstateNoRequest> Nos { get; } = new System.Collections.Generic.List<OrderRealEstateNoRequest>();

        public bool IsLoanSkyFieldsNull
        {
            get
            {
                if (
                    string.IsNullOrEmpty(BusinessUserName) ||// 經辦人名稱
                    string.IsNullOrEmpty(Applicant) ||      // 申請人
                    string.IsNullOrEmpty(BuildingState) || // 建物類型(請參照對照表)
                    //string.IsNullOrEmpty(ParkCategory) || //  車位型態(請參照對照表)
                    string.IsNullOrEmpty(MoiCityCode) || // 縣市代碼(請參照對照表)
                    string.IsNullOrEmpty(MoiTownCode) ||    // 鄉鎮市區代碼(請參照對照表)
                    string.IsNullOrEmpty(Nos.FirstOrDefault().MoiSectionCode)) // 段代碼(請參照對照表)
                {
                    return true;
                }
                return false;
            }
        }
        public List<string> IsRight()
        {
            List<string> errors = new List<string>();
            if (string.IsNullOrEmpty(BusinessUserName))
                errors.Add("經辦人名稱不能為空");
            if (string.IsNullOrEmpty(Applicant))
                errors.Add("申請人不能為空");
            if (string.IsNullOrEmpty(BuildingState))
                errors.Add("建物類型不能為空");
            if (string.IsNullOrEmpty(MoiCityCode))
                errors.Add("縣市代碼不能為空");
            if (string.IsNullOrEmpty(MoiTownCode))
                errors.Add("鄉鎮市區不能為空");

            if (string.IsNullOrEmpty(Nos.FirstOrDefault().MoiSectionCode))
                errors.Add("段代碼不能為空");
            return errors;
        }

    }
}
