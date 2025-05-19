using ProtoBuf;
using System.ComponentModel;

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

    }
}
