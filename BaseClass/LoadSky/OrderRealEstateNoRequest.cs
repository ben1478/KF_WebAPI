using ProtoBuf;
using System.ComponentModel;

namespace LoanSky.Model
{
    [ProtoContract()]
    public partial class OrderRealEstateNoRequest : IExtensible
    {
        private IExtension extensionObject;
        IExtension IExtensible.GetExtensionObject(bool createIfMissing) => Extensible.GetExtensionObject(ref extensionObject, createIfMissing);

        [ProtoMember(1, Name = "moiSectionCode")]
        [DefaultValue("")]
        public string MoiSectionCode { get; set; }

        [ProtoMember(2, Name = "buildNos")]
        [DefaultValue("")]
        public string BuildNos { get; set; } = "";

        [ProtoMember(3, Name = "landNos")]
        [DefaultValue("")]
        public string LandNos { get; set; } = "";

    }
}
