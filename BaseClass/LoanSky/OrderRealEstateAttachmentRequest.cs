using ProtoBuf;
using System.ComponentModel;

namespace LoanSky.Model
{
    [ProtoContract()]
    public partial class OrderRealEstateAttachmentRequest : IExtensible
    {
        private IExtension extensionObject;
        IExtension IExtensible.GetExtensionObject(bool createIfMissing) => Extensible.GetExtensionObject(ref extensionObject, createIfMissing);

        [ProtoMember(1, Name = @"orginalFileName")]
        [DefaultValue("")]
        public string OrginalFileName { get; set; } = "";

        [ProtoMember(2, Name = @"file")]
        public byte[] File { get; set; }
    }
}
