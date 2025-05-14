using ProtoBuf;
using System.ComponentModel;

namespace LoanSky.Model
{
    [ProtoContract()]
    public partial class OrderRealEstateReply : IExtensible
    {
        private IExtension extensionObject;
        IExtension IExtensible.GetExtensionObject(bool createIfMissing) => Extensible.GetExtensionObject(ref extensionObject, createIfMissing);

        [ProtoMember(1, Name = @"result")]
        public bool Result { get; set; }

        [ProtoMember(2, Name = @"message")]
        [DefaultValue("")]
        public string Message { get; set; } = "";

        [ProtoMember(3, DataFormat = DataFormat.WellKnown, Name = @"time")]
        public System.DateTime? Time { get; set; }
    }
}
