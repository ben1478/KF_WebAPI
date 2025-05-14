using ProtoBuf;
using System.ComponentModel;

namespace LoanSky.Model
{
    [ProtoContract()]
    public partial class OrderRealEstateAdapterRequest : IExtensible
    {
        private IExtension extensionObject;
        IExtension IExtensible.GetExtensionObject(bool createIfMissing) => Extensible.GetExtensionObject(ref extensionObject, createIfMissing);

        [ProtoMember(1, Name = "apiKey")]
        [DefaultValue("")]
        public string ApiKey { get; set; } = "";

        [ProtoMember(2,Name = "orderRealEstates")]
        public System.Collections.Generic.List<OrderRealEstateRequest> OrderRealEstates { get; } = new System.Collections.Generic.List<OrderRealEstateRequest>();
    }
}
