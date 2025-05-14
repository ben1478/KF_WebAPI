using LoanSky.Model;
using System.Threading.Tasks;

namespace LoanSky
{
    [System.ServiceModel.ServiceContract(Name = @"company.Company")]
    public interface ICompany
    {
        ValueTask<OrderRealEstateReply> CreateOrderRealEstateAsync(OrderRealEstateAdapterRequest value, ProtoBuf.Grpc.CallContext context = default);
    }
}
