using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using KF_WebAPI.BaseClass;
using KF_WebAPI.Service;
using KF_WebAPI.BaseClass.Bop;

namespace KF_WebAPI.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class VirtualAccountController : ControllerBase
    {
        private readonly VirtualAccountService _vaService;

        public VirtualAccountController(VirtualAccountService vaService)
        {
            _vaService = vaService;
        }

        [HttpPost("RegisterVirtualAccount")]
        public async Task<IActionResult> Register([FromBody] CreateVirtualAccountRequest request)
        {
           
            if (request.Amount <= 0)
                return BadRequest("金額必須大於 0");

            var result = await _vaService.CreateVirtualAccountAsync(request);

            if (!result.Success)
                return StatusCode(500, result);

            return Ok(result);
        }
    }
}
