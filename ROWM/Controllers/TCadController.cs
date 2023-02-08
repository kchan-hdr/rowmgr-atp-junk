using Microsoft.AspNetCore.Mvc;
using ROWM.Dal;
using System;
using System.Threading.Tasks;

namespace ROWM.Controllers
{
    [Route("api/v2")]
    [ApiController]
    public class TCadController : ControllerBase
    {
        [HttpGet("parcels/x/{id}/TCad_Owner")]
        public async Task<ActionResult<OwnerDto2>> GetOwner([FromServices] ROWM_Context context, string id)
        {
            var owners = await context.Database.SqlQuery<string>("SELECT [PartyName] FROM Austin.TCAD_OWNER WHERE [Tracking_Number] = @pid", 
                    new System.Data.SqlClient.SqlParameter("pid", id)
                ).FirstOrDefaultAsync();
            if (owners == null)
                return NotFound();

            return Ok(owners);
        }
    }
}
