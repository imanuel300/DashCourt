using Microsoft.AspNetCore.Mvc;
using DashCourtApi.Services;

namespace DashCourtApi.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class DataController : ControllerBase
    {
        private readonly ExcelDataService _excelDataService;

        public DataController(ExcelDataService excelDataService)
        {
            _excelDataService = excelDataService;
        }

        [HttpGet("cr")]
        public ActionResult<List<Models.CRModel>> GetCRData()
        {
            var data = _excelDataService.GetCRData();
            if (data == null || data.Count == 0)
            {
                return NotFound("CR data not found.");
            }
            return Ok(data);
        }

        [HttpGet("avgo")]
        public ActionResult<List<Models.AVGOModel>> GetAVGOData()
        {
            var data = _excelDataService.GetAVGOData();
            if (data == null || data.Count == 0)
            {
                return NotFound("AVGO data not found.");
            }
            return Ok(data);
        }

        [HttpGet("sit")]
        public ActionResult<List<Models.SITModel>> GetSITData()
        {
            var data = _excelDataService.GetSITData();
            if (data == null || data.Count == 0)
            {
                return NotFound("SIT data not found.");
            }
            return Ok(data);
        }

        [HttpGet("inv3")]
        public ActionResult<List<Models.Inv3Model>> GetInv3Data()
        {
            var data = _excelDataService.GetInv3Data();
            if (data == null || data.Count == 0)
            {
                return NotFound("Inv3 data not found.");
            }
            return Ok(data);
        }
    }
}
