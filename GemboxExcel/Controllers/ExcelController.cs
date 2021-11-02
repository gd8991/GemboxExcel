using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using GemboxExcel.Helpers;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;

namespace GemboxExcel.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelController : ControllerBase
    {
        private readonly ILogger<ExcelController> _logger;

        public ExcelController(ILogger<ExcelController> logger)
        {
            _logger = logger;
        }

        public FileContentResult GetExcel(int noOfRows = 100, string format = "xlsx")
        {
            try
            {
                var message = $"Generating {format} file ...";
                _logger.LogInformation(message);
                var path = GemboxHelper.GenerateFile(noOfRows, format);
                message = $"Generated {format} file at {path} ...";
                _logger.LogInformation(message);
                var bytes = System.IO.File.ReadAllBytes(path);
                return File(bytes, $"application/{format}", path);
            }
            catch (Exception e)
            {
                _logger.LogError(e.Message);
                return null;
            }
        }
    }
}
