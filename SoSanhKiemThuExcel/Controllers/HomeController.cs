using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using SoSanhKiemThuExcel.Models;
using SoSanhKiemThuExcel.Services;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace SoSanhKiemThuExcel.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly IExcelCompareService _compareService;

        public HomeController(ILogger<HomeController> logger, IExcelCompareService compareService)
        {
            _logger = logger;
            _compareService = compareService;
        }

        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public IActionResult KiemTraExcel(
            IFormFile? fileRoot, 
            IFormFile? fileCompare, 
            int sheetRoot, 
            int sheetCompare,
            int headerRowRoot, 
            int headerRowCompare, 
            int dataStartRowRoot, 
            int dataStartRowCompare,
            int roundDigits)
        {
            try
            {
                // Kiểm tra file
                if (fileRoot == null || fileCompare == null)
                {
                    return BadRequest(new { message = "Vui lòng chọn đủ 2 file Excel." });
                }

                // Validate extension & size + magic bytes
                var validExtensions = new[] { ".xlsx", ".xlsm" };
                foreach (var file in new[] { fileRoot, fileCompare })
                {
                    if (file.Length > 50 * 1024 * 1024) // 50MB
                    {
                        return BadRequest(new { message = $"File '{file.FileName}' quá lớn (tối đa 50MB)." });
                    }

                    var ext = Path.GetExtension(file.FileName).ToLowerInvariant();
                    if (!validExtensions.Contains(ext))
                    {
                        return BadRequest(new { message = $"Chỉ chấp nhận các file Excel định dạng: {string.Join(", ", validExtensions)}." });
                    }

                    // Check magic bytes cho Excel hợp lệ
                    using var testStream = file.OpenReadStream();
                    var buffer = new byte[4];
                    var bytesRead = testStream.Read(buffer, 0, 4);
                    if (bytesRead < 4 || !IsValidExcelHeader(buffer, ext))
                    {
                        return BadRequest(new { message = $"File '{file.FileName}' không phải định dạng Excel hợp lệ." });
                    }
                }

                // Validation headerRow
                if (headerRowRoot < 1 || headerRowRoot > 20 || headerRowCompare < 1 || headerRowCompare > 20)
                {
                    return BadRequest(new { message = "Dòng tiêu đề phải từ 1 đến 20." });
                }

                // roundDigits: -1 là Lấy hết (Không làm tròn), 0..10 là số chữ số làm tròn
                if (roundDigits < -1 || roundDigits > 10)
                {
                    return BadRequest(new { message = "Số chữ số làm tròn không hợp lệ (-1 để lấy hết, hoặc từ 0 đến 10)." });
                }

                var uploadsFolder = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "files");

                using var rootStream = fileRoot.OpenReadStream();
                using var compareStream = fileCompare.OpenReadStream();

                var result = _compareService.Compare(
                    rootStream,
                    fileRoot.FileName,
                    compareStream,
                    fileCompare.FileName,
                    sheetRoot,
                    sheetCompare,
                    headerRowRoot,
                    headerRowCompare,
                    dataStartRowRoot,
                    dataStartRowCompare,
                    roundDigits,
                    uploadsFolder);

                return Json(result);
            }
            catch (ArgumentException ex)
            {
                return BadRequest(new { message = ex.Message });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Lỗi trong KiemTraExcel");
                return StatusCode(500, new { message = $"Lỗi xử lý file: {ex.Message}" });
            }
        }

        private bool IsValidExcelHeader(byte[] buffer, string extension)
        {
            if (extension == ".xlsx" || extension == ".xlsm")
            {
                return buffer[0] == 0x50 && buffer[1] == 0x4B && buffer[2] == 0x03 && buffer[3] == 0x04;
            }
            return false;
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
