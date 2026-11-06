using Microsoft.AspNetCore.Http;

namespace SoSanhKiemThuExcel.Models
{
    public class ExcelCompareRequest
    {
        public IFormFile? FileRoot { get; set; }
        public IFormFile? FileCompare { get; set; }
        public int SheetRoot { get; set; } = 1;
        public int SheetCompare { get; set; } = 1;
        public int HeaderRowRoot { get; set; } = 1;
        public int HeaderRowCompare { get; set; } = 1;
        public int DataStartRowRoot { get; set; } = 0;
        public int DataStartRowCompare { get; set; } = 0;
        public int RoundDigits { get; set; } = 2; // -1: Lấy hết (Không làm tròn), 0..10: Số chữ số làm tròn
    }
}
