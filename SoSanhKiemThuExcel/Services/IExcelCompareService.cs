using SoSanhKiemThuExcel.Models;
using System.IO;

namespace SoSanhKiemThuExcel.Services
{
    public interface IExcelCompareService
    {
        CompareSummaryResult Compare(
            Stream rootStream,
            string rootFileName,
            Stream compareStream,
            string compareFileName,
            int sheetRoot,
            int sheetCompare,
            int headerRowRoot,
            int headerRowCompare,
            int dataStartRowRoot,
            int dataStartRowCompare,
            int roundDigits,
            string outputFolder,
            int keyColRoot = 0,
            int keyColCompare = 0);
    }
}
