using System.Collections.Generic;

namespace SoSanhKiemThuExcel.Models
{
    public class CompareDiffItem
    {
        public string MaSo { get; set; } = "";
        public string CotSai { get; set; } = "";
        public string GiaTriCu { get; set; } = "";
        public string GiaTriMoi { get; set; } = "";
        public string LoaiLoi { get; set; } = "";
    }

    public class CompareSummaryResult
    {
        public string FileName { get; set; } = "";
        public string ProcessTime { get; set; } = "";
        public string DownloadUrl { get; set; } = "";
        public bool HasAnyDifference { get; set; }
        public int TotalRowsRoot { get; set; }
        public int TotalRowsCompare { get; set; }
        public int TotalRowsCompared { get; set; }
        public int TotalPassedRows { get; set; }
        public int TotalFailedRows { get; set; }
        public int TotalMissingInCompare { get; set; }
        public int TotalMissingInRoot { get; set; }
        public int TotalDiffCells { get; set; }
        public double MatchRate { get; set; }
        public string RoundDigitsDescription { get; set; } = "";
        public string KeyColumnName { get; set; } = "MASO";
        public string KeyColumnDescription { get; set; } = "";
        public List<CompareDiffItem> TopErrors { get; set; } = new();
    }
}
