using OfficeOpenXml;
using SoSanhKiemThuExcel.Services;
using System;
using System.Collections.Generic;
using System.IO;
using Xunit;

namespace SoSanhKiemThuExcel.Tests
{
    public class ExcelCompareServiceTests
    {
        [Fact]
        public void CompareCellValues_RoundDigitsMinus1_TakesAllAndDetectsDifference()
        {
            // Với roundDigits = -1 (Lấy hết để so luôn)
            // 12.3456789 vs 12.3456788 phải coi là KHÁC NHAU!
            bool isDiff = ExcelCompareService.CompareCellValues(
                "12.3456789", 
                "12.3456788", 
                -1, 
                out string rootDisplay, 
                out string compareDisplay);

            Assert.True(isDiff);
            Assert.Equal("12.3456789", rootDisplay);
            Assert.Equal("12.3456788", compareDisplay);
        }

        [Fact]
        public void CompareCellValues_RoundDigitsMinus1_EqualValuesMatch()
        {
            // Với roundDigits = -1, các giá trị số bằng nhau phải TRÙNG KHỚP
            bool isDiff = ExcelCompareService.CompareCellValues(
                "123.45678901", 
                "123.45678901", 
                -1, 
                out _, 
                out _);

            Assert.False(isDiff);
        }

        [Fact]
        public void CompareCellValues_RoundDigits2_RoundsCorrectly()
        {
            // Với roundDigits = 2: 12.345 vs 12.349 -> làm tròn thành 12.35 vs 12.35 -> TRÙNG NHAU
            bool isDiff = ExcelCompareService.CompareCellValues(
                12.345, 
                12.349, 
                2, 
                out _, 
                out _);

            Assert.False(isDiff);

            // 12.341 vs 12.349 -> 12.34 vs 12.35 -> KHÁC NHAU
            bool isDiff2 = ExcelCompareService.CompareCellValues(
                12.341, 
                12.349, 
                2, 
                out string rootDisplay, 
                out string compareDisplay);

            Assert.True(isDiff2);
            Assert.Equal("12.34", rootDisplay);
            Assert.Equal("12.35", compareDisplay);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("   ")]
        [InlineData("-")]
        [InlineData("--")]
        [InlineData("—")]
        [InlineData("0")]
        [InlineData("0.00")]
        [InlineData("0,00")]
        [InlineData("(0)")]
        [InlineData("(0.00)")]
        [InlineData("N/A")]
        public void IsZeroLike_RecognizesAllZeroRepresentations(object? value)
        {
            Assert.True(ExcelCompareService.IsZeroLike(value));
        }

        [Fact]
        public void CompareCellValues_ZeroLikePairs_Match()
        {
            // Empty string vs "0.00"
            bool isDiff1 = ExcelCompareService.CompareCellValues("", "0.00", 2, out _, out _);
            Assert.False(isDiff1);

            // "-" vs 0
            bool isDiff2 = ExcelCompareService.CompareCellValues("-", 0, -1, out _, out _);
            Assert.False(isDiff2);

            // "(0)" vs "0,00"
            bool isDiff3 = ExcelCompareService.CompareCellValues("(0)", "0,00", -1, out _, out _);
            Assert.False(isDiff3);
        }

        [Theory]
        [InlineData("1234,56", 1234.56)]
        [InlineData("1,234.56", 1234.56)]
        [InlineData("1.234,56", 1234.56)]
        [InlineData("(123.45)", -123.45)]
        [InlineData("100.000 ₫", 100000.0)]
        [InlineData("$500.50", 500.50)]
        public void TryParseNumber_HandlesCulturesAndFormats(string input, double expected)
        {
            bool ok = ExcelCompareService.TryParseNumber(input, out decimal decVal, out double dblVal);
            Assert.True(ok);
            Assert.Equal(expected, (double)decVal, 2);
        }

        [Fact]
        public void FindKeyColumnIndex_FindsVariousHeaderNames()
        {
            Assert.Equal(1, ExcelCompareService.FindKeyColumnIndex(new List<string> { "MASO", "TEN", "TIEN" }));
            Assert.Equal(2, ExcelCompareService.FindKeyColumnIndex(new List<string> { "STT", "Mã số", "Số tiền" }));
            Assert.Equal(1, ExcelCompareService.FindKeyColumnIndex(new List<string> { "MÃ_SỐ", "HỌ TÊN" }));
            Assert.Equal(3, ExcelCompareService.FindKeyColumnIndex(new List<string> { "COL1", "COL2", "Code" }));
            Assert.Equal(1, ExcelCompareService.FindKeyColumnIndex(new List<string> { "ID", "VALUE" }));
        }

        [Fact]
        public void EndToEnd_Compare_WithCustomKeyColumn_STT_MatchesCorrectly()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // File Gốc
            using var rootPkg = new ExcelPackage();
            var wsRoot = rootPkg.Workbook.Worksheets.Add("Sheet1");
            wsRoot.Cells["A1"].Value = "STT";
            wsRoot.Cells["B1"].Value = "HOTEN";
            wsRoot.Cells["C1"].Value = "DIEM";

            wsRoot.Cells["A2"].Value = 1;
            wsRoot.Cells["B2"].Value = "Nguyen Van A";
            wsRoot.Cells["C2"].Value = 8.5;

            wsRoot.Cells["A3"].Value = 2;
            wsRoot.Cells["B3"].Value = "Tran Thi B";
            wsRoot.Cells["C3"].Value = 9.0;

            var rootStream = new MemoryStream();
            rootPkg.SaveAs(rootStream);
            rootStream.Position = 0;

            // File So Sánh: Đảo thứ tự dòng (STT=2 trước, STT=1 sau)
            using var compPkg = new ExcelPackage();
            var wsComp = compPkg.Workbook.Worksheets.Add("Sheet1");
            wsComp.Cells["A1"].Value = "STT";
            wsComp.Cells["B1"].Value = "HOTEN";
            wsComp.Cells["C1"].Value = "DIEM";

            wsComp.Cells["A2"].Value = 2;
            wsComp.Cells["B2"].Value = "Tran Thi B";
            wsComp.Cells["C2"].Value = 9.0; // Khớp 100%

            wsComp.Cells["A3"].Value = 1;
            wsComp.Cells["B3"].Value = "Nguyen Van A";
            wsComp.Cells["C3"].Value = 7.0; // Khác điểm (8.5 vs 7.0)

            var compStream = new MemoryStream();
            compPkg.SaveAs(compStream);
            compStream.Position = 0;

            var tempOutDir = Path.Combine(Path.GetTempPath(), "ExcelCompareKeyColTest_" + Guid.NewGuid().ToString("N"));
            var service = new ExcelCompareService();

            try
            {
                // Chọn rõ cột khóa đối chiếu là Cột 1 (STT)
                var result = service.Compare(
                    rootStream, "Root.xlsx",
                    compStream, "Comp.xlsx",
                    1, 1, 1, 1, 2, 2, 2,
                    tempOutDir,
                    keyColRoot: 1,
                    keyColCompare: 1);

                Assert.Equal(2, result.TotalRowsCompared);
                Assert.Equal(1, result.TotalPassedRows); // STT=2 passed
                Assert.Equal(1, result.TotalFailedRows); // STT=1 failed vì điểm 8.5 != 7.0
                Assert.Equal("STT", result.KeyColumnName);

                // Lỗi ghi nhận mã số là "1" (giá trị của cột STT)
                Assert.Contains(result.TopErrors, e => e.MaSo == "1" && e.CotSai == "DIEM");
            }
            finally
            {
                if (Directory.Exists(tempOutDir))
                {
                    try { Directory.Delete(tempOutDir, true); } catch { }
                }
            }
        }

        [Fact]
        public void EndToEnd_Compare_WithTakeAll_DetectsDifferencesAccurately()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Tạo file gốc
            using var rootPkg = new ExcelPackage();
            var wsRoot = rootPkg.Workbook.Worksheets.Add("Data");
            wsRoot.Cells["A1"].Value = "MASO";
            wsRoot.Cells["B1"].Value = "SOTIEN";
            wsRoot.Cells["C1"].Value = "GHICHU";

            wsRoot.Cells["A2"].Value = "NV01";
            wsRoot.Cells["B2"].Value = 100.123456;
            wsRoot.Cells["C2"].Value = "Đạt";

            wsRoot.Cells["A3"].Value = "NV02";
            wsRoot.Cells["B3"].Value = 50.0;
            wsRoot.Cells["C3"].Value = "-";

            wsRoot.Cells["A4"].Value = "NV03";
            wsRoot.Cells["B4"].Value = 200.0;
            wsRoot.Cells["C4"].Value = "Gốc chỉ có";

            var rootStream = new MemoryStream();
            rootPkg.SaveAs(rootStream);
            rootStream.Position = 0;

            // Tạo file so sánh (với cột đổi thứ tự B và C)
            using var compPkg = new ExcelPackage();
            var wsComp = compPkg.Workbook.Worksheets.Add("Data");
            wsComp.Cells["A1"].Value = "MASO";
            wsComp.Cells["B1"].Value = "GHICHU"; // Đổi thứ tự
            wsComp.Cells["C1"].Value = "SOTIEN";

            wsComp.Cells["A2"].Value = "NV01";
            wsComp.Cells["B2"].Value = "Đạt";
            wsComp.Cells["C2"].Value = 100.123457; // Khác ở chữ số thập phân thứ 6!

            wsComp.Cells["A3"].Value = "NV02";
            wsComp.Cells["B3"].Value = "0.00"; // Zero-like so với "-"
            wsComp.Cells["C3"].Value = 50.00;

            wsComp.Cells["A4"].Value = "NV04";
            wsComp.Cells["B4"].Value = "So sánh chỉ có";
            wsComp.Cells["C4"].Value = 300.0;

            var compStream = new MemoryStream();
            compPkg.SaveAs(compStream);
            compStream.Position = 0;

            var tempOutDir = Path.Combine(Path.GetTempPath(), "ExcelCompareTests_" + Guid.NewGuid().ToString("N"));
            var service = new ExcelCompareService();

            try
            {
                // 1. So sánh với roundDigits = 2 (Làm tròn 2 chữ số) -> NV01 100.12 == 100.12
                var resRound2 = service.Compare(
                    rootStream, "Root.xlsx",
                    compStream, "Comp.xlsx",
                    1, 1, 1, 1, 2, 2, 2,
                    tempOutDir);

                Assert.Equal(1, resRound2.TotalMissingInCompare);
                Assert.Equal(1, resRound2.TotalMissingInRoot);
                Assert.Equal(2, resRound2.TotalPassedRows); // NV01 và NV02 passed

                // 2. So sánh với roundDigits = -1 (LẤY HẾT ĐỂ SO LUÔN)
                rootStream.Position = 0;
                compStream.Position = 0;

                var resTakeAll = service.Compare(
                    rootStream, "Root.xlsx",
                    compStream, "Comp.xlsx",
                    1, 1, 1, 1, 2, 2, -1,
                    tempOutDir);

                // Khi lấy hết: NV01 có 100.123456 != 100.123457 -> NV01 PHẢI BỊ BÁO SAI!
                Assert.Equal(1, resTakeAll.TotalPassedRows); // Chỉ còn NV02 passed
                Assert.True(resTakeAll.TotalDiffCells >= 1);
                Assert.Contains(resTakeAll.TopErrors, e => e.MaSo == "NV01" && e.CotSai == "SOTIEN");

                // 3. Kiểm tra file Excel kết quả được tạo ra có đủ 3 sheet và cấu trúc chuẩn
                var generatedFilePath = Path.Combine(tempOutDir, resTakeAll.FileName);
                Assert.True(File.Exists(generatedFilePath));

                using var resultPkg = new ExcelPackage(new FileInfo(generatedFilePath));
                Assert.Equal(3, resultPkg.Workbook.Worksheets.Count);
                Assert.Equal("TongQuan", resultPkg.Workbook.Worksheets[0].Name);
                Assert.Equal("KetQuaSoSanh", resultPkg.Workbook.Worksheets[1].Name);
                Assert.Equal("DanhSachSai", resultPkg.Workbook.Worksheets[2].Name);

                // Kiểm tra nội dung sheet TongQuan
                var wsOverview = resultPkg.Workbook.Worksheets["TongQuan"];
                Assert.NotNull(wsOverview);
                Assert.Contains("BÁO CÁO KẾT QUẢ SO SÁNH", wsOverview.Cells["A1"].Text);
                Assert.Equal("Lấy hết (Không làm tròn - So sánh chính xác)", wsOverview.Cells["B8"].Text);
            }
            finally
            {
                if (Directory.Exists(tempOutDir))
                {
                    try { Directory.Delete(tempOutDir, true); } catch { }
                }
            }
        }
    }
}
