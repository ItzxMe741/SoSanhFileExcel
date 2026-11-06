using OfficeOpenXml;
using OfficeOpenXml.Style;
using SoSanhKiemThuExcel.Models;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace SoSanhKiemThuExcel.Services
{
    public class ExcelCompareService : IExcelCompareService
    {
        public CompareSummaryResult Compare(
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
            int keyColCompare = 0)
        {
            var stopwatch = Stopwatch.StartNew();

            // Đảm bảo license EPPlus
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using var packageRoot = new ExcelPackage(rootStream);
            using var packageCompare = new ExcelPackage(compareStream);

            if (sheetRoot < 1 || sheetRoot > packageRoot.Workbook.Worksheets.Count)
            {
                throw new ArgumentException($"Sheet {sheetRoot} không tồn tại trong file gốc.");
            }

            if (sheetCompare < 1 || sheetCompare > packageCompare.Workbook.Worksheets.Count)
            {
                throw new ArgumentException($"Sheet {sheetCompare} không tồn tại trong file so sánh.");
            }

            var worksheetRoot = packageRoot.Workbook.Worksheets[sheetRoot - 1];
            var worksheetCompare = packageCompare.Workbook.Worksheets[sheetCompare - 1];

            if (worksheetRoot.Dimension == null || worksheetRoot.Dimension.End.Row < headerRowRoot)
            {
                throw new ArgumentException("Sheet trong file gốc không có dữ liệu hoặc không đủ dòng cho dòng tiêu đề.");
            }

            if (worksheetCompare.Dimension == null || worksheetCompare.Dimension.End.Row < headerRowCompare)
            {
                throw new ArgumentException("Sheet trong file so sánh không có dữ liệu hoặc không đủ dòng cho dòng tiêu đề.");
            }

            int actualDataStartRoot = dataStartRowRoot > 0 ? dataStartRowRoot : headerRowRoot + 1;
            int actualDataStartCompare = dataStartRowCompare > 0 ? dataStartRowCompare : headerRowCompare + 1;

            if (actualDataStartRoot <= headerRowRoot || actualDataStartCompare <= headerRowCompare)
            {
                throw new ArgumentException("Dòng dữ liệu bắt đầu phải lớn hơn dòng tiêu đề.");
            }

            // Lấy danh sách cột
            int maxColsRoot = worksheetRoot.Dimension.End.Column;
            int maxColsCompare = worksheetCompare.Dimension.End.Column;

            var headerRootList = new List<string>();
            var headerCompareList = new List<string>();
            var headerRoot = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var headerCompare = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var commonColumns = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            for (int col = 1; col <= maxColsRoot; col++)
            {
                var colName = worksheetRoot.GetValue(headerRowRoot, col)?.ToString()?.Trim() ?? "";
                headerRootList.Add(colName);
                if (!string.IsNullOrEmpty(colName)) headerRoot.Add(colName);
            }

            for (int col = 1; col <= maxColsCompare; col++)
            {
                var colName = worksheetCompare.GetValue(headerRowCompare, col)?.ToString()?.Trim() ?? "";
                headerCompareList.Add(colName);
                if (!string.IsNullOrEmpty(colName)) headerCompare.Add(colName);
            }

            commonColumns.UnionWith(headerRoot.Intersect(headerCompare));

            if (!commonColumns.Any())
            {
                throw new ArgumentException("Không tìm thấy cột nào trùng tên giữa 2 sheet để so sánh.");
            }

            // Xác định cột định danh dùng để đối chiếu (Key Column)
            int maColRoot = (keyColRoot > 0 && keyColRoot <= maxColsRoot)
                ? keyColRoot
                : FindKeyColumnIndex(headerRootList);

            int maColCompare = (keyColCompare > 0 && keyColCompare <= maxColsCompare)
                ? keyColCompare
                : (keyColRoot > 0
                    ? headerCompareList.FindIndex(s => string.Equals(s?.Trim(), headerRootList[maColRoot - 1]?.Trim(), StringComparison.OrdinalIgnoreCase)) + 1
                    : FindKeyColumnIndex(headerCompareList));

            if (maColCompare < 1)
            {
                maColCompare = FindKeyColumnIndex(headerCompareList);
            }

            if (maColRoot < 1 || maColCompare < 1)
            {
                string missingIn = "";
                if (maColRoot < 1 && maColCompare < 1) missingIn = "cả hai file";
                else if (maColRoot < 1) missingIn = "File Gốc";
                else missingIn = "File So Sánh";

                throw new ArgumentException($"Không tìm thấy cột định danh dùng để đối chiếu ở {missingIn}. Vui lòng chọn cột đối chiếu ở danh sách.");
            }

            string keyColRootName = (maColRoot >= 1 && maColRoot <= headerRootList.Count && !string.IsNullOrWhiteSpace(headerRootList[maColRoot - 1]))
                ? headerRootList[maColRoot - 1]
                : $"Cột {maColRoot}";

            string keyColCompName = (maColCompare >= 1 && maColCompare <= headerCompareList.Count && !string.IsNullOrWhiteSpace(headerCompareList[maColCompare - 1]))
                ? headerCompareList[maColCompare - 1]
                : $"Cột {maColCompare}";

            // Map MASO / Cột Khóa -> Row
            var rootMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            int maxRowsRoot = worksheetRoot.Dimension.End.Row;
            for (int row = actualDataStartRoot; row <= maxRowsRoot; row++)
            {
                var maValue = worksheetRoot.GetValue(row, maColRoot)?.ToString()?.Trim();
                if (!string.IsNullOrEmpty(maValue) && !rootMap.ContainsKey(maValue))
                {
                    rootMap[maValue] = row;
                }
            }

            var compareMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            int maxRowsCompare = worksheetCompare.Dimension.End.Row;
            for (int row = actualDataStartCompare; row <= maxRowsCompare; row++)
            {
                var maValue = worksheetCompare.GetValue(row, maColCompare)?.ToString()?.Trim();
                if (!string.IsNullOrEmpty(maValue) && !compareMap.ContainsKey(maValue))
                {
                    compareMap[maValue] = row;
                }
            }

            // Khởi tạo file kết quả Excel
            using var packageResult = new ExcelPackage();
            var workbookResult = packageResult.Workbook;

            // Sheet 1: KetQuaSoSanh (Chi tiết)
            var worksheetResult = workbookResult.Worksheets.Add("KetQuaSoSanh");

            // Sheet 2: DanhSachSai (Danh sách lỗi)
            var worksheetDiff = workbookResult.Worksheets.Add("DanhSachSai");
            worksheetDiff.Cells[1, 1].Value = keyColRootName; // Tên cột khóa đối chiếu (VD: STT, MASO, ...)
            worksheetDiff.Cells[1, 2].Value = "CotSai";
            worksheetDiff.Cells[1, 3].Value = "GiaTriCu";
            worksheetDiff.Cells[1, 4].Value = "GiaTriMoi";
            worksheetDiff.Cells[1, 5].Value = "LoaiLoi";

            int diffRow = 2;
            bool hasAnyDifference = false;
            int totalPassedRows = 0;
            int totalFailedRows = 0;
            int totalMissingInCompare = 0;
            int totalMissingInRoot = 0;
            int totalDiffCells = 0;

            var topErrors = new List<CompareDiffItem>();

            // Copy toàn bộ dữ liệu từ root làm base
            for (int row = 1; row <= maxRowsRoot; row++)
            {
                for (int col = 1; col <= maxColsRoot; col++)
                {
                    worksheetResult.SetValue(row, col, worksheetRoot.GetValue(row, col));
                }
            }

            int resultCol = maxColsRoot + 1;
            worksheetResult.SetValue(headerRowRoot, resultCol, "Result");
            worksheetResult.Cells.Style.Font.Name = "Times New Roman";

            // So sánh từng dòng root
            foreach (var kvp in rootMap)
            {
                var maValue = kvp.Key;
                int rootRow = kvp.Value;

                if (!compareMap.ContainsKey(maValue))
                {
                    // Dòng có ở file gốc nhưng thiếu ở file so sánh
                    var rowRange = worksheetResult.Cells[rootRow, 1, rootRow, maxColsRoot];
                    rowRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    rowRange.Style.Fill.BackgroundColor.SetColor(Color.LightGoldenrodYellow);
                    worksheetResult.SetValue(rootRow, resultCol, "False - Thiếu ở file so sánh");

                    worksheetDiff.Cells[diffRow, 1].Value = maValue;
                    worksheetDiff.Cells[diffRow, 2].Value = "(Toàn dòng)";
                    worksheetDiff.Cells[diffRow, 3].Value = "Có ở file gốc";
                    worksheetDiff.Cells[diffRow, 4].Value = "Không tồn tại";
                    worksheetDiff.Cells[diffRow, 5].Value = "Thiếu ở file so sánh";

                    if (topErrors.Count < 20)
                    {
                        topErrors.Add(new CompareDiffItem
                        {
                            MaSo = maValue,
                            CotSai = "(Toàn dòng)",
                            GiaTriCu = "Có ở file gốc",
                            GiaTriMoi = "Không tồn tại",
                            LoaiLoi = "Thiếu ở file so sánh"
                        });
                    }

                    diffRow++;
                    hasAnyDifference = true;
                    totalFailedRows++;
                    totalMissingInCompare++;
                    continue;
                }

                int compareRow = compareMap[maValue];
                bool rowHasDifference = false;

                foreach (var commonColName in commonColumns)
                {
                    int rootCol = headerRootList.FindIndex(s => string.Equals(s, commonColName, StringComparison.OrdinalIgnoreCase)) + 1;
                    int compareCol = headerCompareList.FindIndex(s => string.Equals(s, commonColName, StringComparison.OrdinalIgnoreCase)) + 1;

                    if (rootCol < 1 || compareCol < 1) continue;

                    var rootVal = worksheetRoot.GetValue(rootRow, rootCol);
                    var compVal = worksheetCompare.GetValue(compareRow, compareCol);

                    bool isDifferent = CompareCellValues(
                        rootVal, 
                        compVal, 
                        roundDigits, 
                        out string rootDisplay, 
                        out string compareDisplay);

                    if (isDifferent)
                    {
                        string combined = string.IsNullOrEmpty(rootDisplay) ? $"({compareDisplay})" :
                                          string.IsNullOrEmpty(compareDisplay) ? rootDisplay :
                                          $"{rootDisplay} ({compareDisplay})";

                        worksheetResult.SetValue(rootRow, rootCol, combined);
                        worksheetResult.Cells[rootRow, rootCol].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheetResult.Cells[rootRow, rootCol].Style.Fill.BackgroundColor.SetColor(Color.LightPink);

                        // Log lỗi vào sheet DanhSachSai
                        worksheetDiff.Cells[diffRow, 1].Value = maValue;
                        worksheetDiff.Cells[diffRow, 2].Value = commonColName;
                        worksheetDiff.Cells[diffRow, 3].Value = rootDisplay;
                        worksheetDiff.Cells[diffRow, 4].Value = compareDisplay;
                        worksheetDiff.Cells[diffRow, 5].Value = "Khác giá trị";

                        if (topErrors.Count < 20)
                        {
                            topErrors.Add(new CompareDiffItem
                            {
                                MaSo = maValue,
                                CotSai = commonColName,
                                GiaTriCu = rootDisplay,
                                GiaTriMoi = compareDisplay,
                                LoaiLoi = "Khác giá trị"
                            });
                        }

                        diffRow++;
                        rowHasDifference = true;
                        hasAnyDifference = true;
                        totalDiffCells++;
                    }
                }

                if (rowHasDifference)
                {
                    worksheetResult.SetValue(rootRow, resultCol, "False");
                    totalFailedRows++;
                }
                else
                {
                    worksheetResult.SetValue(rootRow, resultCol, "True");
                    totalPassedRows++;
                }
            }

            // Thêm các dòng chỉ có ở file compare
            int currentLastRow = maxRowsRoot + 1;
            foreach (var kvp in compareMap)
            {
                var maValue = kvp.Key;
                if (rootMap.ContainsKey(maValue)) continue;

                int compareRow = kvp.Value;

                // Căn chuẩn theo cột của sheet root
                for (int col = 1; col <= maxColsRoot; col++)
                {
                    var rootColName = headerRootList[col - 1];
                    if (string.IsNullOrEmpty(rootColName)) continue;

                    int compCol = headerCompareList.FindIndex(s => string.Equals(s, rootColName, StringComparison.OrdinalIgnoreCase)) + 1;
                    if (compCol > 0)
                    {
                        worksheetResult.SetValue(currentLastRow, col, worksheetCompare.GetValue(compareRow, compCol));
                    }
                }

                var rowRange = worksheetResult.Cells[currentLastRow, 1, currentLastRow, maxColsRoot];
                rowRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                rowRange.Style.Fill.BackgroundColor.SetColor(Color.LightGoldenrodYellow);

                worksheetResult.SetValue(currentLastRow, resultCol, "False - Thiếu ở file gốc");

                // Log lỗi vào sheet DanhSachSai
                worksheetDiff.Cells[diffRow, 1].Value = maValue;
                worksheetDiff.Cells[diffRow, 2].Value = "(Toàn dòng)";
                worksheetDiff.Cells[diffRow, 3].Value = "Không tồn tại";
                worksheetDiff.Cells[diffRow, 4].Value = "Có ở file so sánh";
                worksheetDiff.Cells[diffRow, 5].Value = "Thiếu ở file gốc";

                if (topErrors.Count < 20)
                {
                    topErrors.Add(new CompareDiffItem
                    {
                        MaSo = maValue,
                        CotSai = "(Toàn dòng)",
                        GiaTriCu = "Không tồn tại",
                        GiaTriMoi = "Có ở file so sánh",
                        LoaiLoi = "Thiếu ở file gốc"
                    });
                }

                diffRow++;
                hasAnyDifference = true;
                totalFailedRows++;
                totalMissingInRoot++;
                currentLastRow++;
            }

            // Thống kê tổng quan
            int totalUniqueRecords = rootMap.Count + totalMissingInRoot;
            double matchRate = totalUniqueRecords > 0
                ? Math.Round((double)totalPassedRows / totalUniqueRecords * 100, 2)
                : 100.0;

            string roundDesc = roundDigits == -1
                ? "Lấy hết (Không làm tròn - So sánh chính xác)"
                : $"Làm tròn {roundDigits} chữ số thập phân";

            // Định dạng sheet KetQuaSoSanh
            worksheetResult.Cells.AutoFitColumns();
            worksheetResult.Cells[headerRowRoot, resultCol].Style.Font.Bold = true;
            worksheetResult.Cells[headerRowRoot, resultCol].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheetResult.Cells[headerRowRoot, resultCol].Style.Fill.BackgroundColor.SetColor(Color.LightSteelBlue);
            worksheetResult.View.FreezePanes(actualDataStartRoot, 1);
            if (worksheetResult.Dimension != null && worksheetResult.Dimension.End.Row >= headerRowRoot)
            {
                worksheetResult.Cells[headerRowRoot, 1, worksheetResult.Dimension.End.Row, resultCol].AutoFilter = true;
            }

            // Định dạng sheet DanhSachSai
            FormatDiffWorksheet(worksheetDiff, diffRow);

            // Tạo sheet Tổng quan (Dashboard) đặt lên đầu workbook
            CreateOverviewWorksheet(
                workbookResult,
                rootFileName,
                compareFileName,
                worksheetRoot.Name,
                worksheetCompare.Name,
                sheetRoot,
                sheetCompare,
                roundDesc,
                keyColRootName,
                keyColCompName,
                totalUniqueRecords,
                totalPassedRows,
                totalFailedRows,
                totalMissingInCompare,
                totalMissingInRoot,
                totalDiffCells,
                matchRate,
                hasAnyDifference);

            // Tự động dọn dẹp file cũ > 24 giờ
            CleanOldResultFiles(outputFolder, 24);

            // Lưu file kết quả
            Directory.CreateDirectory(outputFolder);
            var currentTime = DateTime.Now;
            var fileName = $"compare_result_{currentTime:ddMMyyyy_HHmmss}.xlsx";
            var filePath = Path.Combine(outputFolder, fileName);
            packageResult.SaveAs(new FileInfo(filePath));

            stopwatch.Stop();
            var processTime = (stopwatch.ElapsedMilliseconds / 1000.0).ToString("F1");

            return new CompareSummaryResult
            {
                FileName = fileName,
                ProcessTime = processTime,
                DownloadUrl = $"/files/{fileName}",
                HasAnyDifference = hasAnyDifference,
                TotalRowsRoot = rootMap.Count,
                TotalRowsCompare = compareMap.Count,
                TotalRowsCompared = totalUniqueRecords,
                TotalPassedRows = totalPassedRows,
                TotalFailedRows = totalFailedRows,
                TotalMissingInCompare = totalMissingInCompare,
                TotalMissingInRoot = totalMissingInRoot,
                TotalDiffCells = totalDiffCells,
                MatchRate = matchRate,
                RoundDigitsDescription = roundDesc,
                KeyColumnName = keyColRootName,
                KeyColumnDescription = $"{keyColRootName} ↔ {keyColCompName}",
                TopErrors = topErrors
            };
        }

        public static bool CompareCellValues(
            object? rootVal,
            object? compareVal,
            int roundDigits,
            out string rootDisplay,
            out string compareDisplay)
        {
            string rootStr = (rootVal?.ToString() ?? "").Trim();
            string compareStr = (compareVal?.ToString() ?? "").Trim();
            rootDisplay = rootStr;
            compareDisplay = compareStr;

            // 1. Cả 2 cùng là dạng zero-like (0, 0.00, "-", ô trống, ...)
            if (IsZeroLike(rootVal) && IsZeroLike(compareVal))
            {
                return false; // Trùng nhau
            }

            // 2. Cả 2 là số (numeric)
            bool isRootNum = TryParseNumber(rootVal, out decimal rootDec, out double rootDbl);
            bool isCompNum = TryParseNumber(compareVal, out decimal compDec, out double compDbl);

            if (isRootNum && isCompNum)
            {
                if (roundDigits == -1)
                {
                    // LẤY HẾT ĐỂ SO LUÔN: So sánh chính xác tuyệt đối không làm tròn
                    bool isEqual = (rootDec == compDec) || (Math.Abs(rootDbl - compDbl) < 1e-12);
                    if (isEqual)
                    {
                        return false;
                    }

                    // Khác nhau: hiển thị giá trị gốc đầy đủ không bị cắt bớt
                    rootDisplay = rootStr;
                    compareDisplay = compareStr;
                    return true;
                }
                else
                {
                    // Làm tròn theo số chữ số chỉ định (AwayFromZero chuẩn kiểm toán)
                    decimal rootRounded = Math.Round(rootDec, roundDigits, MidpointRounding.AwayFromZero);
                    decimal compRounded = Math.Round(compDec, roundDigits, MidpointRounding.AwayFromZero);

                    if (rootRounded == compRounded)
                    {
                        return false;
                    }

                    string format = "N" + roundDigits;
                    rootDisplay = rootRounded.ToString(format, CultureInfo.InvariantCulture);
                    compareDisplay = compRounded.ToString(format, CultureInfo.InvariantCulture);
                    return true;
                }
            }

            // 3. So sánh ngày tháng (DateTime)
            if (TryExtractDate(rootVal, rootStr, out DateTime dt1) && TryExtractDate(compareVal, compareStr, out DateTime dt2))
            {
                if (dt1 == dt2 || (dt1.Date == dt2.Date && dt1.TimeOfDay == TimeSpan.Zero && dt2.TimeOfDay == TimeSpan.Zero))
                {
                    return false;
                }
            }

            // 4. So sánh chuỗi (không phân biệt hoa thường và khoảng trắng)
            if (string.Equals(rootStr, compareStr, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            return true; // Khác nhau
        }

        public static bool IsZeroLike(object? val)
        {
            if (val == null) return true;

            if (val is double d && Math.Abs(d) < 1e-12) return true;
            if (val is float f && Math.Abs(f) < 1e-6f) return true;
            if (val is decimal m && m == 0m) return true;
            if (val is int i && i == 0) return true;
            if (val is long l && l == 0) return true;
            if (val is short s && s == 0) return true;
            if (val is byte b && b == 0) return true;

            string str = val.ToString()?.Trim() ?? "";
            if (string.IsNullOrWhiteSpace(str)) return true;

            string cleaned = str.Replace(" ", "").ToUpperInvariant();
            if (cleaned == "-" || cleaned == "--" || cleaned == "—" || cleaned == "–" ||
                cleaned == "N/A" || cleaned == "NA" || cleaned == "NULL" || cleaned == "NIL" || cleaned == "NONE")
            {
                return true;
            }

            string numClean = cleaned.Replace("(", "").Replace(")", "");
            if (numClean == "0" || numClean == "-0" || numClean == "+0")
            {
                return true;
            }

            if (decimal.TryParse(numClean, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal dec) && dec == 0m)
            {
                return true;
            }

            if (decimal.TryParse(numClean.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out decimal decComma) && decComma == 0m)
            {
                return true;
            }

            return false;
        }

        public static bool TryParseNumber(object? val, out decimal decVal, out double dblVal)
        {
            decVal = 0m;
            dblVal = 0.0;

            if (val == null) return false;

            if (val is decimal m)
            {
                decVal = m;
                dblVal = (double)m;
                return true;
            }
            if (val is double d)
            {
                dblVal = d;
                try { decVal = (decimal)d; } catch { decVal = 0m; }
                return true;
            }
            if (val is float f)
            {
                dblVal = f;
                try { decVal = (decimal)f; } catch { decVal = 0m; }
                return true;
            }
            if (val is int i)
            {
                decVal = i;
                dblVal = i;
                return true;
            }
            if (val is long l)
            {
                decVal = l;
                dblVal = l;
                return true;
            }
            if (val is short s)
            {
                decVal = s;
                dblVal = s;
                return true;
            }
            if (val is byte b)
            {
                decVal = b;
                dblVal = b;
                return true;
            }

            string str = val.ToString()?.Trim() ?? "";
            if (string.IsNullOrWhiteSpace(str)) return false;

            // Xử lý số âm kế toán trong ngoặc: "(123.45)" -> "-123.45"
            bool isNegative = false;
            string cleanStr = str.Trim();
            if (cleanStr.StartsWith("(") && cleanStr.EndsWith(")"))
            {
                isNegative = true;
                cleanStr = cleanStr.Substring(1, cleanStr.Length - 2).Trim();
            }

            bool hadVnCurrency = str.Contains("₫") || str.Contains("VND", StringComparison.OrdinalIgnoreCase);

            // Loại bỏ ký hiệu tiền tệ và khoảng trắng
            cleanStr = cleanStr.Replace(" ", "")
                               .Replace("₫", "")
                               .Replace("$", "")
                               .Replace("€", "")
                               .Replace("VND", "", StringComparison.OrdinalIgnoreCase)
                               .Replace("USD", "", StringComparison.OrdinalIgnoreCase);

            int dotCount = cleanStr.Count(c => c == '.');
            int commaCount = cleanStr.Count(c => c == ',');

            string normalized = cleanStr;

            if (hadVnCurrency && commaCount == 0 && dotCount > 0)
            {
                // Với tiền tệ VNĐ (₫, VND), dấu chấm là phân cách hàng nghìn (VD: 100.000 ₫)
                normalized = cleanStr.Replace(".", "");
            }
            else if (dotCount > 0 && commaCount > 0)
            {
                int lastDot = cleanStr.LastIndexOf('.');
                int lastComma = cleanStr.LastIndexOf(',');
                if (lastDot > lastComma)
                {
                    // US format: 1,234.56
                    normalized = cleanStr.Replace(",", "");
                }
                else
                {
                    // VN/EU format: 1.234,56
                    normalized = cleanStr.Replace(".", "").Replace(",", ".");
                }
            }
            else if (commaCount == 1 && dotCount == 0)
            {
                // Dấu phẩy thập phân: 1234,56
                normalized = cleanStr.Replace(",", ".");
            }
            else if (commaCount > 1 && dotCount == 0)
            {
                // Phân cách hàng nghìn bằng dấu phẩy: 1,234,567
                normalized = cleanStr.Replace(",", "");
            }
            else if (dotCount > 1 && commaCount == 0)
            {
                // Phân cách hàng nghìn bằng dấu chấm: 1.234.567
                normalized = cleanStr.Replace(".", "");
            }

            if (isNegative && !normalized.StartsWith("-"))
            {
                normalized = "-" + normalized;
            }

            if (decimal.TryParse(normalized, NumberStyles.Any, CultureInfo.InvariantCulture, out decVal))
            {
                dblVal = (double)decVal;
                return true;
            }

            if (double.TryParse(normalized, NumberStyles.Any, CultureInfo.InvariantCulture, out dblVal))
            {
                try { decVal = (decimal)dblVal; } catch { decVal = 0m; }
                return true;
            }

            return false;
        }

        private static bool TryExtractDate(object? val, string str, out DateTime dt)
        {
            if (val is DateTime d)
            {
                dt = d;
                return true;
            }

            if (!string.IsNullOrWhiteSpace(str))
            {
                if (DateTime.TryParse(str, CultureInfo.InvariantCulture, DateTimeStyles.None, out dt))
                    return true;

                if (DateTime.TryParse(str, new CultureInfo("vi-VN"), DateTimeStyles.None, out dt))
                    return true;
            }

            dt = DateTime.MinValue;
            return false;
        }

        public static int FindKeyColumnIndex(List<string> headers)
        {
            if (headers == null || headers.Count == 0) return -1;

            // 1. Cột MASO chính xác
            int idx = headers.FindIndex(s => string.Equals(s?.Trim(), "MASO", StringComparison.OrdinalIgnoreCase));
            if (idx >= 0) return idx + 1;

            // 2. Các biến thể phổ biến
            var commonNames = new[] { "MÃ SỐ", "MÃ_SỐ", "MA_SO", "MÃSO", "MASO_", "MÃ", "CODE", "ID", "STT", "KEY" };
            foreach (var name in commonNames)
            {
                idx = headers.FindIndex(s => string.Equals(s?.Trim(), name, StringComparison.OrdinalIgnoreCase));
                if (idx >= 0) return idx + 1;
            }

            // 3. Chuẩn hóa không dấu
            idx = headers.FindIndex(s =>
            {
                if (string.IsNullOrWhiteSpace(s)) return false;
                string norm = RemoveDiacritics(s.Trim()).Replace(" ", "").Replace("_", "").ToUpperInvariant();
                return norm == "MASO" || norm == "MA" || norm == "CODE" || norm == "ID";
            });
            if (idx >= 0) return idx + 1;

            return -1;
        }

        private static string RemoveDiacritics(string text)
        {
            var normalizedString = text.Normalize(NormalizationForm.FormD);
            var stringBuilder = new StringBuilder(capacity: normalizedString.Length);

            for (int i = 0; i < normalizedString.Length; i++)
            {
                char c = normalizedString[i];
                var unicodeCategory = CharUnicodeInfo.GetUnicodeCategory(c);
                if (unicodeCategory != UnicodeCategory.NonSpacingMark)
                {
                    stringBuilder.Append(c);
                }
            }

            return stringBuilder.ToString().Normalize(NormalizationForm.FormC).Replace("đ", "d").Replace("Đ", "D");
        }

        private static void FormatDiffWorksheet(ExcelWorksheet worksheetDiff, int diffRow)
        {
            int lastRow = diffRow > 2 ? diffRow - 1 : 2;

            if (diffRow == 2)
            {
                // Không có lỗi
                worksheetDiff.Cells[2, 1].Value = "HOÀN TOÀN TRÙNG KHỚP";
                worksheetDiff.Cells[2, 2].Value = "-";
                worksheetDiff.Cells[2, 3].Value = "-";
                worksheetDiff.Cells[2, 4].Value = "-";
                worksheetDiff.Cells[2, 5].Value = "Không có sai lệch nào giữa 2 file.";

                using var okRange = worksheetDiff.Cells[2, 1, 2, 5];
                okRange.Style.Font.Name = "Times New Roman";
                okRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                okRange.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(235, 247, 238));
                okRange.Style.Font.Color.SetColor(Color.DarkGreen);
                okRange.Style.Font.Bold = true;
            }

            worksheetDiff.Cells[1, 1, lastRow, 5].AutoFitColumns();
            worksheetDiff.Column(1).Width = Math.Max(worksheetDiff.Column(1).Width, 18); // MASO
            worksheetDiff.Column(2).Width = Math.Max(worksheetDiff.Column(2).Width, 22); // CotSai
            worksheetDiff.Column(3).Width = Math.Max(worksheetDiff.Column(3).Width, 28); // GiaTriCu
            worksheetDiff.Column(4).Width = Math.Max(worksheetDiff.Column(4).Width, 28); // GiaTriMoi
            worksheetDiff.Column(5).Width = Math.Max(worksheetDiff.Column(5).Width, 25); // LoaiLoi

            // Header style
            using (var headerRange = worksheetDiff.Cells[1, 1, 1, 5])
            {
                headerRange.Style.Font.Name = "Times New Roman";
                headerRange.Style.Font.Bold = true;
                headerRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                headerRange.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                headerRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                headerRange.Style.Fill.BackgroundColor.SetColor(Color.LightSteelBlue);

                headerRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                headerRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                headerRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                headerRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            }

            // Data styling
            if (diffRow > 2)
            {
                for (int r = 2; r <= lastRow; r++)
                {
                    using var rowRange = worksheetDiff.Cells[r, 1, r, 5];
                    rowRange.Style.Font.Name = "Times New Roman";
                    rowRange.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    if (r % 2 == 0)
                    {
                        rowRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        rowRange.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(245, 245, 245));
                    }

                    rowRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    rowRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    rowRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    rowRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                }
            }

            worksheetDiff.View.FreezePanes(2, 1);
            worksheetDiff.Cells[1, 1, lastRow, 5].AutoFilter = true;
        }

        private static void CreateOverviewWorksheet(
            ExcelWorkbook workbook,
            string rootFileName,
            string compareFileName,
            string sheetRootName,
            string sheetCompareName,
            int sheetRootIndex,
            int sheetCompareIndex,
            string roundDigitsDesc,
            string keyColRootName,
            string keyColCompName,
            int totalRecords,
            int totalPassed,
            int totalFailed,
            int totalMissingInCompare,
            int totalMissingInRoot,
            int totalDiffCells,
            double matchRate,
            bool hasAnyDifference)
        {
            var ws = workbook.Worksheets.Add("TongQuan");
            workbook.Worksheets.MoveToStart("TongQuan");

            // Header Banner
            ws.Cells["A1:E1"].Merge = true;
            ws.Cells["A1"].Value = "BÁO CÁO KẾT QUẢ SO SÁNH & KIỂM THỬ DỮ LIỆU EXCEL";
            ws.Cells["A1"].Style.Font.Name = "Times New Roman";
            ws.Cells["A1"].Style.Font.Size = 15;
            ws.Cells["A1"].Style.Font.Bold = true;
            ws.Cells["A1"].Style.Font.Color.SetColor(Color.Navy);
            ws.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["A1"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            ws.Row(1).Height = 35;

            // Section I: Thông tin kiểm thử
            ws.Cells["A3"].Value = "I. THÔNG TIN THỰC HIỆN KIỂM THỬ";
            ws.Cells["A3"].Style.Font.Name = "Times New Roman";
            ws.Cells["A3"].Style.Font.Bold = true;
            ws.Cells["A3"].Style.Font.Size = 12;

            var infoData = new (string Label, string Value)[]
            {
                ("Thời gian kiểm thử:", DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss")),
                ("File Gốc:", $"{rootFileName} (Sheet {sheetRootIndex}: {sheetRootName})"),
                ("File So Sánh:", $"{compareFileName} (Sheet {sheetCompareIndex}: {sheetCompareName})"),
                ("Cột khóa đối chiếu:", $"{keyColRootName} (File Gốc) ↔ {keyColCompName} (File So Sánh)"),
                ("Chế độ làm tròn số:", roundDigitsDesc),
                ("Đánh giá tổng thể:", hasAnyDifference ? "CÓ SAI LỆCH DỮ LIỆU (FAIL)" : "KHỚP HOÀN TOÀN 100% (PASS)")
            };

            for (int i = 0; i < infoData.Length; i++)
            {
                int r = 4 + i;
                ws.Cells[r, 1].Value = infoData[i].Label;
                ws.Cells[r, 1].Style.Font.Bold = true;
                ws.Cells[r, 1].Style.Font.Name = "Times New Roman";

                ws.Cells[r, 2].Value = infoData[i].Value;
                ws.Cells[r, 2].Style.Font.Name = "Times New Roman";

                if (i == infoData.Length - 1)
                {
                    ws.Cells[r, 2].Style.Font.Bold = true;
                    ws.Cells[r, 2].Style.Font.Color.SetColor(hasAnyDifference ? Color.Red : Color.DarkGreen);
                }
            }

            // Section II: Bảng thống kê
            ws.Cells["A10"].Value = "II. BẢNG TỔNG HỢP SỐ LIỆU KIỂM THỬ";
            ws.Cells["A10"].Style.Font.Name = "Times New Roman";
            ws.Cells["A10"].Style.Font.Bold = true;
            ws.Cells["A10"].Style.Font.Size = 12;

            ws.Cells["A11"].Value = "STT";
            ws.Cells["B11"].Value = "Chỉ số kiểm thử";
            ws.Cells["C11"].Value = "Số lượng";
            ws.Cells["D11"].Value = "Ghi chú";

            using (var tblHeader = ws.Cells["A11:D11"])
            {
                tblHeader.Style.Font.Name = "Times New Roman";
                tblHeader.Style.Font.Bold = true;
                tblHeader.Style.Fill.PatternType = ExcelFillStyle.Solid;
                tblHeader.Style.Fill.BackgroundColor.SetColor(Color.LightSteelBlue);
                tblHeader.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                tblHeader.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                tblHeader.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                tblHeader.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                tblHeader.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            }

            var tableData = new (string No, string Metric, object Count, string Note)[]
            {
                ("1", "Tổng số dòng (Mã số) so sánh", totalRecords, "Tổng bản ghi được đối chiếu"),
                ("2", "Số dòng trùng khớp tuyệt đối (True)", totalPassed, "Khớp toàn bộ các cột chung"),
                ("3", "Số dòng có sai khác (False)", totalFailed, "Có ít nhất 1 ô sai lệch hoặc thiếu"),
                ("4", "Số dòng chỉ có ở file gốc (Thiếu ở file so sánh)", totalMissingInCompare, "Bị thiếu trong file so sánh"),
                ("5", "Số dòng chỉ có ở file so sánh (Thiếu ở file gốc)", totalMissingInRoot, "Chỉ xuất hiện ở file so sánh"),
                ("6", "Tổng số ô dữ liệu bị sai lệch", totalDiffCells, "Chi tiết xem tại sheet 'DanhSachSai'"),
                ("7", "Tỷ lệ trùng khớp (Match Rate)", $"{matchRate:F2}%", matchRate >= 100 ? "Trùng khớp 100%" : "Cần rà soát các ô lệch")
            };

            for (int i = 0; i < tableData.Length; i++)
            {
                int r = 12 + i;
                ws.Cells[r, 1].Value = tableData[i].No;
                ws.Cells[r, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[r, 2].Value = tableData[i].Metric;
                ws.Cells[r, 3].Value = tableData[i].Count;
                ws.Cells[r, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                ws.Cells[r, 4].Value = tableData[i].Note;

                using var rowRange = ws.Cells[r, 1, r, 4];
                rowRange.Style.Font.Name = "Times New Roman";
                rowRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                rowRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                rowRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                rowRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                if (i % 2 == 1)
                {
                    rowRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    rowRange.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(248, 249, 250));
                }
            }

            ws.Cells["A1:D19"].AutoFitColumns();
            ws.Column(1).Width = Math.Max(ws.Column(1).Width, 10);
            ws.Column(2).Width = Math.Max(ws.Column(2).Width, 48);
            ws.Column(3).Width = Math.Max(ws.Column(3).Width, 18);
            ws.Column(4).Width = Math.Max(ws.Column(4).Width, 35);
        }

        public static void CleanOldResultFiles(string uploadsFolder, int olderThanHours = 24)
        {
            try
            {
                if (!Directory.Exists(uploadsFolder)) return;

                var threshold = DateTime.Now.AddHours(-olderThanHours);
                var dirInfo = new DirectoryInfo(uploadsFolder);
                foreach (var file in dirInfo.GetFiles("compare_result_*.xlsx"))
                {
                    if (file.CreationTime < threshold && file.LastWriteTime < threshold)
                    {
                        try { file.Delete(); } catch { }
                    }
                }
            }
            catch
            {
                // Bỏ qua lỗi dọn dẹp nếu có file đang mở
            }
        }
    }
}
