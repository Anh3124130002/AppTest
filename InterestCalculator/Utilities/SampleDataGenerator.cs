using OfficeOpenXml;

namespace InterestCalculator.Utilities
{
    /// <summary>
    /// Utility to create sample Excel template
    /// </summary>
    public class SampleDataGenerator
    {
        public static void CreateSampleExcel(string filePath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("Loans");

            // Add headers
            worksheet.Cells[1, 1].Value = "ID";
            worksheet.Cells[1, 2].Value = "Tên khách hàng";
            worksheet.Cells[1, 3].Value = "Số tiền gốc";
            worksheet.Cells[1, 4].Value = "Lãi suất (%)";
            worksheet.Cells[1, 5].Value = "Ngày vay";
            worksheet.Cells[1, 6].Value = "Kỳ hạn (tháng)";
            worksheet.Cells[1, 7].Value = "Loại lãi";
            worksheet.Cells[1, 8].Value = "Trạng thái";
            worksheet.Cells[1, 9].Value = "Ghi chú";

            // Style header
            using (var range = worksheet.Cells[1, 1, 1, 9])
            {
                range.Style.Font.Bold = true;
                range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
                range.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thick);
            }

            // Add sample data
            var sampleData = new[]
            {
                new { Id = 1, Name = "Nguyễn Văn A", Principal = 100000000m, Rate = 12m, Date = "01/01/2024", Term = 12, Type = "Simple", Status = "Active", Notes = "Vay mua nhà" },
                new { Id = 2, Name = "Trần Thị B", Principal = 50000000m, Rate = 10m, Date = "15/01/2024", Term = 24, Type = "Compound", Status = "Active", Notes = "Vay kinh doanh" },
                new { Id = 3, Name = "Lê Văn C", Principal = 200000000m, Rate = 8.5m, Date = "01/02/2024", Term = 36, Type = "Compound", Status = "Active", Notes = "Vay mua xe" },
                new { Id = 4, Name = "Phạm Thị D", Principal = 75000000m, Rate = 11m, Date = "10/02/2024", Term = 18, Type = "Simple", Status = "Active", Notes = "Vay tiêu dùng" },
                new { Id = 5, Name = "Hoàng Văn E", Principal = 150000000m, Rate = 9m, Date = "20/02/2024", Term = 24, Type = "Compound", Status = "Active", Notes = "Vay đầu tư" }
            };

            for (int i = 0; i < sampleData.Length; i++)
            {
                int row = i + 2;
                var data = sampleData[i];

                worksheet.Cells[row, 1].Value = data.Id;
                worksheet.Cells[row, 2].Value = data.Name;
                worksheet.Cells[row, 3].Value = data.Principal;
                worksheet.Cells[row, 4].Value = data.Rate;
                worksheet.Cells[row, 5].Value = data.Date;
                worksheet.Cells[row, 6].Value = data.Term;
                worksheet.Cells[row, 7].Value = data.Type;
                worksheet.Cells[row, 8].Value = data.Status;
                worksheet.Cells[row, 9].Value = data.Notes;

                // Format currency
                worksheet.Cells[row, 3].Style.Numberformat.Format = "#,##0";
            }

            // Auto-fit columns
            worksheet.Cells.AutoFitColumns();

            // Add instructions sheet
            var instructionsSheet = package.Workbook.Worksheets.Add("Hướng dẫn");
            instructionsSheet.Cells[1, 1].Value = "HƯỚNG DẪN SỬ DỤNG FILE EXCEL MẪU";
            instructionsSheet.Cells[1, 1].Style.Font.Bold = true;
            instructionsSheet.Cells[1, 1].Style.Font.Size = 14;

            int instrRow = 3;
            instructionsSheet.Cells[instrRow++, 1].Value = "Cấu trúc file Excel:";
            instructionsSheet.Cells[instrRow++, 1].Value = "1. ID: Số thứ tự khoản vay (số nguyên)";
            instructionsSheet.Cells[instrRow++, 1].Value = "2. Tên khách hàng: Tên người vay (văn bản)";
            instructionsSheet.Cells[instrRow++, 1].Value = "3. Số tiền gốc: Số tiền vay ban đầu (số)";
            instructionsSheet.Cells[instrRow++, 1].Value = "4. Lãi suất (%): Lãi suất hàng năm (số, ví dụ: 12 = 12%)";
            instructionsSheet.Cells[instrRow++, 1].Value = "5. Ngày vay: Ngày bắt đầu khoản vay (ngày/tháng/năm)";
            instructionsSheet.Cells[instrRow++, 1].Value = "6. Kỳ hạn (tháng): Số tháng vay (số nguyên)";
            instructionsSheet.Cells[instrRow++, 1].Value = "7. Loại lãi: Simple (lãi đơn) hoặc Compound (lãi kép)";
            instructionsSheet.Cells[instrRow++, 1].Value = "8. Trạng thái: Active, Paid, Overdue";
            instructionsSheet.Cells[instrRow++, 1].Value = "9. Ghi chú: Thông tin bổ sung (tùy chọn)";

            instrRow += 2;
            instructionsSheet.Cells[instrRow++, 1].Value = "Lưu ý:";
            instructionsSheet.Cells[instrRow++, 1].Value = "- Không xóa hoặc sửa tên cột (dòng 1)";
            instructionsSheet.Cells[instrRow++, 1].Value = "- Ngày vay theo định dạng: dd/MM/yyyy hoặc dd/MM/yy";
            instructionsSheet.Cells[instrRow++, 1].Value = "- Loại lãi chỉ nhập 'Simple' hoặc 'Compound'";
            instructionsSheet.Cells[instrRow++, 1].Value = "- Số tiền gốc và lãi suất phải là số dương";

            instructionsSheet.Cells.AutoFitColumns();

            // Save file
            package.SaveAs(new FileInfo(filePath));
        }
    }
}
