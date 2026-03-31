using OfficeOpenXml;
using InterestCalculator.Models;

namespace InterestCalculator.Services
{
    /// <summary>
    /// Service for reading and writing Excel files
    /// </summary>
    public class ExcelService
    {
        public ExcelService()
        {
            // Set EPPlus license context
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        /// <summary>
        /// Read loan records from Excel file
        /// </summary>
        public List<LoanRecord> ReadLoansFromExcel(string filePath)
        {
            var loans = new List<LoanRecord>();

            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException($"File not found: {filePath}");
            }

            using var package = new ExcelPackage(new FileInfo(filePath));
            var worksheet = package.Workbook.Worksheets[0];

            if (worksheet == null)
            {
                throw new Exception("No worksheet found in Excel file");
            }

            int rowCount = worksheet.Dimension?.Rows ?? 0;

            // Skip header row, start from row 2
            for (int row = 2; row <= rowCount; row++)
            {
                try
                {
                    var loan = new LoanRecord
                    {
                        Id = GetCellValue<int>(worksheet, row, 1),
                        CustomerName = GetCellValue<string>(worksheet, row, 2) ?? "",
                        PrincipalAmount = GetCellValue<decimal>(worksheet, row, 3),
                        InterestRate = GetCellValue<decimal>(worksheet, row, 4),
                        LoanDate = GetCellValue<DateTime>(worksheet, row, 5),
                        TermInMonths = GetCellValue<int>(worksheet, row, 6),
                        LoanType = GetCellValue<string>(worksheet, row, 7) ?? "Simple",
                        Status = GetCellValue<string>(worksheet, row, 8) ?? "Active",
                        Notes = GetCellValue<string>(worksheet, row, 9)
                    };

                    if (loan.TermInMonths > 0)
                    {
                        loan.DueDate = loan.LoanDate.AddMonths(loan.TermInMonths);
                    }

                    loans.Add(loan);
                }
                catch (Exception ex)
                {
                    // Log error but continue processing other rows
                    Console.WriteLine($"Error reading row {row}: {ex.Message}");
                }
            }

            return loans;
        }

        /// <summary>
        /// Export loan records to Excel file
        /// </summary>
        public void ExportLoansToExcel(List<LoanRecord> loans, string filePath)
        {
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
            worksheet.Cells[1, 10].Value = "Tiền lãi";
            worksheet.Cells[1, 11].Value = "Tổng tiền";
            worksheet.Cells[1, 12].Value = "Trả hàng tháng";

            // Style header
            using (var range = worksheet.Cells[1, 1, 1, 12])
            {
                range.Style.Font.Bold = true;
                range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
            }

            // Add data
            for (int i = 0; i < loans.Count; i++)
            {
                var loan = loans[i];
                int row = i + 2;

                worksheet.Cells[row, 1].Value = loan.Id;
                worksheet.Cells[row, 2].Value = loan.CustomerName;
                worksheet.Cells[row, 3].Value = loan.PrincipalAmount;
                worksheet.Cells[row, 4].Value = loan.InterestRate;
                worksheet.Cells[row, 5].Value = loan.LoanDate.ToString("dd/MM/yyyy");
                worksheet.Cells[row, 6].Value = loan.TermInMonths;
                worksheet.Cells[row, 7].Value = loan.LoanType;
                worksheet.Cells[row, 8].Value = loan.Status;
                worksheet.Cells[row, 9].Value = loan.Notes;
                worksheet.Cells[row, 10].Value = loan.CalculatedInterest;
                worksheet.Cells[row, 11].Value = loan.TotalAmount;
                worksheet.Cells[row, 12].Value = loan.MonthlyPayment;

                // Format currency
                worksheet.Cells[row, 3].Style.Numberformat.Format = "#,##0";
                worksheet.Cells[row, 10].Style.Numberformat.Format = "#,##0";
                worksheet.Cells[row, 11].Style.Numberformat.Format = "#,##0";
                worksheet.Cells[row, 12].Style.Numberformat.Format = "#,##0";
            }

            // Auto-fit columns
            worksheet.Cells.AutoFitColumns();

            // Save file
            package.SaveAs(new FileInfo(filePath));
        }

        /// <summary>
        /// Export payment schedule to Excel
        /// </summary>
        public void ExportPaymentScheduleToExcel(LoanRecord loan, List<PaymentSchedule> schedule, string filePath)
        {
            using var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("Payment Schedule");

            // Add loan summary
            worksheet.Cells[1, 1].Value = "Thông tin khoản vay";
            worksheet.Cells[1, 1].Style.Font.Bold = true;
            worksheet.Cells[1, 1].Style.Font.Size = 14;

            worksheet.Cells[2, 1].Value = "Khách hàng:";
            worksheet.Cells[2, 2].Value = loan.CustomerName;
            worksheet.Cells[3, 1].Value = "Số tiền gốc:";
            worksheet.Cells[3, 2].Value = loan.PrincipalAmount;
            worksheet.Cells[3, 2].Style.Numberformat.Format = "#,##0";
            worksheet.Cells[4, 1].Value = "Lãi suất:";
            worksheet.Cells[4, 2].Value = $"{loan.InterestRate}%";
            worksheet.Cells[5, 1].Value = "Kỳ hạn:";
            worksheet.Cells[5, 2].Value = $"{loan.TermInMonths} tháng";

            // Add payment schedule headers
            int headerRow = 7;
            worksheet.Cells[headerRow, 1].Value = "Kỳ";
            worksheet.Cells[headerRow, 2].Value = "Ngày thanh toán";
            worksheet.Cells[headerRow, 3].Value = "Trả gốc";
            worksheet.Cells[headerRow, 4].Value = "Trả lãi";
            worksheet.Cells[headerRow, 5].Value = "Tổng thanh toán";
            worksheet.Cells[headerRow, 6].Value = "Số dư còn lại";

            // Style header
            using (var range = worksheet.Cells[headerRow, 1, headerRow, 6])
            {
                range.Style.Font.Bold = true;
                range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGreen);
            }

            // Add schedule data
            for (int i = 0; i < schedule.Count; i++)
            {
                var payment = schedule[i];
                int row = headerRow + 1 + i;

                worksheet.Cells[row, 1].Value = payment.PaymentNumber;
                worksheet.Cells[row, 2].Value = payment.PaymentDate.ToString("dd/MM/yyyy");
                worksheet.Cells[row, 3].Value = payment.PrincipalPayment;
                worksheet.Cells[row, 4].Value = payment.InterestPayment;
                worksheet.Cells[row, 5].Value = payment.TotalPayment;
                worksheet.Cells[row, 6].Value = payment.RemainingBalance;

                // Format currency
                worksheet.Cells[row, 3].Style.Numberformat.Format = "#,##0";
                worksheet.Cells[row, 4].Style.Numberformat.Format = "#,##0";
                worksheet.Cells[row, 5].Style.Numberformat.Format = "#,##0";
                worksheet.Cells[row, 6].Style.Numberformat.Format = "#,##0";
            }

            // Auto-fit columns
            worksheet.Cells.AutoFitColumns();

            // Save file
            package.SaveAs(new FileInfo(filePath));
        }

        private T GetCellValue<T>(ExcelWorksheet worksheet, int row, int col)
        {
            var cellValue = worksheet.Cells[row, col].Value;

            if (cellValue == null)
            {
                return default(T)!;
            }

            if (typeof(T) == typeof(DateTime))
            {
                if (cellValue is DateTime dt)
                    return (T)(object)dt;
                if (DateTime.TryParse(cellValue.ToString(), out DateTime parsed))
                    return (T)(object)parsed;
                return default(T)!;
            }

            if (typeof(T) == typeof(decimal))
            {
                if (cellValue is double d)
                    return (T)(object)(decimal)d;
                if (decimal.TryParse(cellValue.ToString(), out decimal parsed))
                    return (T)(object)parsed;
                return default(T)!;
            }

            if (typeof(T) == typeof(int))
            {
                if (cellValue is double d)
                    return (T)(object)(int)d;
                if (int.TryParse(cellValue.ToString(), out int parsed))
                    return (T)(object)parsed;
                return default(T)!;
            }

            return (T)Convert.ChangeType(cellValue, typeof(T));
        }
    }
}
