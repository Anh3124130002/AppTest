namespace InterestCalculator.Models
{
    /// <summary>
    /// Represents a loan transaction record
    /// </summary>
    public class LoanRecord
    {
        public int Id { get; set; }
        public string CustomerName { get; set; } = string.Empty;
        public decimal PrincipalAmount { get; set; }
        public decimal InterestRate { get; set; }
        public DateTime LoanDate { get; set; }
        public DateTime? DueDate { get; set; }
        public int TermInMonths { get; set; }
        public string LoanType { get; set; } = "Simple"; // Simple or Compound
        public string Status { get; set; } = "Active";
        public string? Notes { get; set; }

        // Calculated fields
        public decimal CalculatedInterest { get; set; }
        public decimal TotalAmount { get; set; }
        public decimal MonthlyPayment { get; set; }
    }

    /// <summary>
    /// Payment schedule entry
    /// </summary>
    public class PaymentSchedule
    {
        public int PaymentNumber { get; set; }
        public DateTime PaymentDate { get; set; }
        public decimal PrincipalPayment { get; set; }
        public decimal InterestPayment { get; set; }
        public decimal TotalPayment { get; set; }
        public decimal RemainingBalance { get; set; }
    }
}
