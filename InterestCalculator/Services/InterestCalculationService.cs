using InterestCalculator.Models;

namespace InterestCalculator.Services
{
    /// <summary>
    /// Service for calculating loan interest and payment schedules
    /// </summary>
    public class InterestCalculationService
    {
        /// <summary>
        /// Calculate simple interest
        /// </summary>
        public decimal CalculateSimpleInterest(decimal principal, decimal annualRate, int months)
        {
            return principal * (annualRate / 100) * (months / 12m);
        }

        /// <summary>
        /// Calculate compound interest
        /// </summary>
        public decimal CalculateCompoundInterest(decimal principal, decimal annualRate, int months)
        {
            double rate = (double)(annualRate / 100 / 12);
            double periods = months;
            double amount = (double)principal * Math.Pow(1 + rate, periods);
            return (decimal)amount - principal;
        }

        /// <summary>
        /// Calculate monthly payment for a loan
        /// </summary>
        public decimal CalculateMonthlyPayment(decimal principal, decimal annualRate, int months)
        {
            if (months == 0) return 0;
            if (annualRate == 0) return principal / months;

            double monthlyRate = (double)(annualRate / 100 / 12);
            double periods = months;

            // PMT = P * [r(1+r)^n] / [(1+r)^n - 1]
            double payment = (double)principal *
                (monthlyRate * Math.Pow(1 + monthlyRate, periods)) /
                (Math.Pow(1 + monthlyRate, periods) - 1);

            return (decimal)payment;
        }

        /// <summary>
        /// Generate payment schedule
        /// </summary>
        public List<PaymentSchedule> GeneratePaymentSchedule(LoanRecord loan)
        {
            var schedule = new List<PaymentSchedule>();
            decimal remainingBalance = loan.PrincipalAmount;
            decimal monthlyPayment = CalculateMonthlyPayment(
                loan.PrincipalAmount,
                loan.InterestRate,
                loan.TermInMonths);

            for (int i = 1; i <= loan.TermInMonths; i++)
            {
                decimal monthlyRate = loan.InterestRate / 100 / 12;
                decimal interestPayment = remainingBalance * monthlyRate;
                decimal principalPayment = monthlyPayment - interestPayment;

                // Adjust last payment for rounding
                if (i == loan.TermInMonths)
                {
                    principalPayment = remainingBalance;
                    monthlyPayment = principalPayment + interestPayment;
                }

                remainingBalance -= principalPayment;

                schedule.Add(new PaymentSchedule
                {
                    PaymentNumber = i,
                    PaymentDate = loan.LoanDate.AddMonths(i),
                    PrincipalPayment = Math.Round(principalPayment, 2),
                    InterestPayment = Math.Round(interestPayment, 2),
                    TotalPayment = Math.Round(monthlyPayment, 2),
                    RemainingBalance = Math.Round(remainingBalance, 2)
                });
            }

            return schedule;
        }

        /// <summary>
        /// Calculate loan details
        /// </summary>
        public void CalculateLoanDetails(LoanRecord loan)
        {
            if (loan.LoanType == "Simple")
            {
                loan.CalculatedInterest = CalculateSimpleInterest(
                    loan.PrincipalAmount,
                    loan.InterestRate,
                    loan.TermInMonths);
            }
            else
            {
                loan.CalculatedInterest = CalculateCompoundInterest(
                    loan.PrincipalAmount,
                    loan.InterestRate,
                    loan.TermInMonths);
            }

            loan.TotalAmount = loan.PrincipalAmount + loan.CalculatedInterest;
            loan.MonthlyPayment = CalculateMonthlyPayment(
                loan.PrincipalAmount,
                loan.InterestRate,
                loan.TermInMonths);
        }
    }
}
