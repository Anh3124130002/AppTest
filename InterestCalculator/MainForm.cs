using InterestCalculator.Models;
using InterestCalculator.Services;

namespace InterestCalculator
{
    public partial class MainForm : Form
    {
        private readonly InterestCalculationService _calculationService;
        private readonly ExcelService _excelService;
        private List<LoanRecord> _loans;

        // UI Controls
        private TabControl tabControl;
        private TabPage tabCalculator;
        private TabPage tabLoanList;
        private TabPage tabSchedule;

        // Calculator Tab Controls
        private Label lblCustomerName;
        private TextBox txtCustomerName;
        private Label lblPrincipal;
        private TextBox txtPrincipal;
        private Label lblInterestRate;
        private TextBox txtInterestRate;
        private Label lblTermMonths;
        private TextBox txtTermMonths;
        private Label lblLoanDate;
        private DateTimePicker dtpLoanDate;
        private Label lblLoanType;
        private ComboBox cmbLoanType;
        private Button btnCalculate;
        private Button btnClear;
        private Button btnSave;
        private GroupBox grpResults;
        private Label lblResultInterest;
        private Label lblResultTotal;
        private Label lblResultMonthly;
        private TextBox txtResultInterest;
        private TextBox txtResultTotal;
        private TextBox txtResultMonthly;

        // Loan List Tab Controls
        private DataGridView dgvLoans;
        private Button btnImportExcel;
        private Button btnExportExcel;
        private Button btnRefresh;
        private Button btnViewSchedule;

        // Schedule Tab Controls
        private DataGridView dgvSchedule;
        private Label lblScheduleInfo;
        private Button btnExportSchedule;

        public MainForm()
        {
            InitializeComponent();
            _calculationService = new InterestCalculationService();
            _excelService = new ExcelService();
            _loans = new List<LoanRecord>();

            InitializeCustomComponents();
            SetupModernStyle();
        }

        private void InitializeCustomComponents()
        {
            // Main Tab Control
            tabControl = new TabControl
            {
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 10F)
            };

            tabCalculator = new TabPage("Tính toán");
            tabLoanList = new TabPage("Danh sách khoản vay");
            tabSchedule = new TabPage("Lịch thanh toán");

            tabControl.TabPages.Add(tabCalculator);
            tabControl.TabPages.Add(tabLoanList);
            tabControl.TabPages.Add(tabSchedule);

            this.Controls.Add(tabControl);

            // Initialize Calculator Tab
            InitializeCalculatorTab();

            // Initialize Loan List Tab
            InitializeLoanListTab();

            // Initialize Schedule Tab
            InitializeScheduleTab();
        }

        private void InitializeCalculatorTab()
        {
            var panel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(20)
            };

            // Input Fields
            int yPos = 20;
            int labelWidth = 150;
            int textBoxWidth = 300;
            int spacing = 40;

            lblCustomerName = CreateLabel("Tên khách hàng:", 20, yPos, labelWidth);
            txtCustomerName = CreateTextBox(180, yPos, textBoxWidth);
            panel.Controls.Add(lblCustomerName);
            panel.Controls.Add(txtCustomerName);
            yPos += spacing;

            lblPrincipal = CreateLabel("Số tiền gốc (VND):", 20, yPos, labelWidth);
            txtPrincipal = CreateTextBox(180, yPos, textBoxWidth);
            panel.Controls.Add(lblPrincipal);
            panel.Controls.Add(txtPrincipal);
            yPos += spacing;

            lblInterestRate = CreateLabel("Lãi suất (% năm):", 20, yPos, labelWidth);
            txtInterestRate = CreateTextBox(180, yPos, textBoxWidth);
            panel.Controls.Add(lblInterestRate);
            panel.Controls.Add(txtInterestRate);
            yPos += spacing;

            lblTermMonths = CreateLabel("Kỳ hạn (tháng):", 20, yPos, labelWidth);
            txtTermMonths = CreateTextBox(180, yPos, textBoxWidth);
            panel.Controls.Add(lblTermMonths);
            panel.Controls.Add(txtTermMonths);
            yPos += spacing;

            lblLoanDate = CreateLabel("Ngày vay:", 20, yPos, labelWidth);
            dtpLoanDate = new DateTimePicker
            {
                Location = new Point(180, yPos),
                Width = textBoxWidth,
                Format = DateTimePickerFormat.Short
            };
            panel.Controls.Add(lblLoanDate);
            panel.Controls.Add(dtpLoanDate);
            yPos += spacing;

            lblLoanType = CreateLabel("Loại lãi:", 20, yPos, labelWidth);
            cmbLoanType = new ComboBox
            {
                Location = new Point(180, yPos),
                Width = textBoxWidth,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cmbLoanType.Items.AddRange(new object[] { "Simple", "Compound" });
            cmbLoanType.SelectedIndex = 0;
            panel.Controls.Add(lblLoanType);
            panel.Controls.Add(cmbLoanType);
            yPos += spacing + 10;

            // Buttons
            btnCalculate = CreateButton("Tính toán", 180, yPos, 100, 35);
            btnCalculate.Click += BtnCalculate_Click;
            btnCalculate.BackColor = Color.FromArgb(0, 122, 204);
            btnCalculate.ForeColor = Color.White;
            panel.Controls.Add(btnCalculate);

            btnClear = CreateButton("Xóa", 290, yPos, 100, 35);
            btnClear.Click += BtnClear_Click;
            panel.Controls.Add(btnClear);

            btnSave = CreateButton("Lưu vào danh sách", 400, yPos, 140, 35);
            btnSave.Click += BtnSave_Click;
            btnSave.BackColor = Color.FromArgb(76, 175, 80);
            btnSave.ForeColor = Color.White;
            panel.Controls.Add(btnSave);
            yPos += spacing + 20;

            // Results Group Box
            grpResults = new GroupBox
            {
                Text = "Kết quả",
                Location = new Point(20, yPos),
                Size = new Size(520, 180),
                Font = new Font("Segoe UI", 10F, FontStyle.Bold)
            };

            int grpYPos = 35;
            lblResultInterest = CreateLabel("Tổng tiền lãi:", 20, grpYPos, labelWidth);
            txtResultInterest = CreateTextBox(180, grpYPos, textBoxWidth);
            txtResultInterest.ReadOnly = true;
            txtResultInterest.BackColor = Color.LightYellow;
            grpResults.Controls.Add(lblResultInterest);
            grpResults.Controls.Add(txtResultInterest);
            grpYPos += spacing;

            lblResultTotal = CreateLabel("Tổng tiền phải trả:", 20, grpYPos, labelWidth);
            txtResultTotal = CreateTextBox(180, grpYPos, textBoxWidth);
            txtResultTotal.ReadOnly = true;
            txtResultTotal.BackColor = Color.LightYellow;
            grpResults.Controls.Add(lblResultTotal);
            grpResults.Controls.Add(txtResultTotal);
            grpYPos += spacing;

            lblResultMonthly = CreateLabel("Trả hàng tháng:", 20, grpYPos, labelWidth);
            txtResultMonthly = CreateTextBox(180, grpYPos, textBoxWidth);
            txtResultMonthly.ReadOnly = true;
            txtResultMonthly.BackColor = Color.LightYellow;
            grpResults.Controls.Add(lblResultMonthly);
            grpResults.Controls.Add(txtResultMonthly);

            panel.Controls.Add(grpResults);

            tabCalculator.Controls.Add(panel);
        }

        private void InitializeLoanListTab()
        {
            var panel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(10)
            };

            // Buttons
            var buttonPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 50
            };

            btnImportExcel = CreateButton("Nhập từ Excel", 10, 10, 130, 30);
            btnImportExcel.Click += BtnImportExcel_Click;
            btnImportExcel.BackColor = Color.FromArgb(0, 122, 204);
            btnImportExcel.ForeColor = Color.White;
            buttonPanel.Controls.Add(btnImportExcel);

            btnExportExcel = CreateButton("Xuất ra Excel", 150, 10, 130, 30);
            btnExportExcel.Click += BtnExportExcel_Click;
            btnExportExcel.BackColor = Color.FromArgb(76, 175, 80);
            btnExportExcel.ForeColor = Color.White;
            buttonPanel.Controls.Add(btnExportExcel);

            btnRefresh = CreateButton("Làm mới", 290, 10, 100, 30);
            btnRefresh.Click += BtnRefresh_Click;
            buttonPanel.Controls.Add(btnRefresh);

            btnViewSchedule = CreateButton("Xem lịch thanh toán", 400, 10, 150, 30);
            btnViewSchedule.Click += BtnViewSchedule_Click;
            btnViewSchedule.BackColor = Color.FromArgb(255, 152, 0);
            btnViewSchedule.ForeColor = Color.White;
            buttonPanel.Controls.Add(btnViewSchedule);

            panel.Controls.Add(buttonPanel);

            // DataGridView
            dgvLoans = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                ReadOnly = true,
                AllowUserToAddRows = false,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                Font = new Font("Segoe UI", 9F)
            };

            dgvLoans.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 240, 240);

            panel.Controls.Add(dgvLoans);
            tabLoanList.Controls.Add(panel);
        }

        private void InitializeScheduleTab()
        {
            var panel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(10)
            };

            // Info label
            lblScheduleInfo = new Label
            {
                Dock = DockStyle.Top,
                Height = 50,
                Font = new Font("Segoe UI", 11F, FontStyle.Bold),
                Text = "Chọn một khoản vay từ danh sách để xem lịch thanh toán",
                TextAlign = ContentAlignment.MiddleCenter
            };
            panel.Controls.Add(lblScheduleInfo);

            // Export button
            var buttonPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 40
            };

            btnExportSchedule = CreateButton("Xuất lịch thanh toán", 10, 5, 180, 30);
            btnExportSchedule.Click += BtnExportSchedule_Click;
            btnExportSchedule.BackColor = Color.FromArgb(76, 175, 80);
            btnExportSchedule.ForeColor = Color.White;
            buttonPanel.Controls.Add(btnExportSchedule);

            panel.Controls.Add(buttonPanel);

            // DataGridView
            dgvSchedule = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                ReadOnly = true,
                AllowUserToAddRows = false,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                Font = new Font("Segoe UI", 9F)
            };

            dgvSchedule.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 240, 240);

            panel.Controls.Add(dgvSchedule);
            tabSchedule.Controls.Add(panel);
        }

        private Label CreateLabel(string text, int x, int y, int width)
        {
            return new Label
            {
                Text = text,
                Location = new Point(x, y),
                Width = width,
                Font = new Font("Segoe UI", 10F),
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleLeft
            };
        }

        private TextBox CreateTextBox(int x, int y, int width)
        {
            return new TextBox
            {
                Location = new Point(x, y),
                Width = width,
                Font = new Font("Segoe UI", 10F)
            };
        }

        private Button CreateButton(string text, int x, int y, int width, int height)
        {
            return new Button
            {
                Text = text,
                Location = new Point(x, y),
                Width = width,
                Height = height,
                Font = new Font("Segoe UI", 9F),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
        }

        private void SetupModernStyle()
        {
            this.BackColor = Color.White;
            this.Font = new Font("Segoe UI", 9F);
        }

        // Event Handlers
        private void BtnCalculate_Click(object? sender, EventArgs e)
        {
            try
            {
                if (!ValidateInput())
                    return;

                var loan = new LoanRecord
                {
                    CustomerName = txtCustomerName.Text,
                    PrincipalAmount = decimal.Parse(txtPrincipal.Text),
                    InterestRate = decimal.Parse(txtInterestRate.Text),
                    TermInMonths = int.Parse(txtTermMonths.Text),
                    LoanDate = dtpLoanDate.Value,
                    LoanType = cmbLoanType.SelectedItem?.ToString() ?? "Simple"
                };

                _calculationService.CalculateLoanDetails(loan);

                txtResultInterest.Text = loan.CalculatedInterest.ToString("#,##0");
                txtResultTotal.Text = loan.TotalAmount.ToString("#,##0");
                txtResultMonthly.Text = loan.MonthlyPayment.ToString("#,##0");

                MessageBox.Show("Tính toán thành công!", "Thông báo",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi tính toán: {ex.Message}", "Lỗi",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnClear_Click(object? sender, EventArgs e)
        {
            txtCustomerName.Clear();
            txtPrincipal.Clear();
            txtInterestRate.Clear();
            txtTermMonths.Clear();
            dtpLoanDate.Value = DateTime.Now;
            cmbLoanType.SelectedIndex = 0;
            txtResultInterest.Clear();
            txtResultTotal.Clear();
            txtResultMonthly.Clear();
        }

        private void BtnSave_Click(object? sender, EventArgs e)
        {
            try
            {
                if (!ValidateInput())
                    return;

                var loan = new LoanRecord
                {
                    Id = _loans.Count + 1,
                    CustomerName = txtCustomerName.Text,
                    PrincipalAmount = decimal.Parse(txtPrincipal.Text),
                    InterestRate = decimal.Parse(txtInterestRate.Text),
                    TermInMonths = int.Parse(txtTermMonths.Text),
                    LoanDate = dtpLoanDate.Value,
                    LoanType = cmbLoanType.SelectedItem?.ToString() ?? "Simple",
                    Status = "Active"
                };

                _calculationService.CalculateLoanDetails(loan);
                _loans.Add(loan);

                RefreshLoanList();

                MessageBox.Show("Đã lưu khoản vay vào danh sách!", "Thông báo",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);

                BtnClear_Click(sender, e);
                tabControl.SelectedTab = tabLoanList;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi lưu: {ex.Message}", "Lỗi",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnImportExcel_Click(object? sender, EventArgs e)
        {
            try
            {
                using var openFileDialog = new OpenFileDialog
                {
                    Filter = "Excel Files|*.xlsx;*.xls",
                    Title = "Chọn file Excel"
                };

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    _loans = _excelService.ReadLoansFromExcel(openFileDialog.FileName);

                    // Calculate details for each loan
                    foreach (var loan in _loans)
                    {
                        _calculationService.CalculateLoanDetails(loan);
                    }

                    RefreshLoanList();

                    MessageBox.Show($"Đã nhập {_loans.Count} khoản vay từ Excel!", "Thông báo",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi nhập Excel: {ex.Message}", "Lỗi",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnExportExcel_Click(object? sender, EventArgs e)
        {
            try
            {
                if (_loans.Count == 0)
                {
                    MessageBox.Show("Không có dữ liệu để xuất!", "Thông báo",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                using var saveFileDialog = new SaveFileDialog
                {
                    Filter = "Excel Files|*.xlsx",
                    Title = "Lưu file Excel",
                    FileName = $"DanhSachVay_{DateTime.Now:yyyyMMdd}.xlsx"
                };

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    _excelService.ExportLoansToExcel(_loans, saveFileDialog.FileName);

                    MessageBox.Show("Đã xuất dữ liệu ra Excel thành công!", "Thông báo",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi xuất Excel: {ex.Message}", "Lỗi",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnRefresh_Click(object? sender, EventArgs e)
        {
            RefreshLoanList();
        }

        private void BtnViewSchedule_Click(object? sender, EventArgs e)
        {
            try
            {
                if (dgvLoans.SelectedRows.Count == 0)
                {
                    MessageBox.Show("Vui lòng chọn một khoản vay!", "Thông báo",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var selectedLoan = (LoanRecord)dgvLoans.SelectedRows[0].DataBoundItem;
                var schedule = _calculationService.GeneratePaymentSchedule(selectedLoan);

                dgvSchedule.DataSource = schedule;
                lblScheduleInfo.Text = $"Lịch thanh toán - {selectedLoan.CustomerName} - " +
                    $"Số tiền: {selectedLoan.PrincipalAmount:#,##0} VND";

                tabControl.SelectedTab = tabSchedule;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi xem lịch thanh toán: {ex.Message}", "Lỗi",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnExportSchedule_Click(object? sender, EventArgs e)
        {
            try
            {
                if (dgvLoans.SelectedRows.Count == 0)
                {
                    MessageBox.Show("Vui lòng chọn một khoản vay từ danh sách!", "Thông báo",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var selectedLoan = (LoanRecord)dgvLoans.SelectedRows[0].DataBoundItem;
                var schedule = _calculationService.GeneratePaymentSchedule(selectedLoan);

                using var saveFileDialog = new SaveFileDialog
                {
                    Filter = "Excel Files|*.xlsx",
                    Title = "Lưu lịch thanh toán",
                    FileName = $"LichThanhToan_{selectedLoan.CustomerName}_{DateTime.Now:yyyyMMdd}.xlsx"
                };

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    _excelService.ExportPaymentScheduleToExcel(selectedLoan, schedule, saveFileDialog.FileName);

                    MessageBox.Show("Đã xuất lịch thanh toán ra Excel!", "Thông báo",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi xuất lịch thanh toán: {ex.Message}", "Lỗi",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void RefreshLoanList()
        {
            dgvLoans.DataSource = null;
            dgvLoans.DataSource = _loans;

            if (dgvLoans.Columns.Count > 0)
            {
                dgvLoans.Columns["Id"].HeaderText = "ID";
                dgvLoans.Columns["CustomerName"].HeaderText = "Khách hàng";
                dgvLoans.Columns["PrincipalAmount"].HeaderText = "Tiền gốc";
                dgvLoans.Columns["PrincipalAmount"].DefaultCellStyle.Format = "#,##0";
                dgvLoans.Columns["InterestRate"].HeaderText = "Lãi suất (%)";
                dgvLoans.Columns["LoanDate"].HeaderText = "Ngày vay";
                dgvLoans.Columns["LoanDate"].DefaultCellStyle.Format = "dd/MM/yyyy";
                dgvLoans.Columns["TermInMonths"].HeaderText = "Kỳ hạn";
                dgvLoans.Columns["LoanType"].HeaderText = "Loại lãi";
                dgvLoans.Columns["Status"].HeaderText = "Trạng thái";
                dgvLoans.Columns["CalculatedInterest"].HeaderText = "Tiền lãi";
                dgvLoans.Columns["CalculatedInterest"].DefaultCellStyle.Format = "#,##0";
                dgvLoans.Columns["TotalAmount"].HeaderText = "Tổng tiền";
                dgvLoans.Columns["TotalAmount"].DefaultCellStyle.Format = "#,##0";
                dgvLoans.Columns["MonthlyPayment"].HeaderText = "Trả/tháng";
                dgvLoans.Columns["MonthlyPayment"].DefaultCellStyle.Format = "#,##0";

                // Hide unnecessary columns
                if (dgvLoans.Columns["DueDate"] != null)
                    dgvLoans.Columns["DueDate"].Visible = false;
                if (dgvLoans.Columns["Notes"] != null)
                    dgvLoans.Columns["Notes"].Visible = false;
            }
        }

        private bool ValidateInput()
        {
            if (string.IsNullOrWhiteSpace(txtCustomerName.Text))
            {
                MessageBox.Show("Vui lòng nhập tên khách hàng!", "Lỗi nhập liệu",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtCustomerName.Focus();
                return false;
            }

            if (!decimal.TryParse(txtPrincipal.Text, out decimal principal) || principal <= 0)
            {
                MessageBox.Show("Số tiền gốc phải là số dương!", "Lỗi nhập liệu",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtPrincipal.Focus();
                return false;
            }

            if (!decimal.TryParse(txtInterestRate.Text, out decimal rate) || rate < 0)
            {
                MessageBox.Show("Lãi suất phải là số không âm!", "Lỗi nhập liệu",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtInterestRate.Focus();
                return false;
            }

            if (!int.TryParse(txtTermMonths.Text, out int months) || months <= 0)
            {
                MessageBox.Show("Kỳ hạn phải là số nguyên dương!", "Lỗi nhập liệu",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtTermMonths.Focus();
                return false;
            }

            return true;
        }
    }
}
