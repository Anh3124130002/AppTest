# Chương Trình Tính Lãi - Interest Calculator

## Mô tả (Description)

Ứng dụng Windows Forms hiện đại để tính toán lãi suất và quản lý các khoản vay. Chương trình hỗ trợ:
- Tính lãi đơn (Simple Interest) và lãi kép (Compound Interest)
- Nhập/xuất dữ liệu từ file Excel
- Tạo lịch thanh toán chi tiết
- Giao diện thân thiện, chuyên nghiệp và dễ sử dụng

A modern Windows Forms application for calculating interest and managing loans. The program supports:
- Simple and Compound Interest calculations
- Import/Export data from Excel files
- Generate detailed payment schedules
- Professional, user-friendly interface

## Tính năng (Features)

### 1. Tính toán lãi suất
- Hỗ trợ lãi đơn và lãi kép
- Tính toán tự động: tổng tiền lãi, tổng tiền phải trả, trả hàng tháng
- Xác thực dữ liệu đầu vào
- Hiển thị kết quả rõ ràng

### 2. Quản lý danh sách khoản vay
- Lưu trữ nhiều khoản vay
- Xem danh sách đầy đủ với bảng dữ liệu
- Tìm kiếm và sắp xếp

### 3. Nhập/Xuất Excel
- Nhập dữ liệu từ file Excel mẫu
- Xuất danh sách khoản vay ra Excel
- Xuất lịch thanh toán chi tiết

### 4. Lịch thanh toán
- Xem chi tiết từng kỳ thanh toán
- Hiển thị: tiền gốc, tiền lãi, tổng thanh toán, số dư còn lại
- Xuất lịch ra file Excel

## Yêu cầu hệ thống (System Requirements)

- Windows 10 hoặc cao hơn (Windows 10 or higher)
- .NET 8.0 Runtime
- Microsoft Excel (để xem file xuất ra)

## Cài đặt và Chạy (Installation & Running)

### Biên dịch từ mã nguồn (Build from source):

```bash
# Clone repository
git clone https://github.com/Anh3124130002/AppTest.git
cd AppTest

# Build project
dotnet build InterestCalculator.sln

# Run application
dotnet run --project InterestCalculator/InterestCalculator.csproj
```

### Chạy trực tiếp (Run directly):

```bash
cd InterestCalculator
dotnet run
```

## Hướng dẫn sử dụng (User Guide)

### 1. Tab "Tính toán" (Calculator Tab)

1. Nhập thông tin khoản vay:
   - Tên khách hàng
   - Số tiền gốc (VND)
   - Lãi suất (% năm)
   - Kỳ hạn (tháng)
   - Ngày vay
   - Loại lãi (Simple/Compound)

2. Nhấn "Tính toán" để xem kết quả

3. Nhấn "Lưu vào danh sách" để lưu khoản vay

### 2. Tab "Danh sách khoản vay" (Loan List Tab)

- **Nhập từ Excel**: Chọn file Excel mẫu để nhập dữ liệu
- **Xuất ra Excel**: Xuất danh sách hiện tại ra file Excel
- **Xem lịch thanh toán**: Chọn một khoản vay và xem lịch thanh toán

### 3. Tab "Lịch thanh toán" (Payment Schedule Tab)

- Xem chi tiết từng kỳ thanh toán
- Xuất lịch thanh toán ra Excel

## Cấu trúc File Excel (Excel File Structure)

File Excel cần có cấu trúc như sau:

| ID | Tên khách hàng | Số tiền gốc | Lãi suất (%) | Ngày vay | Kỳ hạn (tháng) | Loại lãi | Trạng thái | Ghi chú |
|----|----------------|-------------|--------------|----------|----------------|----------|------------|---------|
| 1  | Nguyễn Văn A   | 100000000   | 12           | 01/01/2024 | 12           | Simple   | Active     | Vay mua nhà |

### Chú ý:
- Cột tiêu đề phải giữ nguyên
- Ngày vay: dd/MM/yyyy
- Loại lãi: "Simple" hoặc "Compound"
- Trạng thái: "Active", "Paid", "Overdue"

## Công thức tính toán (Calculation Formulas)

### Lãi đơn (Simple Interest):
```
I = P × r × t
Trong đó:
- I: Tiền lãi (Interest)
- P: Tiền gốc (Principal)
- r: Lãi suất năm (Annual rate) / 100
- t: Thời gian (Time in years) = months / 12
```

### Lãi kép (Compound Interest):
```
A = P × (1 + r/n)^(n×t)
I = A - P
Trong đó:
- A: Tổng tiền (Total amount)
- P: Tiền gốc (Principal)
- r: Lãi suất năm (Annual rate) / 100
- n: Số kỳ tính lãi mỗi năm (12 cho tháng)
- t: Thời gian (Time in years)
- I: Tiền lãi (Interest)
```

### Trả hàng tháng (Monthly Payment):
```
PMT = P × [r(1+r)^n] / [(1+r)^n - 1]
Trong đó:
- PMT: Số tiền trả hàng tháng
- P: Tiền gốc
- r: Lãi suất tháng = (annual_rate / 100) / 12
- n: Số tháng
```

## Cấu trúc dự án (Project Structure)

```
InterestCalculator/
├── Models/
│   └── LoanRecord.cs           # Mô hình dữ liệu khoản vay
├── Services/
│   ├── InterestCalculationService.cs  # Dịch vụ tính toán lãi
│   └── ExcelService.cs         # Dịch vụ xử lý Excel
├── Utilities/
│   └── SampleDataGenerator.cs  # Tạo file Excel mẫu
├── MainForm.cs                 # Form chính
├── MainForm.Designer.cs        # Designer form
├── Program.cs                  # Entry point
└── InterestCalculator.csproj   # Project file
```

## Thư viện sử dụng (Dependencies)

- **EPPlus 7.0.0**: Xử lý file Excel (.xlsx)
- **.NET 8.0**: Framework
- **Windows Forms**: Giao diện người dùng

## Ví dụ sử dụng (Example Usage)

### Ví dụ 1: Vay 100 triệu, lãi đơn 12%/năm, 12 tháng

```
Tiền gốc: 100,000,000 VND
Lãi suất: 12% năm
Kỳ hạn: 12 tháng
Loại lãi: Simple

Kết quả:
- Tiền lãi: 12,000,000 VND
- Tổng tiền: 112,000,000 VND
- Trả hàng tháng: 9,333,333 VND
```

### Ví dụ 2: Vay 200 triệu, lãi kép 9%/năm, 24 tháng

```
Tiền gốc: 200,000,000 VND
Lãi suất: 9% năm
Kỳ hạn: 24 tháng
Loại lãi: Compound

Kết quả:
- Tiền lãi: ~19,488,000 VND
- Tổng tiền: ~219,488,000 VND
- Trả hàng tháng: ~9,145,333 VND
```

## Giao diện (User Interface)

Chương trình có giao diện hiện đại với:
- Font chữ Segoe UI dễ đọc
- Màu sắc hài hòa, chuyên nghiệp
- Bố cục rõ ràng, trực quan
- Nút bấm với hiệu ứng hover
- Bảng dữ liệu có màu xen kẽ
- Thông báo rõ ràng cho người dùng

## Xử lý lỗi (Error Handling)

- Xác thực dữ liệu đầu vào
- Thông báo lỗi rõ ràng
- Xử lý ngoại lệ khi đọc/ghi Excel
- Kiểm tra file tồn tại

## Tác giả (Author)

Dự án được phát triển để quản lý và tính toán lãi suất cho các khoản vay một cách chuyên nghiệp và dễ dàng.

## Giấy phép (License)

Dự án này sử dụng EPPlus với giấy phép NonCommercial. Để sử dụng thương mại, vui lòng mua giấy phép EPPlus.

## Hỗ trợ (Support)

Nếu gặp vấn đề hoặc có câu hỏi, vui lòng tạo issue trên GitHub.
