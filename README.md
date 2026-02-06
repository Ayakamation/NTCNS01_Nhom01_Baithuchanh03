NTCNS01 - Nhóm 01 - Bài Thực Hành Buổi 3
Microsoft Excel và Google Sheets
Thông tin chung

Môn học: NTCNS01
Nhóm: Nhóm 01
Giảng viên: N.M. Trung
Repo GitHub: https://github.com/Ayakamation/NTCNS01_Nhom01_Baithuchanh3
Ngày hoàn thành: Tháng 2/2026
Thành viên nhóm
Nguyễn Phát Đạt B2504789, Dương Văn Thành Hiệu B2504794, Nguyễn Ngọc Diệp B2504786, Nguyễn Tuấn Anh B2504784
Mô tả bài thực hành
Buổi thực hành hướng dẫn sử dụng Microsoft Excel 2013 và làm quen với Google Sheets. Nội dung chính bao gồm:

Cấu trúc Workbook và Sheet, thao tác cơ bản (nhập dữ liệu, di chuyển, chọn vùng, Gridline, thao tác Sheet).
Kiểu dữ liệu, công thức, địa chỉ ô, thông báo lỗi.
Thao tác cơ bản (đặt tên vùng, xóa, tự động điền AutoFill, thêm/xóa hàng/cột/ô, đóng bảng Freeze Panes).
Định dạng (hiển thị dữ liệu, canh lề, kẻ khung, tô màu, Format Painter).
Hàm cơ bản (toán học, thống kê, logic, xử lý chuỗi, ngày giờ, tìm kiếm VLOOKUP/HLOOKUP/MATCH).
Thao tác cơ sở dữ liệu (vùng tiêu chuẩn, trích lọc AutoFilter/Advanced Filter, hàm D* như DSUM, DCOUNT).
Tạo và chỉnh sửa biểu đồ (các loại, thành phần, sử dụng Design/Format).
Google Sheets: Hàm FILTER (lọc đơn/đa điều kiện, sắp xếp, tổng hợp) và QUERY (ngôn ngữ truy vấn, các mệnh đề select, where, group by, pivot, order by, limit, offset, label, format).

Buổi học giúp sinh viên xử lý dữ liệu, tính toán, thống kê, vẽ biểu đồ, và làm việc trực tuyến thời gian thực.
Nội dung chính (tóm tắt từ tài liệu "Buoi 3 N.M.Trung.pdf")
1. Microsoft Excel
1.1 Kiến thức cần chuẩn bị

Workbook: Tập tin .XLSX, gồm nhiều Sheet, không giới hạn số lượng (tùy bộ nhớ).
Sheet: Bảng tính với 1.048.576 hàng (1 → 1.048.576) và 16.384 cột (A → XFD), tổng 17.179.869.184 ô. Ô địa chỉ: A5; Vùng: B2:F10.
Di chuyển con trỏ ô: Chuột, bàn phím (↑ ↓ ← →, PageUp/Down, Alt+PageUp/Down, Ctrl+Home), Name Box, thanh cuộn.
Gridline: Bật/tắt View → Gridlines.
Thao tác Sheet: Right-click tab → Rename, Delete, Move/Copy, Hide/Unhide.
Kiểu dữ liệu: Số (canh phải), chuỗi (canh trái), luận lý (canh giữa). Nhập: Enter kết thúc, Esc hủy; Hiệu chỉnh: F2 hoặc Double-click.

1.2 Thao tác cơ bản

Đặt tên vùng: Name Box.
Xóa: Home → Clear (nội dung, định dạng).
AutoFill: Kéo Fill handle để điền dãy (số, ngày, chuỗi).
Thêm/xóa hàng/cột/ô: Home → Insert/Delete.
Freeze Panes: View → Freeze Panes (đóng dòng/cột đầu).

1.3 Định dạng

Hiển thị dữ liệu: Home → Number (Number, Date, Currency, Percentage).
Canh lề: Home → Alignment (top/middle/bottom, left/center/right, orientation, wrap text, merge).
Kẻ khung/tô màu: Home → Format Cells → Border/Fill.
Sao chép định dạng: Format Painter.

1.4 Công thức và hàm

Công thức: Bắt đầu =, toán tử (+ - * / ^ & > < = <>), ưu tiên (^ * / + - = <> > >= < <= NOT AND OR).
Địa chỉ: Tương đối (A1), tuyệt đối ($A$1), hỗn hợp ($A1).
Lỗi: #DIV/0!, #N/A, #NAME?, #NULL!, #NUM!, #REF!, #VALUE!.
Hàm toán học: INT, MOD, ROUND, SUM, SUMIF.
Thống kê/tìm kiếm: AVERAGE, COUNT, COUNTA, COUNTBLANK, COUNTIF, MAX, MIN, RANK.
Logic: AND, OR, NOT, IF.
Xử lý chuỗi: LOWER, UPPER, PROPER, TRIM, LEN, LEFT, RIGHT, MID, VALUE, FIND.
Ngày giờ: TODAY, NOW, DAY, MONTH, YEAR, WEEKDAY, DATE, TIME.
Tìm kiếm: VLOOKUP (tương đối/chính xác), HLOOKUP, MATCH.

1.5 Cơ sở dữ liệu

Danh sách như CSDL phẳng, hàng đầu là header.
Vùng tiêu chuẩn (Criteria range): 2 hàng, điều kiện AND/OR.
Trích lọc: AutoFilter (Data → Filter), Advanced Filter (Data → Advanced).
Hàm D*: DSUM, DAVERAGE, DMAX, DMIN, DCOUNT, DCOUNTA.

1.6 Biểu đồ

Chuẩn bị dữ liệu.
Tạo: Insert → Charts (Column, Line, Pie).
Thành phần: Chart Area, Title, Axis, Labels, Gridlines, Data Series, Legend.
Chỉnh sửa: Chart Tools → Design/Format.

2. Google Sheets

FILTER: Lọc dữ liệu (đơn/đa điều kiện AND/OR, sắp xếp, tổng hợp). Cú pháp: FILTER(range, condition1, [condition2]).
QUERY: Truy vấn SQL-like (select, where, group by, pivot, order by, limit, offset, label, format). Cú pháp: QUERY(data, query, [headers]).

Bài thực hành cụ thể

Phần 1 (Excel): Tạo Workbook với Sheet Tổng hợp điểm, nhập dữ liệu, tính điểm tổng (IF, ROUND), điểm thang 10, điểm chữ (VLOOKUP), trích lọc (Advanced Filter), thống kê (DCOUNT), đóng Freeze Panes, vẽ biểu đồ Pie. Lưu Buoi03_Họ_tên_Bài_1.docx.
Phần 2 (Google Sheets): Tạo gsheet, nhập dữ liệu Sheet Tong_hop, dùng QUERY lọc SV ngành CNTT, lớp DI18T9A1 môn Fundamentals, SV điểm D/F môn System Administration, SV đăng ký >1 môn.

File trong repository

Buoi03_Baithuchanh3.xlsx: File Excel bài thực hành.
Buoi 3 N.M.Trung.pdf: Tài liệu hướng dẫn (32 trang).
README.md: Tài liệu này.

