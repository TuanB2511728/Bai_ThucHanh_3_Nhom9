KHAI THÁC VÀ TRUY VẤN DỮ LIỆU NÂNG CAO TRÊN EXCEL VÀ GOOGLE SHEETS
I.MỤC LỤC
1. Tổng quan về Quản trị Dữ liệu phẳng
2. Hệ thống Công cụ Truy vấn và Tính toán trong Excel
5. Báo cáo Thực hành và Ứng dụng thực tiễn
6. Tổng kết và Đánh giá Chiến lược
7. Tài liệu tham khảo
II.NỘI DUNG
1. TỔNG QUAN VỀ QUẢN TRỊ DỮ LIỆU PHẲNG TRONG MÔI TRƯỜNG BẢNG TÍNH
Trong kiến trúc dữ liệu doanh nghiệp, Excel và Google Sheets không chỉ là bảng tính mà là các hệ quản trị cơ sở dữ liệu phẳng (flat databases). Việc làm chủ cấu trúc phân cấp Workbook - Sheet - Cell là bước đầu tiên để xây dựng các hệ thống báo cáo có khả năng mở rộng.
• Kiến trúc và Giới hạn hệ thống: Một Workbook là thực thể chứa nhiều Sheet. Trong khi số lượng Sheet không bị giới hạn cứng (phụ thuộc vào bộ nhớ khả dụng - Available Memory của hệ thống), mỗi Sheet có giới hạn nghiêm ngặt là 1.048.576 dòng và 16.384 cột (đến cột XFD). Đối với các chuyên gia, việc hiểu giới hạn này giúp quyết định thời điểm cần chuyển dịch dữ liệu sang các hệ quản trị SQL chuyên dụng.
• Phân loại và Định dạng dữ liệu: Dữ liệu được phần mềm tự động nhận diện dựa trên Windows Regional Settings (đặc biệt quan trọng với định dạng Ngày tháng và Số).
    ◦ Dữ liệu số/ngày tháng: Canh lề phải, lưu trữ dưới dạng trị số để tính toán.
    ◦ Dữ liệu chuỗi: Canh lề trái.
    ◦ Dữ liệu công thức: Bắt đầu bằng dấu = hoặc +.
• Điều hướng chiến lược: Hiệu suất làm việc với dữ liệu lớn được tối ưu qua hệ thống phím tắt: Ctrl + Home (về ô A1), PageUp/Down (di chuyển theo trang), và Alt + PageUp/Down (di chuyển ngang).

2. HỆ THỐNG CÔNG CỤ TRUY VẤN VÀ TÍNH TOÁN TRONG MICROSOFT EXCEL
Để biến một bảng dữ liệu tĩnh thành một công cụ phân tích động, chuyên viên cần làm chủ kỹ thuật tham chiếu địa chỉ và hệ thống hàm logic.
2.1. Địa chỉ ô: Nền tảng của các Template có khả năng tái sử dụng
Việc lựa chọn loại địa chỉ không chỉ là thao tác kỹ thuật mà là tư duy thiết kế hệ thống:
• Địa chỉ tương đối (A1): Tự động thay đổi theo hướng sao chép.
• Địa chỉ tuyệt đối (A1): Cố định hoàn toàn, dùng cho các tham số chuẩn (tỷ giá, hệ số).
• Địa chỉ hỗn hợp (A1hoặcA1): Đây là chìa khóa để xây dựng các ma trận tính toán phức tạp. Bằng cách cố định dòng hoặc cột, một công thức duy nhất có thể được kéo cho toàn bộ bảng tính, giảm thiểu sai sót và thời gian bảo trì.
2.2. Danh mục Hàm chiến lược và Kỹ thuật Xử lý lỗi
Hệ thống hàm được chia thành các nhóm phục vụ mục đích truy vấn và làm sạch dữ liệu. Đồng thời, việc hiểu rõ bản chất 7 mã lỗi phổ biến theo nguồn tài liệu gốc là bắt buộc:
Mã lỗi	Nguyên nhân và Ý nghĩa kỹ thuật
#DIV/0!	Phép tính thực hiện chia cho giá trị 0.
#N/A	Không tìm thấy giá trị trong hàm tham chiếu hoặc hàm thiếu đối số.
#NAME?	Excel không nhận diện được tên vùng hoặc tên hàm.
#NULL!	Xảy ra khi xác định giao của 2 vùng nhưng thực tế chúng không giao nhau.
#NUM!	Phát hiện lỗi đối với dữ liệu kiểu số (ví dụ số quá lớn hoặc quá nhỏ).
#REF!	Tham chiếu đến một địa chỉ ô không hợp lệ (thường do dòng/cột bị xóa).
#VALUE!	Công thức chứa các toán hạng hoặc toán tử sai kiểu dữ liệu.

3. KHAI THÁC DỮ LIỆU NÂNG CAO THEO MÔ HÌNH DATABASE
Khi xử lý các tập dữ liệu có nhiều điều kiện ràng buộc, phương pháp sử dụng Criteria Range (Vùng tiêu chuẩn) là tiền thân quan trọng của tư duy truy vấn SQL.
• Thiết lập Criteria Range: Yêu cầu tối thiểu 2 hàng (hàng tiêu đề và hàng điều kiện).
    ◦ Logic AND: Các điều kiện nằm trên cùng một hàng (ví dụ: =AND(H4 >= 4, K4 = "C")).
    ◦ Logic OR: Các điều kiện nằm trên các hàng khác nhau.
• Advanced Filter: Công cụ này vượt trội hơn AutoFilter nhờ khả năng trích xuất dữ liệu sang vị trí khác (Copy to another location), giữ cho dữ liệu gốc luôn sạch.
• Nhóm hàm Database (D-Functions): Các hàm DSUM, DAVERAGE, DCOUNT, DMAX, DMIN cho phép thực hiện thống kê trực tiếp trên cơ sở dữ liệu dựa trên vùng tiêu chuẩn, giúp báo cáo trở nên linh hoạt và mạnh mẽ hơn so với các hàm IF lồng nhau.

4. CHIẾN LƯỢC TRUY VẤN DỮ LIỆU ĐỘNG TRÊN GOOGLE SHEETS
Google Sheets cung cấp các hàm hiện đại giúp xử lý dữ liệu theo thời gian thực với cú pháp tối giản nhưng hiệu quả cao.
• Hàm FILTER: =FILTER(range, condition1, ...). Tạo ra các tập dữ liệu con (sub-datasets) có tính cập nhật tự động.
• Sức mạnh của hàm QUERY: Sử dụng ngôn ngữ Google Visualization API (SQL-like).
    ◦ Lưu ý kỹ thuật: Nếu dữ liệu là dải ô trực tiếp, sử dụng ký hiệu cột (A, B, C). Nếu dữ liệu là kết quả của hàm khác, sử dụng Col1, Col2.
    ◦ Mệnh đề nâng cao: Hỗ trợ matches (biểu thức chính quy - Regex), contains, starts with cho các kịch bản lọc chuỗi phức tạp.

5. TỔNG QUAN NỘI DUNG THỰC HÀNH
5.1. Excel Case Study: Hệ thống Quản lý Điểm Tổng hợp
Trong bài tập thực hành, chúng ta xây dựng mô hình quản lý dựa trên các ràng buộc kinh doanh (Business Rules) cụ thể:
5.1.1. Truy xuất dữ liệu: Sử dụng VLOOKUP kết hợp hàm chuỗi để lấy điểm từ các Sheet thành phần. Thuật toán yêu cầu tách 8 ký tự cuối của Mã SV (RIGHT(MaSV, 8)) để làm giá trị tìm kiếm chuẩn trong danh sách.
5.1.2. Logic tính toán :
    ◦ Nếu Điểm Lý thuyết (Điểm LT) = -3 (Vắng thi), Điểm tổng mặc định là 0.
    ◦ Nếu Điểm Bài tập lớn = 0 hoặc vắng quá 2 buổi, sinh viên không đủ điều kiện hoàn thành.
    ◦ Trọng số tính toán: Điểm tổng = Điểm TH1 + Điểm TH2 + Điểm TH3 + Điểm BT lớn + 0.6 * Điểm LT.
5.1.3. Trực quan hóa: Sử dụng Biểu đồ tròn (Pie Chart) để thể hiện phân bố điểm chữ (A, B, C, D, F). Điều này giúp nhà quản lý nắm bắt nhanh tỷ lệ sinh viên đạt/trượt qua từng học kỳ.
5.2. Google Sheets Case Study
Ứng dụng hàm QUERY để thực hiện 4 loại báo cáo chiến lược từ sheet Tong_hop:
• SV Ngành CNTT: SELECT * WHERE D = 'Information Technology'
• SV Lớp DI18T9A1 môn NT_CNTT: SELECT * WHERE C = 'DI18T9A1' AND E = 'Fundamentals of Information Technology'
• SV Kết quả yếu (D, F): SELECT * WHERE E = 'D' OR E = 'F'
• SV Đăng ký hơn 1 môn: Sử dụng SELECT A, B, C, D, COUNT(E) GROUP BY A, B, C, D HAVING COUNT(E) > 1.

6. TỔNG KẾT VÀ ĐÁNH GIÁ CHIẾN LƯỢC
Việc chuyển đổi từ tư duy "bảng tính thủ công" sang "hệ quản trị dữ liệu" là bước tiến quan trọng của một chuyên viên phân tích.
• Excel vẫn là công cụ thống trị cho các phép tính toán sâu, nặng về công thức và xử lý dữ liệu cục bộ.
• Google Sheets chiếm ưu thế tuyệt đối trong việc cộng tác và truy vấn động nhờ ngôn ngữ QUERY mạnh mẽ.

7. TÀI LIỆU THAM KHẢO
• Nguồn gốc: "Buổi thực hành 2: Microsoft Excel và Google Sheets" – N.M.Trung.
• Hỗ trợ biên tập: Nội dung được cấu trúc hóa và tối ưu chuyên môn bởi mô hình ngôn ngữ lớn (LLM) dựa trên tiêu chuẩn tài liệu kỹ thuật quốc tế.

