<div align="center">

# 🤖 Tự động Tổng Hợp Báo Cáo Kinh Doanh từ Dữ Liệu Thô bằng AI <br> trong Môi Trường Số

</div>

<div align="center">
  <img src="logo.jpg" alt="Logo dự án" width="2000"/>
</div>
## 📌 Giới Thiệu
Ứng dụng web Flask hỗ trợ **tự động xử lý dữ liệu Excel/CSV và sinh báo cáo kinh doanh** từ dữ liệu thô.  
Người dùng chỉ cần tải file dữ liệu, hệ thống sẽ:
- Chuẩn hóa dữ liệu  
- Tính toán KPI (doanh thu, lợi nhuận, biên lợi nhuận, tăng trưởng MoM…)  
- Vẽ biểu đồ xu hướng và top sản phẩm  
- Sinh báo cáo tự động bằng AI rule-based  

## ⚙️ Cài Đặt Môi Trường
Yêu cầu: **Python 3.10+**  

Cài đặt thư viện cần thiết:

pip install flask pandas matplotlib python-docx openpyxl
Chạy ứng dụng:
python app.py
Ứng dụng sẽ chạy tại: http://127.0.0.1:5000

🚀 Tính Năng
📂 Tải lên file dữ liệu (.xlsx, .csv)

🧹 Tiền xử lý dữ liệu, chuẩn hóa tên cột tự động

📈 Tính KPI: Tổng doanh thu, tổng lợi nhuận, biên lợi nhuận gộp, doanh thu trung bình tháng, tăng trưởng MoM

📊 Biểu đồ trực quan: Xu hướng doanh thu, Top sản phẩm bán chạy

📝 Sinh báo cáo tự động (AI rule-based) và hỗ trợ xuất ra file Word

📊 Kết Quả
Người dùng tải file dữ liệu → hệ thống trả về KPI + biểu đồ + báo cáo chi tiết

Báo cáo có thể xem trực tiếp trên web hoặc xuất ra file Word

Giao diện thân thiện, dễ sử dụng, hỗ trợ chọn file hoặc kéo-thả

📸 Demo Giao Diện
<div align="center"> <img src="demo.jpg" alt="Demo ứng dụng" width="800"/> </div> 
📌 Giải thích:

logo.jpg → logo bạn đã upload.

images/demo.png → bạn có thể chụp màn hình giao diện ứng dụng, lưu vào thư mục images/, rồi thay tên ảnh cho đúng.

