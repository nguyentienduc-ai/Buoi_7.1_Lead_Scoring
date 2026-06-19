# AI Lead Scoring & Automation System (Real Estate)

Dự án tự động hóa thu thập dữ liệu khách hàng từ Google Sheets, chấm điểm tiềm năng (Lead Scoring) dựa trên AI/Heuristics và cung cấp giao diện kiểm duyệt (Human-in-the-loop) hỗ trợ xuất báo cáo Excel.

---

## 🛠️ Hướng dẫn cài đặt

Cài đặt các thư viện cần thiết trước khi bắt đầu:
```bash
pip install -r requirements.txt
```

---

## 🚀 Cách khởi chạy ứng dụng

Dự án cung cấp **2 lựa chọn giao diện** tùy thuộc vào nhu cầu sử dụng của bạn:

### Lựa chọn A: Giao diện Glassmorphism Custom (Khuyên dùng)
Giao diện Web cao cấp được thiết kế riêng bằng HTML/JS/CSS dạng kính mờ (Glassmorphism), có hiệu ứng nhảy số mượt mà và giao diện tối hiện đại.
1. Khởi chạy máy chủ:
   ```bash
   python server.py
   ```
2. Truy cập cổng: [http://localhost:9090](http://localhost:9090)

### Lựa chọn B: Giao diện Streamlit App
Giao diện tiêu chuẩn được xây dựng hoàn toàn bằng thư viện Streamlit của Python, đã được tối ưu hóa giao diện tối (Dark Mode) đồng bộ.
1. Khởi chạy Streamlit:
   ```bash
   streamlit run app.py
   ```
2. Truy cập cổng: [http://localhost:8501](http://localhost:8501)

---

## 📋 Cơ chế chấm điểm (Lead Scoring Rules)
Hệ thống sử dụng điểm số cơ bản là **100**. Điểm được điều chỉnh tự động:
- **Cộng 50 điểm (VIP)** khi nhận diện: Ngân sách lớn (>=20 tỷ VND, tài chính mạnh), loại hình cao cấp (Biệt thự, Penthouse, Shophouse, sàn văn phòng lớn), vị trí vàng (Q1, Phú Mỹ Hưng), đối tượng khách hàng (Chủ doanh nghiệp, Mua sỉ), pháp lý minh bạch (Sổ hồng riêng).
- **Trừ 50 điểm (Junk)** khi có các dấu hiệu: Yêu cầu phi thực tế (Giá Quận 1 siêu rẻ dưới 2 tỷ VND), không có nhu cầu (Nhầm số, dữ liệu cũ), spam/quảng cáo bảo hiểm, lỗi thông tin (Thuê bao, không liên lạc được).
- **Giữ nguyên 100 điểm (Neutral)** cho các nhu cầu chung cư/nhà phố tầm trung (3-10 tỷ VND), cần tư vấn vay ngân hàng thông thường.

*Hệ thống bảo toàn tuyệt đối các thay đổi/chú thích được phê duyệt thủ công bởi con người khi thực hiện Đồng bộ dữ liệu mới.*
