# [MINI PROJECT] HỆ THỐNG QUẢN TRỊ HẠ TẦNG & NHẬT KÝ KỸ THUẬT


## 📌 1. Phân Tích: Mô Hình IPO & DIKW (Dễ Hiểu)

* **IPO (Quy trình xử lý):**
  * **Input (Đầu vào):** File log thô (`log-he-thong.log`) và danh sách linh kiện.
  * **Process (Xử lý):** Dùng lệnh lọc các dòng lỗi; dùng hàm `VLOOKUP` tra tên thiết bị.
  * **Output (Đầu ra):** File báo cáo lỗi và bảng Excel đối soát hoàn chỉnh.

* **DIKW (Tháp dữ liệu):**
  * **Data:** Các dòng chữ trong file log (`2026-10-25 ERROR 404...`).
  * **Information:** Lọc ra được: *"Hôm nay có 15 lỗi CRITICAL"*.
  * **Knowledge:** Biết được: *"Trạm số 3 hay bị lỗi nhất vào buổi trưa"*.
  * **Wisdom:** Quyết định: *"Thay ổ cứng mới cho trạm số 3"*.

* **Bẫy đổi mạng (Mbps sang MB/s):**
  * Mạng $40\text{ Mbps}$ chia $8 = 5\text{ MB/s}$.
  * File log $500\text{ MB}$ mất: $500 / 5 = 100\text{ giây}$ để tải.

---

## 🛠️ 2. Triển Khai 4 Module

### Module 1: Tạo Cây Thư Mục (Terminal)

* **Gõ 1 dòng lệnh duy nhất để tạo đủ các thư mục:**
  ```cmd
  mkdir greentech-system\docs greentech-system\scripts greentech-system\data greentech-system\backup
Vào thư mục và tạo file .env bằng đường dẫn tương đối:DOScd greentech-system
echo MAT_KHAU=123456 > .env
Module 2: File Lọc Log Tự Động (auto_report.bat)Tạo file auto_report.bat nằm trong thư mục scripts/ với nội dung ngắn gọn:DOS@echo off
echo BAO CAO SU CO HE THONG > ..\backup\bao-cao-su-co.txt
echo ======================= >> ..\backup\bao-cao-su-co.txt

:: Dùng Pipe (|) và findstr để lọc lỗi, ghi tiếp (>>) vào file báo cáo
type ..\data\log-he-thong.log | findstr "CRITICAL ERROR" >> ..\backup\bao-cao-su-co.txt

echo Da loc xong loi!
pause
Module 3: Viết Tài Liệu (Markdown & Word)1. File docs/README.md (Đơn giản cho GitHub)Markdown# He Thong GreenTech IoT

## Huong Dan Chay Script
Chay file `scripts/auto_report.bat` de loc loi he thong.

## Checklist Kiem Tra
- [x] Kiem tra nguon dien
- [ ] Chay file loc log
- [ ] Kiem tra file Excel

## Danh Muc Thiet Bi
| Ma LK | Ten Thiet Bi | Loai |
|---|---|---|
| LK01 | O cung SSD | SSD |
| LK02 | Thanh RAM 8GB | RAM |
2. File docs/srs-he-thong.docx (Word)Trang 1: Trang bìa + Chèn Mục lục tự động (References $\rightarrow$ Table of Contents).Trang 2:Gán Heading 1: 1. Tong quan he thongGán Heading 2: 1.1 Muc tieu, 1.2 Danh sach thu mucTrang 3:Gán Heading 1: 2. Quan ly thiet biGán Heading 2: 2.1 Tra cuu linh kien, 2.2 Xu ly log loiModule 4: Đối Soát Bằng Excel (quan-ly-thiet-bi.xlsx)Quy tắc: Không gộp ô (Merge Cell), dòng 1 là tiêu đề, không để dòng trống.Bảng 1 (Sheet danh-muc-linh-kien): Cột A là ma_linh_kien, Cột B là ten_thiet_bi.Bảng 2 (Sheet trang-thai-tram): Cột B là ma_linh_kien, Cột C là ten_thiet_bi.Công thức tại ô C2 (kết hợp TRIM, $ và FALSE):Excel=VLOOKUP(TRIM(B2), 'danh-muc-linh-kien'!$A$2:$B$50, 2, FALSE)
TRIM(B2): Xóa dấu cách thừa tránh lỗi #N/A.$A$2:$B$50: Dấu $ để khóa chặt bảng khi kéo công thức xuống.2: Lấy cột tên thiết bị.FALSE: Tìm kiếm chính xác 100%.
