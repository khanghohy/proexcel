# 📊 Ứng Dụng Bảng Tính Excel & Quản Lý Ảnh Lưu Trữ SQL

Ứng dụng bảng tính phong cách Microsoft Excel trực quan, hiện đại, cho phép thêm hàng, chèn/dán ảnh chụp toàn bộ (**Ảnh chụp full**) và ảnh chụp khu vực (**Ảnh chụp 1 nửa**), lưu trữ dữ liệu và ảnh nhị phân trực tiếp vào **Cơ sở dữ liệu SQL** ngay trên máy tính của bạn.

---

## 🌟 Tính Năng Nổi Bật

1. **Giao diện bảng tính phong cách Excel hiện đại**:
   - Thanh Ribbon thao tác nhanh, thanh công thức `fx` hiển thị tọa độ ô đang chọn (A1, B2, C3...).
   - **Cột A**: `Tên` (cho phép bấm sửa nội dung trực tiếp, tự động lưu vào SQL).
   - **Cột B**: `Ảnh chụp full` (cho phép đính kèm ảnh chụp toàn màn hình).
   - **Cột C**: `Ảnh chụp 1 nửa` (cho phép đính kèm ảnh chụp 1 phần / khu vực).
   - Thao tác: Bấm **`+ Thêm hàng`** (hoặc phím tắt **Enter**, **Ctrl + N**), thêm nhanh 5 hàng, xóa hàng.
   - Hỗ trợ đổi giao diện **Sáng / Tối (Dark mode)**.

2. **Cơ chế chèn ảnh cực kỳ trực quan & tiện lợi**:
   - **Dán trực tiếp từ Clipboard (`Ctrl + V`)**: Bạn chụp ảnh màn hình bằng công cụ Snipping Tool hoặc PrtScn, bấm vào ô ảnh và nhấn `Ctrl + V`, ảnh sẽ ngay lập tức được dán và lưu vào SQL!
   - **Kéo thả (Drag & Drop)** tệp ảnh từ máy tính trực tiếp vào ô tương ứng.
   - **Bấm chọn tệp ảnh** từ hộp thoại duyệt file (`png`, `jpg`, `jpeg`, `webp`...).
   - **Xem phóng to (Lightbox Zoom)**: Bấm vào ảnh để mở trình xem toàn màn hình, hỗ trợ phóng to/thu nhỏ (Zoom in/out) và xoay ảnh 90°.
   - Đổi ảnh, tải ảnh về máy hoặc xóa ảnh bất kỳ lúc nào.

3. **Cơ sở dữ liệu SQL trên máy tính**:
   - Sử dụng chuẩn SQL **SQLite 3** độc lập, không cần cấu hình mật khẩu hay cài đặt rườm rà.
   - Vị trí tệp CSDL: `data\excel_app.db`.
   - **Lưu trữ ảnh trực tiếp trong SQL**: Tệp ảnh được lưu trữ dưới dạng nhị phân `BLOB` (`full_image_data`, `half_image_data`) kèm theo kích thước, định dạng mime và tên file trong bảng `records`.
   - **Trình Quản Lý & Truy Vấn SQL (SQL Studio)** tích hợp sẵn:
     - Xem đường dẫn tệp CSDL trên máy, dung lượng file DB, tổng số ảnh BLOB.
     - Xem cấu trúc bảng (Schema).
     - Soạn và thực thi câu lệnh SQL trực tiếp (`SELECT`, `PRAGMA`...) và xem kết quả dạng bảng.

4. **Xuất tệp Excel (.xlsx)**:
   - Tích hợp thư viện xử lý bảng tính chuyên nghiệp, cho phép tải về tệp `.xlsx` có nhúng sẵn ảnh thumbnail trực quan trong từng ô.

---

## 🚀 Hướng Dẫn Khởi Động & Truy Cập

Ứng dụng hiện đang được **host và chạy trực tiếp** trên máy tính của bạn tại:
- **Địa chỉ máy nội bộ (Localhost)**: [http://localhost:8000](http://localhost:8000)
- **Địa chỉ mạng LAN**: [http://0.0.0.0:8000](http://0.0.0.0:8000) (cho phép các thiết bị khác trong cùng mạng Wi-Fi/LAN truy cập).

### Khởi động lại khi cần:
1. Nhấp đúp vào tệp **`start_app.bat`** tại thư mục này.
2. Hoặc mở Terminal/PowerShell tại thư mục này và gõ:
   ```bash
   python -m uvicorn main:app --host 0.0.0.0 --port 8000 --reload
   ```

---

## 🗄️ Cấu Trúc Bảng SQL (`records`)

```sql
CREATE TABLE records (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    row_order INTEGER DEFAULT 0,
    name TEXT NOT NULL DEFAULT '',
    full_image_data BLOB,
    full_image_mime TEXT,
    full_image_name TEXT,
    full_image_size INTEGER DEFAULT 0,
    half_image_data BLOB,
    half_image_mime TEXT,
    half_image_name TEXT,
    half_image_size INTEGER DEFAULT 0,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);
```

---

## 📁 Cấu Trúc Thư Mục Dự Án

```
d:\New folder\
├── app\
│   ├── __init__.py
│   ├── database.py       # Tầng kết nối & thao tác CSDL SQL (BLOB storage)
│   └── routes.py         # Bộ định tuyến API (upload ảnh, xuất Excel, truy vấn SQL)
├── data\
│   └── excel_app.db      # Tệp cơ sở dữ liệu SQL lưu trữ toàn bộ dữ liệu & ảnh
├── static\
│   ├── css\
│   │   └── style.css     # Giao diện phong cách Excel & Dark/Light mode
│   ├── js\
│   │   └── app.js        # Logic bảng tính, dán ảnh Ctrl+V, kéo thả, Lightbox
│   └── index.html        # Giao diện chính người dùng
├── main.py               # Máy chủ FastAPI
├── start_app.bat         # Phím tắt 1-click khởi chạy trên Windows
└── README.md             # Hướng dẫn chi tiết
```
