# InvoiceFlow E2E — Hệ thống Tự động hóa Nhập liệu và Đối soát Hóa đơn trên n8n

Hệ thống tự động hóa quy trình đối chiếu báo cáo tờ khai hải quan (`BaoCaoToKhai`), dữ liệu hóa đơn SAP (`HK63` qua `LOCAL_FOLDER` hoặc SAP API `YINVLST` / `TCURR`), cập nhật tệp làm việc tháng (`1. Sale invoices to upload*.xlsx`), xử lý các ngoại lệ nghiệp vụ, tạo tệp chuẩn hóa đơn Viettel (`00. UPLOAD INVOICES VIETTEL.xlsx`), lưu trữ thư mục tờ khai đã xử lý và xuất nhật ký kiểm toán Excel (`*_AUDIT.xlsx`) trên nền tảng **n8n**.

> **Lưu ý bảo mật dữ liệu:** Toàn bộ dữ liệu Excel mẫu và tệp kết quả (`.xls`, `.xlsx`) trong `finance_share/` và `InvoiceFlow_E2E_N8N/data/` đã được loại trừ khỏi repository công khai thông qua [`.gitignore`](.gitignore) nhằm bảo vệ dữ liệu nghiệp vụ của doanh nghiệp.

---

## 1. Cấu Trúc Dự Án & Thư Mục Dữ Liệu (`finance_share/`)

Docker Compose mount trực tiếp thư mục `../finance_share` vào đường dẫn `/files/finance` bên trong container n8n:

```text
invoiceflow-n8n-e2e/
├── .gitignore
├── README.md
├── InvoiceFlow_E2E_N8N/
│   ├── docker-compose.yml                 <-- Cấu hình container n8n (v2.4.6) & mount finance_share
│   ├── Dockerfile                         <-- Cài đặt exceljs và xlsx cho n8n Task Runner
│   ├── InvoiceFlow_E2E_N8N_IMPORT.json    <-- File workflow n8n (20 node, Stage 1 song song + Merge 08M)
│   ├── WORKFLOW_ANALYSIS.md               <-- Đặc tả chi tiết luồng xử lý và từng node
│   ├── code/                              <-- Mã nguồn JavaScript của các Code Node (02, 04, 06, 07b, 08L, 09, 14)
│   ├── config/                            <-- Hợp đồng JSON cho SAP API và ánh xạ tờ khai hải quan
│   ├── data/                              <-- Thư mục đặt dữ liệu tham chiếu cục bộ (không push Excel lên Git)
│   └── tests/                             <-- Bộ kiểm thử tự động hợp đồng, đường dẫn và ngoại lệ (17/17 PASS)
└── finance_share/                         <-- Thư mục chia sẻ dữ liệu đầu vào/đầu ra cho n8n
    ├── Bảng kê tờ khai HQ/                <-- Chứa các thư mục TK* (BaoCaoToKhai*.xls*) và lưu trữ {MM}.{YYYY}/
    ├── Desktop/                           <-- Chứa file SAP export cục bộ ({Ma_cong_ty} {mm}.{yy}.XLSX)
    ├── 5. Sale invoices/                  <-- Chứa 00. UPLOAD INVOICES VIETTEL.xlsx và {YYYY}/{MM}.{YYYY}/1. Sale invoices to upload*.xlsx
    └── results/                           <-- Thư mục nhận tệp Viettel đầu ra và nhật ký kiểm toán (*_AUDIT.xlsx)
```

---

## 2. Hướng Dẫn Triển Khai & Vận Hành Trên n8n

### Bước 1: Chuẩn bị thư mục dữ liệu & khởi động Docker n8n
1. Đặt các tệp dữ liệu nguồn Excel (tờ khai `TK*/BaoCaoToKhai*.xls*`, tệp SAP `Desktop/{Ma_cong_ty} {mm}.{yy}.XLSX` nếu chạy `LOCAL_FOLDER`, mẫu `5. Sale invoices/00. UPLOAD INVOICES VIETTEL.xlsx`, tệp làm việc tháng `5. Sale invoices/{YYYY}/{MM}.{YYYY}/1. Sale invoices to upload*.xlsx`) vào đúng cấu trúc `finance_share/`.
2. Mở terminal tại `InvoiceFlow_E2E_N8N/` và chạy:
```powershell
docker volume create n8n_data
docker compose up -d
```
3. Mở trình duyệt truy cập `http://localhost:5678`, import tệp [`InvoiceFlow_E2E_N8N_IMPORT.json`](InvoiceFlow_E2E_N8N/InvoiceFlow_E2E_N8N_IMPORT.json) và cấu hình thông tin xác thực nếu dùng nhánh `SAP_API`.

### Bước 2: Nhập Thông Số Khởi Chạy Tại Form (`01 - Form 1 Click`)
Mở workflow **`InvoiceFlow | SAP API + Baocaotokhai + Viettel + Audit | E2E Review`**, bấm **Execute Workflow** và nhập các trường:
- **`Profile`**: chọn `FINANCE_VN`.
- **`Ngay_xu_ly`**: ngày hóa đơn cần xử lý theo định dạng `YYYY-MM-DD` (ví dụ: `2026-08-14`).
- **`Ma_cong_ty`**: mã tổ chức bán hàng trên SAP (ví dụ: `HK63`).
- **`Nguon_SAP`**:
  - `LOCAL_FOLDER` *(Mặc định)*: đọc file SAP Excel trực tiếp trong `finance_share/Desktop/`.
  - `SAP_API`: gọi API SAP `YINVLST` & `TCURR`.
- **`Ty_gia_USD_VND`**: tỷ giá USD/VND (ví dụ: `25890`). Bắt buộc nhập ở chế độ `LOCAL_FOLDER`; nếu nhập khi chạy `SAP_API`, hệ thống sẽ dùng tỷ giá này và bỏ qua node `08 - SAP TCURR API` qua nhánh `07c`.
- **`Ghi_de_path_to_khai` / `Ghi_de_path_template` / `Ghi_de_path_sap` / `Ghi_de_thu_muc_ket_qua`**:
  - **Để trống** nếu sử dụng cấu trúc chuẩn tại Mục 1.
  - Có thể truyền đường dẫn tương đối trong `finance_share/`, đường dẫn Docker (`/files/finance/...`) hoặc đường dẫn Windows được chuẩn hóa (mọi chuỗi chứa `..` đều bị chặn bởi `Node 02`).

### Bước 3: Kết Quả Đầu Ra Tại `finance_share/results/`
- **Khi đạt (`READY_FOR_REVIEW`):**
  - Sinh 3 tệp trong `finance_share/results/`: `00. UPLOAD INVOICES VIETTEL.xlsx` (giữ định dạng template gốc, `fullCalcOnLoad = true`, sẵn sàng để nhân viên kiểm tra trước khi tải lên cổng Viettel), tệp làm việc `1. Sale invoices to upload ...xlsx` đã cập nhật, và báo cáo kiểm toán `*_AUDIT.xlsx` (2 sheet `Tom_tat` & `Nhat_ky_xu_ly`).
  - Tự động di chuyển thư mục `TK*` vừa xử lý vào `Bảng kê tờ khai HQ/{MM}.{YYYY}/` (ghi log `ĐÃ CHUYỂN`).
- **Khi bị chặn (`BLOCKED`):**
  - Không xuất `00. UPLOAD INVOICES VIETTEL.xlsx` và không di chuyển thư mục `TK*` (`CHƯA CHUYỂN`).
  - Chỉ xuất `*_AUDIT.xlsx` để phục vụ rà soát nguyên nhân lỗi.

---

## 3. Cơ Chế Xử Lý Ngoại Lệ & Kiểm Soát An Toàn (Fail-Closed)

- **Chặn sai lệch & chống tự suy đoán:** Tự động dừng (`BLOCKED`) khi thiếu tệp, thiếu tỷ giá, chênh lệch tổng tờ khai `> ±0.03` (bao gồm kiểm tra trị tuyệt đối `absCheck` chống cộng/trừ triệt tiêu), sai khác từng nhóm hóa đơn (`colAmount: AN`), số lượng không phải số hợp lệ (`toQtyStrict`), thiếu hóa đơn hai chiều Hải quan ↔ SAP hoặc thiếu đơn giá tham chiếu `PIR price` (`#N/A`).
- **Xử lý ngoại lệ nghiệp vụ tự động:**
  - Loại bỏ dòng có `Quantity = 0` và ghi nhận ngoại lệ (`EXCEPTION_QTY_ZERO`).
  - Hóa đơn gắn nhiều tờ khai: giữ mã kép tại cột `R`, nối số tờ khai bổ sung vào cột Diễn giải `P` (`EXCEPTION_MULTI_DECL_P`) và xóa công thức trùng trên sheet `Tờ khai` (`EXCEPTION_TK_DUP_INVOICE`).
  - Lệch cột kiểm tra `AC`: tự động kiểm thử bỏ phép chia `/10` tại cột `AA` nếu làm khớp thành tiền (`EXCEPTION_AA_DIV10`); nếu vẫn vượt ngưỡng dung sai sẽ dừng (`EXCEPTION_AC_UNRESOLVED`).
- **Lưu ý vận hành:** Đóng toàn bộ tệp Excel trong `results/` trước khi chạy n8n để tránh lỗi khóa tệp trên Windows (`EACCES`).

---

## 4. Chạy Bộ Kiểm Thử Tự Động (Không Cần Dữ Liệu Excel Thật)

Từ thư mục `InvoiceFlow_E2E_N8N/`, chạy lần lượt 3 bộ kiểm thử Node.js:

```powershell
cd InvoiceFlow_E2E_N8N
node tests/test_config_mock.js
node tests/test_engine_mock.js
node tests/test_exceptions_and_real_e2e.js
```

Toàn bộ `17/17` ca kiểm thử hợp đồng cấu hình, bảo vệ Path Traversal, đối soát lõi và xử lý ngoại lệ chạy hoàn toàn bằng mock contract/in-memory workbook (không phụ thuộc tệp Excel mẫu).
