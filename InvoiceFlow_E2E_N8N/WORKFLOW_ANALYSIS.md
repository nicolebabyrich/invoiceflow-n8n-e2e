# Phân Tích Chi Tiết Workflow n8n — `InvoiceFlow | SAP API + Baocaotokhai + Viettel + Audit | E2E Review`

Tài liệu mô tả cấu trúc, luồng xử lý, input/output đầu-cuối và chi tiết từng node của file [InvoiceFlow_E2E_N8N_IMPORT.json](InvoiceFlow_E2E_N8N_IMPORT.json) (gồm **19 node xử lý** và 1 `stickyNote` ghi chú = **20 node**).

---

## 1. Input Đầu Tiên & Output Cuối Cùng

### Input đầu tiên
1. **Tham số nhập từ Form (`01 - Form 1 Click`):**
   - `Profile`: hồ sơ cấu hình (`FINANCE_VN`).
   - `Nguon_SAP`: `LOCAL_FOLDER` (đọc Excel SAP trên share) hoặc `SAP_API` (gọi API SAP).
   - `Ngay_xu_ly`: ngày hóa đơn cần xử lý (`YYYY-MM-DD`, ví dụ `2026-08-14`).
   - `Ma_cong_ty`: mã công ty SAP (để trống, placeholder `HK63`).
   - `Ty_gia_USD_VND`: tỷ giá USD/VND (để trống, placeholder `25890`; bắt buộc nhập nếu chọn `LOCAL_FOLDER`, và dùng làm tỷ giá ghi đè bỏ qua `TCURR API` nếu nhập ở chế độ `SAP_API`).
   - `Nguoi_thuc_hien`: mã nhân viên (tùy chọn).
   - `Ghi_de_path_to_khai`, `Ghi_de_path_template`, `Ghi_de_path_sap`, `Ghi_de_thu_muc_ket_qua`: tùy chọn ghi đè đường dẫn (để trống = dùng mặc định).
2. **Dữ liệu nguồn đọc từ thư mục share (`/files/finance`):**
   - Tờ khai hải quan: `Bảng kê tờ khai HQ/TK*/BaoCaoToKhai*.xls*` (hỗ trợ gom nhiều thư mục `TK*` cùng ngày).
   - Template Viettel: `5. Sale invoices/00. UPLOAD INVOICES VIETTEL.xlsx`
   - Dữ liệu SAP: `Desktop/{Ma_cong_ty} {mm}.{yy}.XLSX` (nếu `LOCAL_FOLDER`) hoặc API SAP `YINVLST` & `TCURR` (nếu `SAP_API`).
   - File làm việc tháng (chứa `Tờ khai`, `Ms Tiên invoices`, `PIR price`): `5. Sale invoices/{YYYY}/{MM}.{YYYY}/1. Sale invoices to upload*.xlsx`.

### Output cuối cùng
- **Nhánh ĐẠT (`READY_FOR_REVIEW`):**
  1. **Thư mục `results/` (`/files/finance/results`):**
     - `00. UPLOAD INVOICES VIETTEL.xlsx` (giữ định dạng template gốc, `fullCalcOnLoad = true`, chưa upload Viettel).
     - `1. Sale invoices to upload ...xlsx` (đã cập nhật `Tờ khai`, `Ms Tiên invoices` và điền đầy đủ công thức).
     - `HK63_{YYYYMMDD}_{ts}_AUDIT.xlsx` (gồm 2 sheet tiếng Việt `Tom_tat` và `Nhat_ky_xu_ly`).
  2. **Lưu trữ tự động:** Di chuyển thư mục `TK*` đã xử lý vào `Bảng kê tờ khai HQ/{MM}.{YYYY}/` và ghi trạng thái `ĐÃ CHUYỂN` cùng danh sách số tờ khai vào Audit Log.
  3. **Màn hình Form (`16 - End Form Tai file`):** Hiển thị thông báo hoàn tất (`showText`) tổng hợp số dòng và đường dẫn `results/`.
- **Nhánh BỊ CHẶN (`BLOCKED`):**
  1. **Thư mục `results/`:** chỉ lưu `*_AUDIT.xlsx` để kiểm tra nguyên nhân lỗi; không tạo/ghi `00. UPLOAD INVOICES VIETTEL.xlsx` và không di chuyển thư mục `TK*` (ghi `CHƯA CHUYỂN`).
  2. **Màn hình Form (`18 - End Form Tai log loi`):** Hiển thị danh sách lỗi chặn và hướng dẫn mở `*_AUDIT.xlsx` trong `results/`.

---

## 2. Sơ Đồ Flow Xử Lý (Stage 1 Song Song + Hội Tụ `08M`)

```text
01 (Form 1 Click) ─► 02 (Config & bảo vệ path)
                       ├─► [Nhánh 1: Hải quan & Template]
                       │     03 (Đọc Baocaotokhai) ─► 04 (Chọn theo từng thư mục TK*) ─► 05 (Đọc Template) ─► 06 (Buffer Template) ──┐
                       │                                                                                                             │
                       └─► [Nhánh 2: SAP & Tỷ giá]                                                                                   ▼
                             06b (IF Nguon_SAP == 'SAP_API'?)                                                                 08M (Merge)
                               ├─ [TRUE: SAP_API]                                                                                    ▲
                               │    └─► 07 (YINVLST) ─► 07b (Gom) ─► 07c (Có tỷ giá Form?)                                           │
                               │                                       ├─ [YES] ─────────────────────────────────────────────────────┤
                               │                                       └─ [NO] ─► 08 (SAP TCURR) ────────────────────────────────────┤
                               └─ [FALSE: LOCAL_FOLDER]                                                                              │
                                    └─► 07L (Đọc SAP Excel) ─► 08L (Chuẩn hóa SAP & PIR) ────────────────────────────────────────────┘
                                                                                                                                     │
  ┌──────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┘
  ▼
09 (Engine đối soát, cập nhật workbook, tạo Excel & Audit, lưu trữ TK*)
  │
  ▼
10 (IF status == 'READY_FOR_REVIEW'?)
  ├─ [TRUE]  ─► 11 (Lưu 00. UPLOAD INVOICES VIETTEL.xlsx vào results/) ─► 16 (End Form Hoàn tất)
  └─ [FALSE] ─► 18 (End Form Báo lỗi)
```

---

## 3. Chi Tiết Từng Node

| # | Tên Node | Loại Node | Công dụng | Input | Output |
|---|---|---|---|---|---|---|
| 1 | `01 - Form 1 Click` | `n8n-nodes-base.formTrigger` | Giao diện web để nhân viên nhập thông số chạy (`n8nUserAuth`, `lastNode`) | Dữ liệu form từ trình duyệt | JSON gồm `Profile`, `Nguon_SAP`, `Ngay_xu_ly`, `Ma_cong_ty`, `Ty_gia_USD_VND`, `Nguoi_thuc_hien`, `Ghi_de_path_*` |
| 2 | `02 - Config va bao ve path` | `n8n-nodes-base.code` | Nạp profile `FINANCE_VN` (`colAmount: 'AN'`), kiểm tra ngày/công ty/tỷ giá, chuẩn hóa path (tương đối, Windows trong `finance_share`, Docker) và chặn `..` | JSON từ node `01` | `{ cfg }` (`customsPath`, `templatePath`, `sapLocalPath`, `outputFolder`, `sapSourceMode`, `sapLocalRate`, mapping cột) |
| 3 | `03 - Doc Baocaotokhai tu share` | `n8n-nodes-base.readWriteFile` | *(Nhánh 1)* Đọc các file `BaoCaoToKhai*.xls*` từ thư mục share | `cfg.customsPath` | Danh sách item chứa binary `data`, `fileName`, `directory` |
| 4 | `04 - Chon file to khai` | `n8n-nodes-base.code` | *(Nhánh 1)* Gom theo từng thư mục `TK*`, chọn 1 file mới nhất trong mỗi thư mục và mã hóa `base64` | Binary từ `03` + `cfg` từ `02` | `{ cfg, customs, customsFiles[] }` |
| 5 | `05 - Doc template Viettel` | `n8n-nodes-base.readWriteFile` | *(Nhánh 1)* Đọc file mẫu `00. UPLOAD INVOICES VIETTEL.xlsx` | `cfg.templatePath` | Item chứa binary `data` của template |
| 6 | `06 - Buffer template` | `n8n-nodes-base.code` | *(Nhánh 1)* Mã hóa template sang `base64` và đóng gói cùng `customsFiles` | Binary từ `05` + JSON từ `04` | `{ cfg, customs, customsFiles, template }` |
| 7 | `06b - Chon nguon SAP` | `n8n-nodes-base.if` | *(Nhánh 2)* Rẽ nhánh theo `cfg.sapSourceMode === 'SAP_API'` (đọc từ `02`) | JSON từ `02` | `true` → `07` (`SAP_API`); `false` → `07L` (`LOCAL_FOLDER`) |
| 8 | `07 - SAP YINVLST API` | `n8n-nodes-base.httpRequest` | *(Nhánh API)* Gọi API SAP lấy danh sách invoice trong tháng (`httpBasicAuth`, timeout 120s) | `cfg.sapInvoiceUrl`, `companyCode`, `monthStart`, `monthEnd` | JSON response danh sách invoice từ SAP |
| 9 | `07b - Gom response SAP` | `n8n-nodes-base.code` | *(Nhánh API)* Gom các item trả về từ `07` thành một cấu trúc đồng nhất | Output từ `07` | JSON response SAP đã gom |
| 10 | `07c - Kiem tra ty gia Form` | `n8n-nodes-base.if` | *(Nhánh API)* Kiểm tra `Number(cfg.sapLocalRate || 0) > 0` để dùng tỷ giá nhập tay trên Form | `cfg.sapLocalRate` từ `02` | `YES` → `08M` (bỏ qua `08`); `NO` → `08` |
| 11 | `08 - SAP TCURR API` | `n8n-nodes-base.httpRequest` | *(Nhánh API)* Gọi API SAP lấy tỷ giá USD/VND ngày `Ngay_xu_ly - 1` | `cfg.sapRateUrl`, `validOn` | `{ rateVndPerUsd, validOn }` |
| 12 | `07L - Doc file SAP tu folder` | `n8n-nodes-base.readWriteFile` | *(Nhánh Local)* Đọc file `Desktop/{Ma_cong_ty} {mm}.{yy}.XLSX` từ share | `cfg.sapLocalPath` từ `02` | Item chứa binary `data` của file SAP Excel |
| 13 | `08L - Chuan hoa SAP tu Excel` | `n8n-nodes-base.code` | *(Nhánh Local)* Đọc `HK63` theo tên header, `toQtyStrict` chống ép chuỗi lỗi thành `0`, lọc `ZEX`/`VN20`/`UNITED WELL`/không hủy, tra `PIR price` từ `1. Sale invoices to upload*.xlsx` | Binary `data` từ `07L` + `cfg` từ `02` | `{ sapSourceMode, workFilePath, pirCount, skippedCounts, missingPirMaterials, droppedZeroRows, sapPayload, fxPayload }` |
| 14 | `08M - Hop luong Du lieu Stage 1` | `n8n-nodes-base.merge` | Hội tụ 2 nhánh song song (`combineAll`): chờ cả Nhánh 1 (`06`) và Nhánh 2 (`07c`/`08`/`08L`) hoàn thành | Input 0 từ `06`; Input 1 từ `07c`/`08`/`08L` | Dữ liệu hợp nhất chuyển sang `09` |
| 15 | `09 - Tao Excel Viettel va Audit` | `n8n-nodes-base.code` | Lõi nghiệp vụ: tách hóa đơn kép, đối chiếu 2 chiều (`Customs ↔ SAP`), kiểm `absCheck` và chênh lệch từng nhóm `±0.03`, chặn thiếu PIR (`#N/A`) hoặc lệch `AC`, tự bỏ `/10` ở `AA` nếu khớp, tạo `00. UPLOAD INVOICES VIETTEL.xlsx`, cập nhật `1. Sale invoices to upload*.xlsx`, chuyển thư mục `TK*` vào `MM.YYYY` khi đạt, và lưu `*_AUDIT.xlsx` | Dữ liệu từ `08M` / `06` / `07` / `08` | JSON `{ status, summary, errors, warnings, report, filenames, outputPath }` + Binary `{ audit_xlsx, excel? }` |
| 16 | `10 - Trang thai doi soat` | `n8n-nodes-base.if` | Kiểm tra `status === 'READY_FOR_REVIEW'` | Output từ `09` | `true` → `11`; `false` → `18` |
| 17 | `11 - Luu Excel tren server share` | `n8n-nodes-base.readWriteFile` | *(Nhánh Đạt)* Ghi binary `excel` (`00. UPLOAD INVOICES VIETTEL.xlsx`) vào `cfg.outputFolder` | Binary `excel` từ `09` | Item xác nhận đã ghi file ra `/files/finance/results/` |
| 18 | `16 - End Form Tai file` | `n8n-nodes-base.form` | *(Nhánh Đạt)* Hiển thị thông báo hoàn tất (`showText`) và đường dẫn `results/` | Dữ liệu từ `11` + `09` | Màn hình hoàn tất trên trình duyệt |
| 19 | `18 - End Form Tai log loi` | `n8n-nodes-base.form` | *(Nhánh Chặn)* Hiển thị danh sách lỗi chặn (`showText`) và hướng dẫn xem `*_AUDIT.xlsx` trong `results/` | Dữ liệu từ `10` + `09` | Màn hình báo lỗi trên trình duyệt |
