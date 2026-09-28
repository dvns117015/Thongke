# Thongke — trích xuất số liệu và danh sách từ báo cáo KTGS

Công cụ đọc các file Word báo cáo công tác kiểm tra, giám sát, kỷ luật đảng
(hằng tháng, 6 tháng, năm) rồi tổng hợp **số liệu** và **danh sách** theo
**tháng, quý, năm** ra một file Excel.

## Hai loại file đầu vào

| Loại | Ví dụ tên file | Công cụ lấy ra |
|---|---|---|
| **Biểu thống kê số liệu** (bảng 191 dòng, mã TT `I, a, aa, ab…`, 4 cột cấp TW/tỉnh/huyện/cơ sở) | `5.1_Thống kê số liệu.docx` | Số liệu từng dòng, từng cấp. Kỳ đọc từ dòng "Tháng 5 năm 2025" dưới tiêu đề THỐNG KÊ |
| **Báo cáo lời văn** | `5_BAO CAO tháng 5-2025.doc`, `6_BAO CAO 6 tháng 2025.doc` | Số chi bộ, số đảng viên; danh sách đảng viên bị kỷ luật; hoạt động kiểm tra/giám sát/đơn thư; nhiệm vụ kỳ tới |

- Đọc được cả `.docx` và `.doc` (Word 97-2003) mà **không cần cài Word hay LibreOffice**.
- Loại file và kỳ báo cáo được nhận diện **theo nội dung**, không phụ thuộc tên file.

## Cài đặt (một lần)

Cần Python 3.10 trở lên.

```
pip install -r requirements.txt
```

## Cách dùng

Chép tất cả file Word vào thư mục `du_lieu/` (thư mục con cũng được), rồi:

```
# Một hoặc nhiều kỳ cụ thể (mỗi kỳ là một nhóm cột/dòng trong cùng file Excel)
python -m thongke tong-hop du_lieu --ky 2025-05 --ky 2025-Q2 --ky 2025 -o ket_qua/TongHop.xlsx

# Tự tạo mọi tháng / quý / năm có dữ liệu
python -m thongke tong-hop du_lieu --theo thang -o ket_qua/theo_thang.xlsx
python -m thongke tong-hop du_lieu --theo quy --nam 2025 -o ket_qua/quy_2025.xlsx

# Xem nhanh máy đọc được gì từ một file (để kiểm tra)
python -m thongke xem "du_lieu/6_BAO CAO 6 tháng 2025.doc"
```

Trên Windows có thể nhấp đúp `tong_hop.bat` (tạo sẵn 3 file theo tháng, quý, năm trong `ket_qua/`).

Cách viết kỳ: `2025-05` (tháng 5), `2025-Q2` (quý II), `2025-6T` (6 tháng đầu năm), `2025` (cả năm).

## File Excel kết quả

| Sheet | Nội dung |
|---|---|
| Tổng quan | Mỗi kỳ: đơn vị, số chi bộ, số đảng viên, file nào đã dùng, **tháng nào thiếu biểu**, số đảng viên đã/đang xử lý kỷ luật |
| Số liệu | Toàn bộ biểu thống kê, mỗi kỳ một cột (chỉ hiện các cấp có số liệu; luôn có Cấp cơ sở) |
| Kỷ luật | Mỗi đảng viên một dòng: họ tên, chi bộ, hình thức, trạng thái, cấp xử lý, lý do, **diễn biến qua các báo cáo** |
| Kiểm tra - Giám sát | Hoạt động kiểm tra / giám sát / đơn thư theo chủ thể (Đảng ủy, cấp ủy chi bộ, UBKT), đối tượng, tình trạng |
| Hoạt động khác | Tham mưu, xây dựng ngành, đánh giá chung… |
| Phương hướng | Nhiệm vụ trọng tâm kỳ tới |
| Cảnh báo | Những điểm cần xem lại (xem dưới) |

## Quy tắc tổng hợp

**Số liệu**
- Có biểu thống kê đúng bằng kỳ cần lấy (vd biểu "Quý II năm 2025") → dùng luôn biểu đó.
- Không có → **cộng các biểu tháng** trong kỳ. Tháng nào thiếu biểu được ghi ở Tổng quan và Cảnh báo.
- Ô trống giữ là trống (khác với số 0); khi cộng, ô trống tính là 0.
- Mẫu biểu có vài mã TT bị trùng (`pb`, `rb`, `rc`) — công cụ tự tách thành `pb`, `pb#2`… nên không bị cộng lẫn.
- Nhiều file cùng một tháng → lấy file sửa gần nhất và ghi cảnh báo.

**Danh sách**
- Lấy từ các báo cáo **nằm trong** kỳ. Báo cáo rộng hơn kỳ (vd báo cáo 6 tháng khi tổng hợp quý II)
  mặc định bị bỏ qua và ghi cảnh báo; thêm `--lay-ca-bao-cao-giao-nhau` để tính cả.
- Đảng viên được nhận ra qua cụm "đồng chí / đ/c Họ Tên – Đảng viên Chi bộ N".
- Trạng thái kỷ luật:
  - **Đã thi hành** — "Thi hành kỷ luật Khiển trách…", "thi hành kỷ luật bằng hình thức…", "Quyết định thi hành kỷ luật…"
  - **Đề nghị / đang xem xét** — "đề nghị thi hành kỷ luật", "báo cáo … xem xét thi hành kỷ luật"
  - **Dự kiến** — nằm trong phần phương hướng / nhiệm vụ trọng tâm
- Cùng một người xuất hiện ở nhiều báo cáo → gộp một dòng, lấy trạng thái cao nhất, giữ diễn biến.

**Cảnh báo tự động**
- Số đảng viên "đã thi hành kỷ luật" trong báo cáo lời văn **khác** tổng hàng `da` + `la` của biểu thống kê.
- Báo cáo của đơn vị này nhưng nhắc tên đơn vị khác (thường sót do sao chép mẫu cũ).
- Không xác định được kỳ, thiếu biểu tháng, file không đọc được.

## Để công cụ đọc chính xác

- Giữ nguyên mẫu biểu thống kê (cột TT và 4 cột cấp). Dòng kỳ ngay dưới tiêu đề: `Tháng 5 năm 2025` / `Quý II năm 2025` / `Năm 2025`.
- Tiêu đề báo cáo giữ dạng "… kỷ luật đảng tháng 5 và phương hướng …" hoặc "… 6 tháng đầu năm …".
- Khi nhắc đảng viên, viết đủ "đồng chí Họ Tên – Đảng viên Chi bộ N".
- Biểu thống kê nên để ở `.docx`. File `.doc` chỉ đọc được phần chữ (đủ cho báo cáo lời văn, không đủ cho bảng số liệu).

## Kiểm thử

```
pip install -r requirements-dev.txt
python -m pytest
```

Test dùng file Word tự tạo với dữ liệu giả. **Không đưa file báo cáo thật lên git** (thư mục `du_lieu/`
và mọi file `.xlsx` đã nằm trong `.gitignore`).
