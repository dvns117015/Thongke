"""Kiểm thử với file Word tự tạo (dữ liệu giả, không dùng file thật)."""

from pathlib import Path

import docx
import openpyxl
import pytest

from thongke.__main__ import main
from thongke.bao_cao import DA_THI_HANH, DE_NGHI, DU_KIEN, doc_bao_cao
from thongke.bieu_mau import doc_bieu_mau
from thongke.doc_reader import _bo_ma_truong
from thongke.ky import Ky, phan_tich_ky, phan_tich_tham_so_ky
from thongke.tong_hop import nap_du_lieu, tong_hop_ky

HANG_MAU = [
    ("I", "KIỂM TRA, GIÁM SÁT CỦA CẤP ỦY CÁC CẤP VÀ CHI BỘ"),
    ("d", "Thi hành kỷ luật"),
    ("da", "Số đảng viên bị thi hành kỷ luật"),
    ("dc", "Hình thức kỷ luật: Khiển trách"),
    ("II", "KIỂM TRA, GIÁM SÁT CỦA UBKT CÁC CẤP"),
    ("l", "Thi hành kỷ luật"),
    ("la", "Số đảng viên bị thi hành kỷ luật"),
    ("p", "Kiểm tra tài chính đảng"),
    ("pb", "Số đảng viên có vi phạm đến mức phải kỷ luật"),
    ("pđ", "Tổng số tiền vi phạm"),
    ("q", "Tiếp nhận, xử lý đơn thư"),
    ("pb", "Số đảng viên bị phản ánh đã thành lập đoàn giải quyết"),
] + [(f"x{c}", f"Hàng đệm {c}") for c in "abcdefghik"]


def tao_bieu_mau(duong_dan: Path, tieu_de_ky: str, so: dict[str, str]) -> Path:
    d = docx.Document()
    d.add_paragraph("ĐẢNG BỘ CATP. HỒ CHÍ MINH\tĐẢNG CỘNG SẢN VIỆT NAM")
    d.add_paragraph("ĐẢNG ỦY PC09")
    d.add_paragraph("THỐNG KÊ")
    d.add_paragraph("TÌNH HÌNH THỰC HIỆN NHIỆM VỤ KIỂM TRA, GIÁM SÁT")
    d.add_paragraph(tieu_de_ky)
    d.add_paragraph("(Kèm theo Báo cáo số -BC/ĐU)")
    bang = d.add_table(rows=0, cols=6)
    bang.add_row().cells[0].text = "TT"
    dem: dict[str, int] = {}
    for tt, nv in HANG_MAU:
        dem[tt] = dem.get(tt, 0) + 1
        khoa = tt if dem[tt] == 1 else f"{tt}#{dem[tt]}"
        o = bang.add_row().cells
        o[0].text, o[1].text, o[5].text = tt, nv, so.get(khoa, "")
    d.save(duong_dan)
    return duong_dan


def tao_bao_cao(duong_dan: Path, doan: list[str]) -> Path:
    d = docx.Document()
    for p in ["ĐẢNG BỘ CATP. HỒ CHÍ MINH\tĐẢNG CỘNG SẢN VIỆT NAM", "ĐẢNG ỦY PC09", "BÁO CÁO", *doan]:
        d.add_paragraph(p)
    d.save(duong_dan)
    return duong_dan


BAO_CAO_THANG_4 = [
    "Kết quả công tác kiểm tra, giám sát, kỷ luật đảng",
    "tháng 4 và phương hướng công tác tháng 5 năm 2025",
    "-----",
    "I. ĐẶC ĐIỂM TÌNH HÌNH",
    "Đảng bộ PC09 có 05 chi bộ trực thuộc; Tổng số đảng viên trong toàn Đảng bộ là 120 đồng chí.",
    "II. KẾT QUẢ CÔNG TÁC THÁNG 4 NĂM 2025",
    "2.1. Kiểm tra tổ chức đảng, đảng viên:",
    "a) Đảng ủy: Không.",
    "b) UBKT Đảng ủy:",
    "Triển khai Quyết định kiểm tra Chi bộ 3 khi có dấu hiệu vi phạm.",
    "2.3. Thi hành kỷ luật tổ chức đảng và đảng viên:",
    "a) Đảng ủy:",
    "- Thi hành kỷ luật Khiển trách đối với đồng chí Trần Văn An – Đảng viên Chi bộ 2 do vi phạm quy định về sinh hoạt đảng.",
    "- Báo cáo UBKT Đảng ủy CATP xem xét thi hành kỷ luật đối với đồng chí Lê Thị Bình – Đảng viên Chi bộ 5 do vi phạm pháp luật.",
    "b) Cấp ủy chi bộ trực thuộc:",
    "Chi bộ 1 thi hành kỷ luật bằng hình thức Cảnh cáo đối với đ/c Phạm Quốc Cường – Đảng viên Chi bộ 1 do vi phạm chính sách dân số.",
    "III. NHIỆM VỤ TRỌNG TÂM CÔNG TÁC THÁNG 5",
    "1. Tham mưu Đảng ủy xem xét, xử lý kỷ luật đ/c Lê Thị Bình – Đảng viên Chi bộ 5.",
    "2. Giám sát thực hiện quy trình tiếp nhận đơn thư./.",
]


@pytest.fixture
def thu_muc(tmp_path: Path) -> Path:
    tao_bieu_mau(tmp_path / "4.1 Thong ke.docx", "Tháng 4 năm 2025", {"d": "2", "da": "2", "dc": "2", "pđ": "1.500.000"})
    tao_bieu_mau(tmp_path / "5.1 Thong ke.docx", "Tháng 5 năm 2025", {"la": "1", "pb#2": "3"})
    tao_bao_cao(tmp_path / "4 BAO CAO thang 4.docx", BAO_CAO_THANG_4)
    return tmp_path


@pytest.mark.parametrize(
    "text, ky",
    [
        ("Tháng 5  năm 2025", Ky.thang(2025, 5)),
        ("kỷ luật đảng tháng 5 và phương hướng công tác tháng 6 năm 2025", Ky.thang(2025, 5)),
        ("6 tháng đầu năm và phương hướng công tác 6 tháng cuối năm 2025", Ky(2025, 1, 6)),
        ("Quý II năm 2025", Ky.quy(2025, 2)),
        ("9 tháng năm 2025", Ky(2025, 1, 9)),
        ("THỐNG KÊ TÌNH HÌNH THỰC HIỆN NHIỆM VỤ KIỂM TRA, GIÁM SÁT Năm 2025", Ky.ca_nam(2025)),
        ("không có năm", None),
    ],
)
def test_phan_tich_ky(text, ky):
    assert phan_tich_ky(text) == ky


def test_tham_so_ky():
    assert phan_tich_tham_so_ky("2025-06") == Ky.thang(2025, 6)
    assert phan_tich_tham_so_ky("2025-q2") == Ky.quy(2025, 2)
    assert phan_tich_tham_so_ky("2025-6T") == Ky(2025, 1, 6)
    assert phan_tich_tham_so_ky("2025") == Ky.ca_nam(2025)
    with pytest.raises(ValueError):
        phan_tich_tham_so_ky("quy 2")


def test_doc_bieu_mau(thu_muc):
    bm = doc_bieu_mau(thu_muc / "4.1 Thong ke.docx")
    assert bm.ky == Ky.thang(2025, 4) and bm.don_vi == "PC09"
    hang = bm.theo_khoa()
    assert hang["da"].gia_tri == [None, None, None, 2]
    assert hang["pđ"].gia_tri[3] == 1500000
    assert hang["d"].la_nhom and not hang["da"].la_nhom
    assert "pb#2" in hang  # mã trùng trong mẫu được tách riêng


def test_doc_bao_cao(thu_muc):
    bc = doc_bao_cao(thu_muc / "4 BAO CAO thang 4.docx")
    assert bc.ky == Ky.thang(2025, 4)
    assert (bc.so_chi_bo, bc.so_dang_vien) == (5, 120)
    kl = {(k.ho_ten, k.phan): k for k in bc.ky_luat}
    an = kl[("Trần Văn An", "Kết quả")]
    assert (an.chi_bo, an.hinh_thuc, an.trang_thai, an.chu_the) == ("Chi bộ 2", "Khiển trách", DA_THI_HANH, "Đảng ủy")
    assert an.ly_do == "vi phạm quy định về sinh hoạt đảng"
    assert kl[("Lê Thị Bình", "Kết quả")].trang_thai == DE_NGHI
    assert kl[("Lê Thị Bình", "Phương hướng")].trang_thai == DU_KIEN
    cuong = kl[("Phạm Quốc Cường", "Kết quả")]
    assert (cuong.hinh_thuc, cuong.trang_thai, cuong.chu_the) == ("Cảnh cáo", DA_THI_HANH, "Cấp ủy chi bộ")
    kt = [h for h in bc.hoat_dong if h.loai == "Kiểm tra"]
    assert [(h.chu_the, h.doi_tuong, h.tinh_trang) for h in kt] == [
        ("Đảng ủy", None, "Không"),
        ("UBKT Đảng ủy", "Chi bộ 3", "Đang thực hiện"),
    ]
    assert bc.phuong_huong[-1] == "Giám sát thực hiện quy trình tiếp nhận đơn thư."


def test_tong_hop_quy_cong_cac_thang(thu_muc):
    kho = nap_du_lieu([thu_muc])
    kq = tong_hop_ky(kho, Ky.quy(2025, 2))
    assert kq.so_lieu.thang_thieu == [6]
    assert kq.so_lieu.gia_tri["da"][3] == 2 and kq.so_lieu.gia_tri["la"][3] == 1
    assert kq.so_lieu.gia_tri["pb#2"][3] == 3
    assert kq.so_lieu.gia_tri["xa"][3] is None
    # báo cáo tháng 4: 2 người đã thi hành; biểu: da + la = 3 -> cảnh báo đối chiếu
    assert kq.so_da_thi_hanh == 2
    assert any("đối chiếu" in c for c in kq.canh_bao)
    # Lê Thị Bình gộp một dòng, giữ trạng thái cao nhất
    binh = [d for d in kq.ky_luat if d.ho_ten == "Lê Thị Bình"]
    assert len(binh) == 1 and binh[0].trang_thai == DE_NGHI and len(binh[0].dien_bien) == 2


def test_tong_hop_thang_khong_lay_thang_khac(thu_muc):
    kq = tong_hop_ky(nap_du_lieu([thu_muc]), Ky.thang(2025, 5))
    assert kq.bao_cao == [] and kq.so_lieu.gia_tri["da"][3] is None


def test_dong_lenh_xuat_excel(thu_muc, tmp_path, capsys):
    ra = tmp_path / "kq.xlsx"
    assert main(["tong-hop", str(thu_muc), "--ky", "2025-04", "--ky", "2025-Q2", "-o", str(ra)]) == 0
    wb = openpyxl.load_workbook(ra)
    assert wb.sheetnames == ["Tổng quan", "Số liệu", "Kỷ luật", "Kiểm tra - Giám sát",
                             "Hoạt động khác", "Phương hướng", "Cảnh báo"]
    so_lieu = {r[0]: r for r in wb["Số liệu"].iter_rows(min_row=2, values_only=True)}
    assert so_lieu["da"][2:] == (2, 2)
    ten = [r[2] for r in wb["Kỷ luật"].iter_rows(min_row=2, values_only=True)]
    assert "Trần Văn An" in ten


def test_bo_ma_truong():
    assert _bo_ma_truong("Trang \x13 PAGE \x141\x15 xong\x01") == "Trang 1 xong"
