"""Đọc file "Báo cáo" lời văn (tháng / 6 tháng / năm) và trích danh sách.

Trích ra:
- Thông tin chung: số chi bộ, tổng số đảng viên.
- Danh sách đảng viên liên quan kỷ luật (họ tên, chi bộ, hình thức, trạng thái, lý do).
- Danh sách hoạt động kiểm tra / giám sát / đơn thư (theo chủ thể: Đảng ủy, chi bộ, UBKT).
- Các hoạt động khác trong phần kết quả.
- Nhiệm vụ trọng tâm / phương hướng kỳ tới.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

from .bieu_mau import tim_don_vi
from .doc_reader import TaiLieu, doc_file
from .ky import Ky, chuan_hoa, phan_tich_ky

HINH_THUC = ["Khiển trách", "Cảnh cáo", "Cách chức", "Khai trừ", "Giải tán"]

DA_THI_HANH = "Đã thi hành"
DE_NGHI = "Đề nghị / đang xem xét"
DU_KIEN = "Dự kiến (phương hướng)"
MUC_TRANG_THAI = {DU_KIEN: 0, DE_NGHI: 1, DA_THI_HANH: 2}

LOAI_KIEM_TRA = "Kiểm tra"
LOAI_GIAM_SAT = "Giám sát"
LOAI_KY_LUAT = "Kỷ luật"
LOAI_DON_THU = "Đơn thư, tố cáo, khiếu nại"


@dataclass
class DangVienKyLuat:
    ho_ten: str
    chi_bo: str | None
    hinh_thuc: str | None
    trang_thai: str
    chu_the: str | None  # Đảng ủy / Cấp ủy chi bộ / UBKT ...
    ly_do: str | None
    noi_dung: str  # đoạn văn gốc
    phan: str  # "Kết quả" | "Phương hướng"


@dataclass
class HoatDong:
    loai: str  # Kiểm tra / Giám sát / Kỷ luật / Đơn thư ... / Khác
    chu_the: str | None
    doi_tuong: str | None
    noi_dung: str
    tinh_trang: str | None


@dataclass
class BaoCao:
    duong_dan: Path
    ky: Ky | None
    don_vi: str | None
    tieu_de: str
    so_chi_bo: int | None = None
    so_dang_vien: int | None = None
    ky_luat: list[DangVienKyLuat] = field(default_factory=list)
    hoat_dong: list[HoatDong] = field(default_factory=list)
    phuong_huong: list[str] = field(default_factory=list)
    canh_bao: list[str] = field(default_factory=list)


# ---------------------------------------------------------------- nhận dạng tiêu đề mục

RE_MUC_CHU = re.compile(r"^([A-D])\.\s+(.*)$")  # A. B.
RE_MUC_LA_MA = re.compile(r"^(I|II|III|IV|V|VI|VII)\.\s+(.*)$")
RE_MUC_SO = re.compile(r"^(\d+(?:\.\d+)*)\.?\s+(.*)$")  # 1. / 1.2. / 2.3.
RE_MUC_NHO = re.compile(r"^([a-eđ])[.)]\s+(.*)$")  # a. / b)
RE_GACH = re.compile(r"^[-–+•]\s*")


def _loai_tu_tieu_de(text: str) -> str | None:
    t = text.lower()
    if "kỷ luật" in t and "thi hành" in t:
        return LOAI_KY_LUAT
    if "đơn thư" in t or "tố cáo" in t or "khiếu nại" in t:
        return LOAI_DON_THU
    if "giám sát" in t and "kiểm tra" not in t:
        return LOAI_GIAM_SAT
    if "kiểm tra" in t and "giám sát" not in t:
        return LOAI_KIEM_TRA
    return None


def _chu_the_tu_tieu_de(text: str) -> str | None:
    t = text.lower()
    if t.startswith("ubkt") or t.startswith("ủy ban kiểm tra"):
        return "UBKT Đảng ủy"
    if t.startswith("cấp ủy chi bộ") or t.startswith("chi bộ") or t.startswith("các chi bộ"):
        return "Cấp ủy chi bộ"
    if t.startswith("đảng ủy"):
        return "Đảng ủy"
    return None


def _tach_noi_dung_sau_tieu_de(text: str) -> tuple[str, str | None]:
    """'Kiểm tra tổ chức đảng, đảng viên: Không thực hiện.' -> (tiêu đề, 'Không thực hiện')."""
    if ":" in text:
        dau, sau = text.split(":", 1)
        sau = sau.strip().rstrip(".").strip()
        return dau.strip(), (sau or None)
    # 'd. Tiếp nhận, xử lý đơn thư, ... đảng viên. Không có.'
    m = re.match(r"^(.*?)\.\s+(Không[^.]*|Đang thực hiện[^.]*)\.?$", text)
    if m:
        return m.group(1).strip(), m.group(2).strip()
    return text.rstrip(":.").strip(), None


# ---------------------------------------------------------------- nhận dạng họ tên

RE_DANH_XUNG = re.compile(r"(?:đồng chí|đ/c)\s+", re.IGNORECASE)
RE_TU = re.compile(r"[^\W\d_]+")


def _tim_ho_ten(doan: str) -> list[tuple[int, int, str]]:
    """Trả về [(vị trí đầu, vị trí cuối, họ tên)] cho các cụm 'đồng chí/đ/c Họ Tên'."""
    ket_qua = []
    for m in RE_DANH_XUNG.finditer(doan):
        vt = m.end()
        cac_tu = []
        while len(cac_tu) < 5:
            t = RE_TU.match(doan, vt)
            if not t or not t.group(0)[0].isupper():
                break
            cac_tu.append(t.group(0))
            vt = t.end()
            khoang = re.match(r"[  ]+", doan[vt:])
            if not khoang:
                break
            vt += khoang.end()
        if len(cac_tu) >= 2:
            ten = " ".join(cac_tu)
            ket_qua.append((m.start(), m.start() + len(m.group(0)) + len(ten), ten))
    return ket_qua


def _trang_thai(doan_con: str, phuong_huong: bool) -> str:
    if phuong_huong:
        return DU_KIEN
    t = doan_con.lower()
    if "quyết định thi hành kỷ luật" in t or "quyết định kỷ luật" in t:
        return DA_THI_HANH
    hinh = "|".join(h.lower() for h in HINH_THUC)
    for m in re.finditer(rf"thi hành kỷ luật\s+(?:bằng hình thức\s+)?(?:{hinh})", t):
        truoc = t[max(0, m.start() - 12) : m.start()]
        if "xem xét" not in truoc and "đề nghị" not in truoc:
            return DA_THI_HANH
    return DE_NGHI


def _hinh_thuc(doan_con: str) -> str | None:
    tim = [(doan_con.lower().find(h.lower()), h) for h in HINH_THUC]
    tim = [x for x in tim if x[0] >= 0]
    return min(tim)[1] if tim else None


def _ly_do(doan_con: str) -> str | None:
    m = re.search(r"\bdo\s+(.+?)(?:\.\s|\.$|;|$)", doan_con)
    return m.group(1).strip() if m else None


def _chi_bo_sau_ten(doan: str, vt: int) -> str | None:
    m = re.match(r"\s*[–\-—]\s*Đảng viên\s+(Chi bộ\s*\d+)", doan[vt:], re.IGNORECASE)
    if m:
        return re.sub(r"\s+", " ", m.group(1)).replace("chi bộ", "Chi bộ")
    return None


def trich_ky_luat(doan: str, phuong_huong: bool, chu_the: str | None) -> list[DangVienKyLuat]:
    ten_list = _tim_ho_ten(doan)
    # gộp các lần nhắc lại cùng một người trong cùng đoạn
    nguoi: list[tuple[int, int, str]] = []
    for dau, cuoi, ten in ten_list:
        if not any(n[2] == ten for n in nguoi):
            nguoi.append((dau, cuoi, ten))
    ket_qua = []
    for i, (dau, cuoi, ten) in enumerate(nguoi):
        # đoạn con: từ sau người trước (hoặc đầu đoạn) đến trước người sau (hoặc cuối đoạn)
        tu = nguoi[i - 1][1] if i > 0 else 0
        den = nguoi[i + 1][0] if i + 1 < len(nguoi) else len(doan)
        doan_con = doan[tu:den]
        ket_qua.append(
            DangVienKyLuat(
                ho_ten=ten,
                chi_bo=_chi_bo_sau_ten(doan, cuoi),
                hinh_thuc=_hinh_thuc(doan_con),
                trang_thai=_trang_thai(doan_con, phuong_huong),
                chu_the=chu_the,
                ly_do=_ly_do(doan[cuoi:den]),
                noi_dung=doan,
                phan="Phương hướng" if phuong_huong else "Kết quả",
            )
        )
    return ket_qua


# ---------------------------------------------------------------- đọc báo cáo


def la_bao_cao(tl: TaiLieu) -> bool:
    return any(chuan_hoa(d).upper() == "BÁO CÁO" for d in tl.doan_van[:15])


def _tieu_de(doan_van: list[str]) -> tuple[str, int]:
    """Trả về (tiêu đề, chỉ số đoạn đầu tiên của thân báo cáo)."""
    for i, d in enumerate(doan_van):
        if chuan_hoa(d).upper() == "BÁO CÁO":
            tieu_de = []
            j = i + 1
            while j < len(doan_van) and not doan_van[j].startswith("---") and len(tieu_de) < 5:
                tieu_de.append(chuan_hoa(doan_van[j]))
                j += 1
            if j < len(doan_van) and doan_van[j].startswith("---"):
                j += 1
            return " ".join(tieu_de), j
    return "", 0


def doc_bao_cao(nguon: str | Path | TaiLieu) -> BaoCao:
    tl = nguon if isinstance(nguon, TaiLieu) else doc_file(nguon)
    tieu_de, bat_dau = _tieu_de(tl.doan_van)
    ky = phan_tich_ky(tieu_de) or phan_tich_ky(tl.duong_dan.stem.replace("_", " "))
    bc = BaoCao(tl.duong_dan, ky, tim_don_vi(tl.doan_van), tieu_de)
    if ky is None:
        bc.canh_bao.append("Không xác định được kỳ báo cáo từ tiêu đề.")

    phuong_huong = False
    loai_cap_so: str | None = None  # loại theo mục số (1.2, 2.3...)
    loai_cap_nho: str | None = None  # loại theo mục a./b.
    chu_the: str | None = None

    for goc in tl.doan_van[bat_dau:]:
        d = chuan_hoa(goc)
        if d.startswith("Nơi nhận") or d.startswith("- UBKT Đảng ủy CATP") or d.startswith("|"):
            break
        if d.startswith("Trên đây là báo cáo") or d.endswith("như sau:"):
            continue

        if m := RE_MUC_CHU.match(d) or RE_MUC_LA_MA.match(d):
            tieu = m.group(2).upper()
            phuong_huong = "PHƯƠNG HƯỚNG" in tieu or "NHIỆM VỤ TRỌNG" in tieu
            loai_cap_so = loai_cap_nho = chu_the = None
            continue

        if phuong_huong:
            if m := RE_MUC_SO.match(d):
                bc.phuong_huong.append(_bo_ket(m.group(2)))
            elif d:
                bc.phuong_huong.append(_bo_ket(RE_GACH.sub("", d)))
            if "kỷ luật" in d.lower():
                bc.ky_luat.extend(trich_ky_luat(d, True, None))
            continue

        if m := RE_MUC_SO.match(d):
            tieu, noi_dung = _tach_noi_dung_sau_tieu_de(m.group(2))
            loai_cap_so = _loai_tu_tieu_de(tieu)
            loai_cap_nho = chu_the = None
            if noi_dung and loai_cap_so:
                bc.hoat_dong.append(HoatDong(loai_cap_so, None, None, noi_dung, noi_dung))
            continue

        if m := RE_MUC_NHO.match(d):
            tieu, noi_dung = _tach_noi_dung_sau_tieu_de(m.group(2))
            chu_the_moi = _chu_the_tu_tieu_de(tieu)
            if chu_the_moi:
                chu_the, loai_cap_nho = chu_the_moi, None
            else:
                loai_cap_nho, chu_the = _loai_tu_tieu_de(tieu), None
            if noi_dung:
                loai = loai_cap_nho or loai_cap_so or "Khác"
                bc.hoat_dong.append(HoatDong(loai, chu_the, None, noi_dung, noi_dung))
            continue

        # đoạn nội dung
        _thong_tin_chung(bc, d)
        loai = loai_cap_nho or loai_cap_so
        noi_dung = RE_GACH.sub("", d)
        if loai == LOAI_KY_LUAT or (
            "kỷ luật" in d.lower() and _tim_ho_ten(d) and loai != LOAI_DON_THU
        ):
            bc.ky_luat.extend(trich_ky_luat(d, False, chu_the))
            continue
        if bc.so_dang_vien is not None and "tổng số đảng viên" in d.lower():
            continue  # đoạn đặc điểm tình hình đã lấy số liệu
        bc.hoat_dong.append(
            HoatDong(loai or "Khác", chu_the, _doi_tuong(noi_dung), noi_dung, _tinh_trang(noi_dung))
        )

    _kiem_tra_don_vi(bc, tl.doan_van[bat_dau:])
    return bc


def _bo_ket(text: str) -> str:
    """Bỏ ký hiệu kết thúc văn bản './.'"""
    return re.sub(r"\s*\./\.$", ".", text)


def _thong_tin_chung(bc: BaoCao, d: str) -> None:
    if bc.so_chi_bo is None and (m := re.search(r"có\s+(\d+)\s+chi bộ", d, re.IGNORECASE)):
        bc.so_chi_bo = int(m.group(1))
    if bc.so_dang_vien is None and (
        m := re.search(r"tổng số đảng viên[^0-9]{0,40}?(\d[\d.]*)", d, re.IGNORECASE)
    ):
        bc.so_dang_vien = int(m.group(1).replace(".", ""))


def _doi_tuong(noi_dung: str) -> str | None:
    cac_chi_bo = sorted(set(re.findall(r"Chi bộ\s*\d+", noi_dung)), key=lambda s: int(s.split()[-1]))
    return ", ".join(cac_chi_bo) or None


def _tinh_trang(noi_dung: str) -> str | None:
    t = noi_dung.lower()
    if t.startswith("hoàn thành") or t.startswith("đã "):
        return "Hoàn thành"
    if t.startswith("triển khai") or t.startswith("tiếp tục") or "đang thực hiện" in t:
        return "Đang thực hiện"
    return None


def _kiem_tra_don_vi(bc: BaoCao, doan_van: list[str]) -> None:
    """Cảnh báo tên đơn vị khác lẫn trong báo cáo (thường do sao chép từ mẫu cũ)."""
    if not bc.don_vi:
        return
    ma = re.match(r"([A-ZĐ]+)\d+", bc.don_vi)
    if not ma:
        return
    lan: dict[str, int] = {}
    for d in doan_van:
        for m in re.finditer(rf"\b{ma.group(1)}\d+\b", d):
            if m.group(0) != bc.don_vi:
                lan[m.group(0)] = lan.get(m.group(0), 0) + 1
    for dv, n in lan.items():
        bc.canh_bao.append(
            f"Báo cáo của {bc.don_vi} nhưng nhắc tới '{dv}' {n} lần — kiểm tra lại có phải sót do sao chép mẫu."
        )
