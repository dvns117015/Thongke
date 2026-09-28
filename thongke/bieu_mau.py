"""Đọc file "Thống kê số liệu" (biểu mẫu tình hình thực hiện nhiệm vụ kiểm tra, giám sát).

Biểu mẫu là một bảng cố định: TT | Nhiệm vụ | Cấp Trung ương | Cấp tỉnh | Cấp huyện | Cấp cơ sở.
Một số mã TT trong mẫu bị trùng (pb, rb, rc) nên khóa của hàng là mã TT kèm số thứ tự
lần xuất hiện: 'pb', 'pb#2', ...
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

from .doc_reader import TaiLieu, doc_file
from .ky import Ky, chuan_hoa, phan_tich_ky

CAC_CAP = ["Cấp Trung ương", "Cấp tỉnh", "Cấp huyện", "Cấp cơ sở"]
MA_TT = re.compile(r"^(?:[IVX]+|[a-zđ]{1,2})$")
MA_NHOM = re.compile(r"^(?:[IVX]+|[a-zđ])$")  # hàng tổng nhóm: I, II, a, b, ...


@dataclass
class HangSoLieu:
    khoa: str  # mã TT duy nhất trong biểu (vd 'da', 'pb#2')
    tt: str  # mã TT như trong biểu
    nhiem_vu: str
    la_nhom: bool
    gia_tri: list[float | None]  # theo CAC_CAP; None = ô trống


@dataclass
class BieuMau:
    duong_dan: Path
    ky: Ky | None
    don_vi: str | None
    hang: list[HangSoLieu] = field(default_factory=list)
    canh_bao: list[str] = field(default_factory=list)

    def theo_khoa(self) -> dict[str, HangSoLieu]:
        return {h.khoa: h for h in self.hang}


def la_bieu_mau(tl: TaiLieu) -> bool:
    return any(_bang_so_lieu(b) for b in tl.bang)


def _bang_so_lieu(bang: list[list[str]]) -> bool:
    so_hang_ma = sum(1 for h in bang if len(h) >= 6 and MA_TT.match(h[0].strip()))
    return so_hang_ma >= 20


def doc_bieu_mau(nguon: str | Path | TaiLieu) -> BieuMau:
    tl = nguon if isinstance(nguon, TaiLieu) else doc_file(nguon)
    bm = BieuMau(tl.duong_dan, _tim_ky(tl), tim_don_vi(tl.doan_van))
    if bm.ky is None:
        bm.canh_bao.append("Không xác định được kỳ (tháng/quý/năm) của biểu mẫu.")

    bang = next((b for b in tl.bang if _bang_so_lieu(b)), None)
    if bang is None:
        if tl.duong_dan.suffix.lower() == ".doc" or not tl.bang:
            bm.canh_bao.append(
                "Không đọc được bảng số liệu. Nếu là file .doc, hãy mở bằng Word và "
                "'Save As' sang .docx rồi chạy lại."
            )
        return bm

    dem: dict[str, int] = {}
    for o in bang:
        if len(o) < 6:
            continue
        tt = o[0].strip()
        if not MA_TT.match(tt):
            continue
        dem[tt] = dem.get(tt, 0) + 1
        khoa = tt if dem[tt] == 1 else f"{tt}#{dem[tt]}"
        gia_tri = []
        for i, text in enumerate(o[-4:]):
            so, loi = _doc_so(text)
            if loi:
                bm.canh_bao.append(f"Hàng {tt} ({CAC_CAP[i]}): không đọc được số '{text}'.")
            gia_tri.append(so)
        bm.hang.append(
            HangSoLieu(khoa, tt, chuan_hoa(o[1]), bool(MA_NHOM.match(tt)), gia_tri)
        )
    return bm


def _doc_so(text: str) -> tuple[float | None, bool]:
    t = re.sub(r"[\s/]", "", text)
    if not t or t in {"-", "x", "X"}:
        return None, False
    # 1.500.000 hoặc 1,500,000 là số tiền; 1,5 là số thập phân
    if re.fullmatch(r"\d{1,3}([.,]\d{3})+", t):
        t = re.sub(r"[.,]", "", t)
    t = t.replace(",", ".")
    try:
        so = float(t)
    except ValueError:
        return None, True
    return (int(so) if so.is_integer() else so), False


def _tim_ky(tl: TaiLieu) -> Ky | None:
    # Tiêu đề kỳ nằm sau dòng "THỐNG KÊ", trước dòng "(Kèm theo ...)"
    dv = tl.doan_van
    bat_dau = next((i for i, d in enumerate(dv) if chuan_hoa(d).upper().startswith("THỐNG KÊ")), None)
    if bat_dau is not None:
        tieu_de = []
        for d in dv[bat_dau : bat_dau + 6]:
            if d.lstrip().startswith("(") or d.startswith("---"):
                break
            tieu_de.append(d)
        ky = phan_tich_ky(" ".join(tieu_de))
        if ky:
            return ky
    return phan_tich_ky(tl.duong_dan.stem.replace("_", " "))


def tim_don_vi(doan_van: list[str]) -> str | None:
    for d in doan_van[:6]:
        m = re.search(r"ĐẢNG ỦY\s+([A-ZĐ0-9][\w\-]*)", chuan_hoa(d))
        if m:
            return m.group(1)
    return None
