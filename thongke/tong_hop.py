"""Tổng hợp số liệu và danh sách theo tháng / quý / năm từ một thư mục file Word."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from .bao_cao import DA_THI_HANH, MUC_TRANG_THAI, BaoCao, DangVienKyLuat, HoatDong, doc_bao_cao, la_bao_cao
from .bieu_mau import CAC_CAP, BieuMau, doc_bieu_mau, la_bieu_mau
from .doc_reader import doc_file
from .ky import Ky

DUOI_WORD = {".doc", ".docx"}


@dataclass
class KhoDuLieu:
    bieu_mau: list[BieuMau] = field(default_factory=list)
    bao_cao: list[BaoCao] = field(default_factory=list)
    canh_bao: list[str] = field(default_factory=list)

    @property
    def cac_ky(self) -> list[Ky]:
        return sorted({x.ky for x in [*self.bieu_mau, *self.bao_cao] if x.ky})


def nap_du_lieu(nguon: list[str | Path]) -> KhoDuLieu:
    kho = KhoDuLieu()
    files: list[Path] = []
    for n in map(Path, nguon):
        if n.is_dir():
            files += sorted(p for p in n.rglob("*") if p.suffix.lower() in DUOI_WORD)
        else:
            files.append(n)
    for f in files:
        if f.name.startswith("~$"):  # file tạm khi Word đang mở
            continue
        try:
            tl = doc_file(f)
        except Exception as e:  # noqa: BLE001 - báo lỗi từng file, không dừng cả lô
            kho.canh_bao.append(f"{f.name}: không đọc được ({e}).")
            continue
        if la_bieu_mau(tl):
            kho.bieu_mau.append(doc_bieu_mau(tl))
        elif la_bao_cao(tl):
            kho.bao_cao.append(doc_bao_cao(tl))
        else:
            kho.canh_bao.append(f"{f.name}: không phải biểu thống kê hay báo cáo, bỏ qua.")
    for x in [*kho.bieu_mau, *kho.bao_cao]:
        kho.canh_bao += [f"{x.duong_dan.name}: {c}" for c in x.canh_bao]
    return kho


# ---------------------------------------------------------------- số liệu


@dataclass
class SoLieuKy:
    ky: Ky
    # khóa hàng -> giá trị theo CAC_CAP (None = không có dữ liệu)
    gia_tri: dict[str, list[float | None]]
    nguon: list[str]
    thang_thieu: list[int]
    ghi_chu: list[str] = field(default_factory=list)


def tong_hop_so_lieu(kho: KhoDuLieu, ky: Ky) -> SoLieuKy:
    """Nếu có biểu mẫu đúng bằng kỳ cần tổng hợp thì dùng luôn;
    nếu không thì cộng các biểu mẫu tháng nằm trong kỳ."""
    ghi_chu: list[str] = []
    dung_ky = [b for b in kho.bieu_mau if b.ky == ky]
    if dung_ky:
        chon = [_moi_nhat(dung_ky, ghi_chu)]
        thieu: list[int] = []
    else:
        chon, thieu = [], []
        for t in ky.cac_thang:
            cua_thang = [b for b in kho.bieu_mau if b.ky == Ky.thang(ky.nam, t)]
            if cua_thang:
                chon.append(_moi_nhat(cua_thang, ghi_chu))
            else:
                thieu.append(t)

    kq = SoLieuKy(ky, {}, [b.duong_dan.name for b in chon], thieu, ghi_chu)
    for b in chon:
        for h in b.hang:
            dong = kq.gia_tri.setdefault(h.khoa, [None] * len(CAC_CAP))
            for i, v in enumerate(h.gia_tri):
                if v is not None:
                    dong[i] = (dong[i] or 0) + v
    if len(ky.cac_thang) > 1 and thieu and chon:
        kq.ghi_chu.append(
            f"{ky.ten}: thiếu biểu thống kê tháng {', '.join(map(str, thieu))} — số liệu chưa đầy đủ."
        )
    elif not chon:
        kq.ghi_chu.append(f"{ky.ten}: không có biểu thống kê nào.")
    return kq


def _moi_nhat(ds: list[BieuMau], ghi_chu: list[str]) -> BieuMau:
    """Nhiều file cùng kỳ: lấy file sửa gần nhất và ghi chú lại."""
    chon = max(ds, key=lambda b: b.duong_dan.stat().st_mtime)
    if len(ds) > 1:
        ghi_chu.append(
            f"{chon.ky.ten}: có {len(ds)} biểu thống kê cùng kỳ, dùng file sửa gần nhất '{chon.duong_dan.name}'."
        )
    return chon


def mau_hang(kho: KhoDuLieu) -> list[tuple[str, str, str, bool]]:
    """Danh sách hàng chuẩn (khóa, TT, nhiệm vụ, là nhóm) lấy từ biểu mẫu đầy đủ nhất."""
    if not kho.bieu_mau:
        return []
    mau = max(kho.bieu_mau, key=lambda b: len(b.hang))
    return [(h.khoa, h.tt, h.nhiem_vu, h.la_nhom) for h in mau.hang]


# ---------------------------------------------------------------- danh sách


@dataclass
class DongKyLuat:
    ky: Ky
    ho_ten: str
    chi_bo: str | None
    hinh_thuc: str | None
    trang_thai: str
    chu_the: str | None
    ly_do: str | None
    dien_bien: list[str]
    nguon: list[str]


def cac_bao_cao_trong_ky(kho: KhoDuLieu, ky: Ky, lay_ca_giao_nhau: bool) -> tuple[list[BaoCao], list[str]]:
    chon, ghi_chu = [], []
    for bc in sorted(kho.bao_cao, key=lambda b: _thu_tu(b.ky)):
        if not bc.ky:
            continue
        if ky.chua(bc.ky):
            chon.append(bc)
        elif ky.giao(bc.ky):
            if lay_ca_giao_nhau:
                chon.append(bc)
                ghi_chu.append(f"{ky.ten}: có dùng báo cáo '{bc.ky.ten}' (rộng hơn kỳ) — sự kiện có thể nằm ngoài kỳ.")
            else:
                ghi_chu.append(
                    f"{ky.ten}: bỏ qua báo cáo '{bc.ky.ten}' vì rộng hơn kỳ "
                    "(dùng --lay-ca-bao-cao-giao-nhau để tính cả)."
                )
    return chon, ghi_chu


def _thu_tu(ky: Ky | None) -> tuple[int, int, int]:
    """Sắp theo thời điểm kết thúc kỳ: báo cáo kết thúc muộn hơn là mới hơn."""
    return (ky.nam, ky.thang_cuoi, ky.thang_dau) if ky else (0, 0, 0)


def gop_ky_luat(ky: Ky, cac_bc: list[BaoCao]) -> list[DongKyLuat]:
    """Mỗi đảng viên một dòng; trạng thái lấy mức cao nhất, giữ lại diễn biến qua các báo cáo."""
    theo_ten: dict[str, DongKyLuat] = {}
    for bc in cac_bc:
        for kl in bc.ky_luat:
            d = theo_ten.get(kl.ho_ten)
            buoc = f"{bc.ky.ten}: {kl.trang_thai}" + (f" ({kl.hinh_thuc})" if kl.hinh_thuc else "")
            if d is None:
                theo_ten[kl.ho_ten] = DongKyLuat(
                    ky, kl.ho_ten, kl.chi_bo, kl.hinh_thuc, kl.trang_thai, kl.chu_the,
                    kl.ly_do, [buoc], [bc.duong_dan.name],
                )
                continue
            _cap_nhat(d, kl, bc)
            if buoc not in d.dien_bien:
                d.dien_bien.append(buoc)
            if bc.duong_dan.name not in d.nguon:
                d.nguon.append(bc.duong_dan.name)
    return sorted(theo_ten.values(), key=lambda d: (-MUC_TRANG_THAI[d.trang_thai], d.chi_bo or "", d.ho_ten))


def _cap_nhat(d: DongKyLuat, kl: DangVienKyLuat, bc: BaoCao) -> None:
    d.chi_bo = d.chi_bo or kl.chi_bo
    d.ly_do = d.ly_do or kl.ly_do
    if MUC_TRANG_THAI[kl.trang_thai] >= MUC_TRANG_THAI[d.trang_thai]:
        d.trang_thai = kl.trang_thai
        d.hinh_thuc = kl.hinh_thuc or d.hinh_thuc
        d.chu_the = kl.chu_the or d.chu_the


@dataclass
class KetQuaKy:
    ky: Ky
    so_lieu: SoLieuKy
    bao_cao: list[BaoCao]
    ky_luat: list[DongKyLuat]
    hoat_dong: list[tuple[BaoCao, HoatDong]]
    phuong_huong: list[tuple[BaoCao, int, str]]
    canh_bao: list[str]

    @property
    def so_da_thi_hanh(self) -> int:
        return sum(1 for d in self.ky_luat if d.trang_thai == DA_THI_HANH)


def tong_hop_ky(kho: KhoDuLieu, ky: Ky, lay_ca_giao_nhau: bool = False) -> KetQuaKy:
    so_lieu = tong_hop_so_lieu(kho, ky)
    cac_bc, ghi_chu = cac_bao_cao_trong_ky(kho, ky, lay_ca_giao_nhau)
    kq = KetQuaKy(
        ky,
        so_lieu,
        cac_bc,
        gop_ky_luat(ky, cac_bc),
        [(bc, hd) for bc in cac_bc for hd in bc.hoat_dong],
        [(bc, i + 1, nd) for bc in cac_bc for i, nd in enumerate(bc.phuong_huong)],
        so_lieu.ghi_chu + ghi_chu,
    )
    kq.canh_bao += doi_chieu(kq)
    return kq


def doi_chieu(kq: KetQuaKy) -> list[str]:
    """So số đảng viên đã bị kỷ luật trong báo cáo lời văn với biểu thống kê (hàng da + la)."""
    if not kq.bao_cao or not kq.so_lieu.nguon:
        return []
    tong = 0.0
    for khoa in ("da", "la"):
        tong += sum(v or 0 for v in kq.so_lieu.gia_tri.get(khoa, []))
    so_bc = kq.so_da_thi_hanh
    if so_bc != tong:
        ten = ", ".join(d.ho_ten for d in kq.ky_luat if d.trang_thai == DA_THI_HANH) or "không có"
        return [
            f"{kq.ky.ten}: báo cáo ghi {so_bc} đảng viên đã bị thi hành kỷ luật ({ten}) "
            f"nhưng biểu thống kê (hàng da + la) ghi {tong:g}. Cần đối chiếu lại."
        ]
    return []
