"""Xuất kết quả tổng hợp ra file Excel."""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .bieu_mau import CAC_CAP
from .tong_hop import KetQuaKy, KhoDuLieu, mau_hang

DAM = Font(bold=True)
NEN_TIEU_DE = PatternFill("solid", fgColor="DDEBF7")
NEN_NHOM = PatternFill("solid", fgColor="F2F2F2")
VIEN = Border(*(Side(style="thin", color="BFBFBF"),) * 4)
XUONG_DONG = Alignment(wrap_text=True, vertical="top")


def xuat(kho: KhoDuLieu, cac_kq: list[KetQuaKy], duong_dan: str | Path) -> Path:
    wb = Workbook()
    _tong_quan(wb.active, cac_kq)
    _so_lieu(wb.create_sheet("Số liệu"), kho, cac_kq)
    _ky_luat(wb.create_sheet("Kỷ luật"), cac_kq)
    _hoat_dong(wb.create_sheet("Kiểm tra - Giám sát"), cac_kq, khac=False)
    _hoat_dong(wb.create_sheet("Hoạt động khác"), cac_kq, khac=True)
    _phuong_huong(wb.create_sheet("Phương hướng"), cac_kq)
    _canh_bao(wb.create_sheet("Cảnh báo"), kho, cac_kq)
    duong_dan = Path(duong_dan)
    duong_dan.parent.mkdir(parents=True, exist_ok=True)
    wb.save(duong_dan)
    return duong_dan


def _bang(ws, tieu_de: list[str], hang: list[list], rong: list[int]) -> None:
    ws.append(tieu_de)
    for o in ws[1]:
        o.font, o.fill, o.border, o.alignment = DAM, NEN_TIEU_DE, VIEN, XUONG_DONG
    for h in hang:
        ws.append(h)
    for dong in ws.iter_rows(min_row=2):
        for o in dong:
            o.border, o.alignment = VIEN, XUONG_DONG
    for i, r in enumerate(rong, 1):
        ws.column_dimensions[get_column_letter(i)].width = r
    ws.freeze_panes = "A2"
    if hang:
        ws.auto_filter.ref = ws.dimensions


def _tong_quan(ws, cac_kq: list[KetQuaKy]) -> None:
    ws.title = "Tổng quan"
    hang = []
    for kq in cac_kq:
        bc_moi = kq.bao_cao[-1] if kq.bao_cao else None
        hang.append([
            kq.ky.ten,
            bc_moi.don_vi if bc_moi else None,
            bc_moi.so_chi_bo if bc_moi else None,
            bc_moi.so_dang_vien if bc_moi else None,
            "\n".join(kq.so_lieu.nguon) or "(không có)",
            ", ".join(map(str, kq.so_lieu.thang_thieu)) or None,
            "\n".join(b.duong_dan.name for b in kq.bao_cao) or "(không có)",
            kq.so_da_thi_hanh,
            sum(1 for d in kq.ky_luat if d.trang_thai != "Đã thi hành"),
            len(kq.canh_bao),
        ])
    _bang(
        ws,
        ["Kỳ", "Đơn vị", "Số chi bộ", "Số đảng viên", "Biểu thống kê dùng", "Tháng thiếu biểu",
         "Báo cáo dùng", "Đảng viên đã thi hành KL", "Đảng viên đang đề nghị / dự kiến KL", "Số cảnh báo"],
        hang,
        [22, 10, 10, 12, 40, 14, 40, 14, 16, 10],
    )


def _so_lieu(ws, kho: KhoDuLieu, cac_kq: list[KetQuaKy]) -> None:
    # Chỉ hiện các cấp có dữ liệu; luôn hiện Cấp cơ sở.
    cap_hien = [
        i for i, _ in enumerate(CAC_CAP)
        if i == 3 or any(kq.so_lieu.gia_tri.get(k, [None] * 4)[i] is not None
                         for kq in cac_kq for k in kq.so_lieu.gia_tri)
    ]
    tieu_de = ["TT", "Nhiệm vụ"]
    for kq in cac_kq:
        for i in cap_hien:
            tieu_de.append(f"{kq.ky.ten}\n{CAC_CAP[i]}" if len(cap_hien) > 1 else kq.ky.ten)
    hang, la_nhom = [], []
    for khoa, tt, nhiem_vu, nhom in mau_hang(kho):
        dong = [tt, nhiem_vu]
        for kq in cac_kq:
            gt = kq.so_lieu.gia_tri.get(khoa)
            dong += [(gt[i] if gt else None) for i in cap_hien]
        hang.append(dong)
        la_nhom.append(nhom)
    _bang(ws, tieu_de, hang, [6, 60] + [16] * (len(tieu_de) - 2))
    for r, nhom in enumerate(la_nhom, start=2):
        if nhom:
            for o in ws[r]:
                o.font, o.fill = DAM, NEN_NHOM
    ws.freeze_panes = "C2"


def _ky_luat(ws, cac_kq: list[KetQuaKy]) -> None:
    hang = [
        [kq.ky.ten, i, d.ho_ten, d.chi_bo, d.hinh_thuc, d.trang_thai, d.chu_the, d.ly_do,
         "\n".join(d.dien_bien), "\n".join(d.nguon)]
        for kq in cac_kq
        for i, d in enumerate(kq.ky_luat, 1)
    ]
    _bang(
        ws,
        ["Kỳ", "STT", "Họ tên", "Chi bộ", "Hình thức", "Trạng thái", "Cấp xử lý", "Lý do",
         "Diễn biến qua các báo cáo", "Nguồn"],
        hang,
        [20, 5, 22, 10, 12, 20, 16, 40, 45, 30],
    )


def _hoat_dong(ws, cac_kq: list[KetQuaKy], khac: bool) -> None:
    hang = []
    for kq in cac_kq:
        for bc, hd in kq.hoat_dong:
            if (hd.loai == "Khác") != khac:
                continue
            if khac:
                hang.append([kq.ky.ten, bc.ky.ten, hd.noi_dung, bc.duong_dan.name])
            else:
                hang.append([kq.ky.ten, bc.ky.ten, hd.loai, hd.chu_the, hd.doi_tuong, hd.tinh_trang,
                             hd.noi_dung, bc.duong_dan.name])
    if khac:
        _bang(ws, ["Kỳ", "Kỳ báo cáo", "Nội dung", "Nguồn"], hang, [20, 22, 100, 30])
    else:
        _bang(
            ws,
            ["Kỳ", "Kỳ báo cáo", "Loại", "Chủ thể", "Đối tượng", "Tình trạng", "Nội dung", "Nguồn"],
            hang,
            [20, 22, 14, 16, 14, 16, 80, 30],
        )


def _phuong_huong(ws, cac_kq: list[KetQuaKy]) -> None:
    hang = [[kq.ky.ten, bc.ky.ten, stt, nd, bc.duong_dan.name]
            for kq in cac_kq for bc, stt, nd in kq.phuong_huong]
    _bang(ws, ["Kỳ", "Báo cáo kỳ", "STT", "Nhiệm vụ kỳ tới", "Nguồn"], hang, [20, 22, 5, 100, 30])


def _canh_bao(ws, kho: KhoDuLieu, cac_kq: list[KetQuaKy]) -> None:
    hang = [["Chung", c] for c in kho.canh_bao]
    hang += [[kq.ky.ten, c] for kq in cac_kq for c in kq.canh_bao]
    _bang(ws, ["Kỳ", "Nội dung"], hang, [22, 120])
