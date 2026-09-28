"""Dòng lệnh.

Ví dụ:
  python -m thongke tong-hop du_lieu --ky 2025-05 --ky 2025-Q2 --ky 2025
  python -m thongke tong-hop du_lieu --theo quy --nam 2025
  python -m thongke xem "du_lieu/6_BAO CAO 6 tháng 2025.doc"
"""

from __future__ import annotations

import argparse
import json
import sys
from dataclasses import asdict, is_dataclass
from pathlib import Path

from .ky import Ky, phan_tich_tham_so_ky
from .tong_hop import KhoDuLieu, nap_du_lieu, tong_hop_ky


def _cac_ky_theo(kho: KhoDuLieu, theo: str, nam: int | None) -> list[Ky]:
    cac_nam = sorted({k.nam for k in kho.cac_ky if nam is None or k.nam == nam})
    if not cac_nam and nam:
        cac_nam = [nam]
    ra: list[Ky] = []
    for n in cac_nam:
        if theo == "thang":
            ra += [Ky.thang(n, t) for t in range(1, 13)]
        elif theo == "quy":
            ra += [Ky.quy(n, q) for q in range(1, 5)]
        else:
            ra.append(Ky.ca_nam(n))
    # bỏ các kỳ không có dữ liệu nào
    return [k for k in ra if any(k.giao(x) for x in kho.cac_ky)]


def lenh_tong_hop(a: argparse.Namespace) -> int:
    from .xuat_excel import xuat

    kho = nap_du_lieu(a.nguon)
    print(f"Đã đọc {len(kho.bieu_mau)} biểu thống kê, {len(kho.bao_cao)} báo cáo.")
    for x in [*kho.bieu_mau, *kho.bao_cao]:
        loai = "Biểu TK " if x in kho.bieu_mau else "Báo cáo "
        print(f"  {loai} {x.ky.ten if x.ky else '??':<26} {x.duong_dan.name}")

    cac_ky = [phan_tich_tham_so_ky(k) for k in a.ky or []]
    if a.theo:
        cac_ky += _cac_ky_theo(kho, a.theo, a.nam)
    if not cac_ky:
        print("Chưa chọn kỳ: dùng --ky (vd 2025-05, 2025-Q2, 2025) hoặc --theo thang|quy|nam.", file=sys.stderr)
        return 2

    cac_kq = [tong_hop_ky(kho, k, a.lay_ca_bao_cao_giao_nhau) for k in cac_ky]
    ra = Path(a.ra or f"TongHop_{'_'.join(k.ma for k in cac_ky[:3])}{'_...' if len(cac_ky) > 3 else ''}.xlsx")
    xuat(kho, cac_kq, ra)
    print(f"\nĐã xuất: {ra}")
    for kq in cac_kq:
        print(f"  {kq.ky.ten}: {kq.so_da_thi_hanh} đảng viên đã thi hành KL, "
              f"{len(kq.ky_luat) - kq.so_da_thi_hanh} đang đề nghị/dự kiến, "
              f"{len(kq.hoat_dong)} hoạt động, {len(kq.canh_bao)} cảnh báo")
    so_cb = len(kho.canh_bao) + sum(len(kq.canh_bao) for kq in cac_kq)
    if so_cb:
        print(f"\nCó {so_cb} cảnh báo — xem sheet 'Cảnh báo'.")
    return 0


def _sang_dict(o):
    if is_dataclass(o):
        return {k: _sang_dict(v) for k, v in asdict(o).items()}
    if isinstance(o, Path):
        return str(o)
    if isinstance(o, dict):
        return {k: _sang_dict(v) for k, v in o.items()}
    if isinstance(o, (list, tuple)):
        return [_sang_dict(v) for v in o]
    return o


def lenh_xem(a: argparse.Namespace) -> int:
    kho = nap_du_lieu(a.file)
    for x in [*kho.bieu_mau, *kho.bao_cao]:
        d = _sang_dict(x)
        if "hang" in d:  # biểu mẫu: chỉ in hàng có số liệu khác 0
            d["hang"] = [h for h in d["hang"] if any(v not in (None, 0) for v in h["gia_tri"])]
        print(json.dumps(d, ensure_ascii=False, indent=2, default=str))
    for c in kho.canh_bao:
        print("CẢNH BÁO:", c, file=sys.stderr)
    return 0


def main(argv: list[str] | None = None) -> int:
    for s in (sys.stdout, sys.stderr):
        if hasattr(s, "reconfigure"):
            s.reconfigure(encoding="utf-8")
    p = argparse.ArgumentParser(prog="thongke", description="Trích xuất số liệu, danh sách từ báo cáo KTGS.")
    sub = p.add_subparsers(dest="lenh", required=True)

    th = sub.add_parser("tong-hop", help="Tổng hợp theo tháng/quý/năm ra Excel")
    th.add_argument("nguon", nargs="+", help="Thư mục hoặc file .doc/.docx")
    th.add_argument("--ky", action="append", help="Kỳ: 2025-05 | 2025-Q2 | 2025-6T | 2025 (lặp lại được)")
    th.add_argument("--theo", choices=["thang", "quy", "nam"], help="Tạo mọi kỳ có dữ liệu theo tháng/quý/năm")
    th.add_argument("--nam", type=int, help="Giới hạn --theo trong một năm")
    th.add_argument("--lay-ca-bao-cao-giao-nhau", action="store_true",
                    help="Tính cả báo cáo có kỳ rộng hơn kỳ tổng hợp (vd báo cáo 6 tháng khi tổng hợp quý II)")
    th.add_argument("-o", "--ra", help="File Excel đầu ra")
    th.set_defaults(ham=lenh_tong_hop)

    xem = sub.add_parser("xem", help="In nội dung đã trích của từng file (JSON) để kiểm tra")
    xem.add_argument("file", nargs="+")
    xem.set_defaults(ham=lenh_xem)

    a = p.parse_args(argv)
    try:
        return a.ham(a)
    except ValueError as e:
        print(f"Lỗi: {e}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
