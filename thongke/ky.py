"""Kỳ báo cáo (tháng / quý / 6 tháng / 9 tháng / năm) và cách nhận diện từ văn bản."""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass

LA_MA = {"I": 1, "II": 2, "III": 3, "IV": 4}


@dataclass(frozen=True, order=True)
class Ky:
    """Một khoảng tháng liên tục trong một năm: [thang_dau, thang_cuoi]."""

    nam: int
    thang_dau: int
    thang_cuoi: int

    @classmethod
    def thang(cls, nam: int, thang: int) -> "Ky":
        return cls(nam, thang, thang)

    @classmethod
    def quy(cls, nam: int, quy: int) -> "Ky":
        return cls(nam, 3 * quy - 2, 3 * quy)

    @classmethod
    def ca_nam(cls, nam: int) -> "Ky":
        return cls(nam, 1, 12)

    @property
    def cac_thang(self) -> range:
        return range(self.thang_dau, self.thang_cuoi + 1)

    def chua(self, khac: "Ky") -> bool:
        return self.nam == khac.nam and self.thang_dau <= khac.thang_dau and khac.thang_cuoi <= self.thang_cuoi

    def giao(self, khac: "Ky") -> bool:
        return self.nam == khac.nam and self.thang_dau <= khac.thang_cuoi and khac.thang_dau <= self.thang_cuoi

    @property
    def ten(self) -> str:
        d, c = self.thang_dau, self.thang_cuoi
        if d == c:
            return f"Tháng {d}/{self.nam}"
        if d == 1 and c == 12:
            return f"Năm {self.nam}"
        if c - d == 2 and d % 3 == 1:
            return f"Quý {(d + 2) // 3}/{self.nam}"
        if d == 1:
            return f"{c} tháng đầu năm {self.nam}"
        return f"Tháng {d}-{c}/{self.nam}"

    @property
    def ma(self) -> str:
        """Mã ngắn dùng làm tên cột / tên file."""
        d, c = self.thang_dau, self.thang_cuoi
        if d == c:
            return f"{self.nam}-T{d:02d}"
        if d == 1 and c == 12:
            return f"{self.nam}"
        if c - d == 2 and d % 3 == 1:
            return f"{self.nam}-Q{(d + 2) // 3}"
        return f"{self.nam}-T{d:02d}-{c:02d}"


def chuan_hoa(text: str) -> str:
    """Chuẩn hóa Unicode (NFC) và gộp khoảng trắng."""
    return re.sub(r"\s+", " ", unicodedata.normalize("NFC", text)).strip()


def phan_tich_ky(text: str) -> Ky | None:
    """Nhận diện kỳ từ tiêu đề, ví dụ:

    'Tháng 5 năm 2025', 'tháng 5 và phương hướng công tác tháng 6 năm 2025',
    '6 tháng đầu năm ...  2025', 'Quý II năm 2025', '9 tháng năm 2025', 'năm 2025'.
    Kỳ tiếp theo trong phần "phương hướng" bị bỏ qua.
    """
    t = chuan_hoa(text).lower()
    nam_tim = re.findall(r"(?:19|20)\d{2}", t)
    if not nam_tim:
        return None
    nam = int(nam_tim[-1])
    # bỏ phần "phương hướng ..." để không nhầm với kỳ tới
    t_chinh = re.split(r"phương hướng|nhiệm vụ trọng tâm|kỳ tới", t)[0]

    m = re.search(r"(\d{1,2})\s*tháng\s*(?:đầu\s*)?năm", t_chinh)
    if m and not re.search(r"tháng\s*\d", t_chinh[: m.start()]):
        return Ky(nam, 1, int(m.group(1)))
    m = re.search(r"(\d{1,2})\s*tháng\s*cuối\s*năm", t_chinh)
    if m:
        return Ky(nam, 13 - int(m.group(1)), 12)
    m = re.search(r"quý\s*(iv|iii|ii|i|[1-4])\b", t_chinh)
    if m:
        q = m.group(1).upper()
        return Ky.quy(nam, LA_MA.get(q) or int(q))
    m = re.search(r"tháng\s*(\d{1,2})", t_chinh)
    if m and 1 <= int(m.group(1)) <= 12:
        return Ky.thang(nam, int(m.group(1)))
    if re.search(r"(?:cả\s*)?năm\s*(?:19|20)\d{2}", t_chinh) or "tổng kết" in t_chinh:
        return Ky.ca_nam(nam)
    return None


def phan_tich_tham_so_ky(chuoi: str) -> Ky:
    """Đọc tham số dòng lệnh: 2025-06 | 2025-T6 | 2025-Q2 | 2025-6T | 2025 ."""
    s = chuoi.strip().upper()
    if m := re.fullmatch(r"(\d{4})", s):
        return Ky.ca_nam(int(m.group(1)))
    if m := re.fullmatch(r"(\d{4})-?Q([1-4])", s):
        return Ky.quy(int(m.group(1)), int(m.group(2)))
    if m := re.fullmatch(r"(\d{4})-?(?:T)?(\d{1,2})", s):
        return Ky.thang(int(m.group(1)), int(m.group(2)))
    if m := re.fullmatch(r"(\d{4})-?(\d{1,2})T", s):  # 2025-6T = 6 tháng đầu năm
        return Ky(int(m.group(1)), 1, int(m.group(2)))
    raise ValueError(f"Kỳ không hợp lệ: {chuoi!r} (ví dụ: 2025-06, 2025-Q2, 2025-6T, 2025)")
