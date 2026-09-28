"""Đọc nội dung file Word (.docx và .doc) thành đoạn văn và bảng.

- .docx: dùng python-docx, giữ nguyên bảng.
- .doc (Word 97-2003): tự giải mã bằng Python thuần (olefile), không cần
  cài Word hay LibreOffice. Chỉ lấy được văn bản (đủ cho file báo cáo lời văn).

Định dạng được nhận theo nội dung file, không theo đuôi file.
"""

from __future__ import annotations

import struct
from dataclasses import dataclass, field
from pathlib import Path

OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
ZIP_MAGIC = b"PK\x03\x04"


@dataclass
class TaiLieu:
    """Nội dung thô của một file Word."""

    duong_dan: Path
    doan_van: list[str] = field(default_factory=list)
    # Mỗi bảng là danh sách hàng, mỗi hàng là danh sách ô (đã gộp ô merge).
    bang: list[list[list[str]]] = field(default_factory=list)


def doc_file(duong_dan: str | Path) -> TaiLieu:
    duong_dan = Path(duong_dan)
    with open(duong_dan, "rb") as f:
        dau = f.read(8)
    if dau.startswith(ZIP_MAGIC):
        return _doc_docx(duong_dan)
    if dau == OLE_MAGIC:
        return _doc_doc(duong_dan)
    raise ValueError(f"Không nhận dạng được định dạng Word: {duong_dan}")


def _doc_docx(duong_dan: Path) -> TaiLieu:
    import docx
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    d = docx.Document(str(duong_dan))
    tl = TaiLieu(duong_dan)
    for phan_tu in d.element.body.iterchildren():
        if phan_tu.tag.endswith("}p"):
            text = Paragraph(phan_tu, d).text.strip()
            if text:
                tl.doan_van.append(text)
        elif phan_tu.tag.endswith("}tbl"):
            bang = []
            for hang in Table(phan_tu, d).rows:
                o_list, truoc = [], None
                for o in hang.cells:
                    if o._tc is truoc:  # ô gộp ngang lặp lại
                        continue
                    truoc = o._tc
                    o_list.append(o.text.strip())
                bang.append(o_list)
            tl.bang.append(bang)
    return tl


def _doc_doc(duong_dan: Path) -> TaiLieu:
    text = doc_text_word97(duong_dan)
    tl = TaiLieu(duong_dan)
    for dong in text.split("\r"):
        # \x07: hết ô bảng; \x0b: xuống dòng mềm; \x0c: ngắt trang
        for phan in dong.replace("\x0b", "\n").replace("\x0c", "\n").split("\x07"):
            phan = phan.strip()
            if phan:
                tl.doan_van.append(phan)
    return tl


def doc_text_word97(duong_dan: str | Path) -> str:
    """Trích văn bản thân tài liệu từ file .doc theo đặc tả MS-DOC (piece table)."""
    import olefile

    with olefile.OleFileIO(str(duong_dan)) as ole:
        wd = ole.openstream("WordDocument").read()
        co = struct.unpack_from("<H", wd, 0x0A)[0]
        if co & 0x0100:
            raise ValueError(f"File .doc bị đặt mật khẩu: {duong_dan}")
        bang_stream = "1Table" if co & 0x0200 else "0Table"
        tbl = ole.openstream(bang_stream).read()

    ccp_text = struct.unpack_from("<i", wd, 0x4C)[0]
    fc_clx, lcb_clx = struct.unpack_from("<II", wd, 0x1A2)
    clx = tbl[fc_clx : fc_clx + lcb_clx]

    i = 0
    while clx[i] == 0x01:  # bỏ qua các Prc
        i += 3 + struct.unpack_from("<H", clx, i + 1)[0]
    if clx[i] != 0x02:
        raise ValueError(f"Cấu trúc .doc không hợp lệ: {duong_dan}")
    lcb = struct.unpack_from("<I", clx, i + 1)[0]
    plc = clx[i + 5 : i + 5 + lcb]
    n = (lcb - 4) // 12
    cps = struct.unpack_from(f"<{n + 1}i", plc, 0)

    phan = []
    for k in range(n):
        fc = struct.unpack_from("<I", plc, 4 * (n + 1) + 8 * k + 2)[0]
        so_ky_tu = cps[k + 1] - cps[k]
        if fc & 0x40000000:
            vt = (fc & ~0x40000000) // 2
            phan.append(wd[vt : vt + so_ky_tu].decode("cp1252", errors="replace"))
        else:
            phan.append(wd[fc : fc + 2 * so_ky_tu].decode("utf-16le", errors="replace"))
    text = "".join(phan)[:ccp_text]
    return _bo_ma_truong(text)


def _bo_ma_truong(text: str) -> str:
    """Bỏ mã trường (field code) giữa \\x13..\\x14, giữ kết quả hiển thị."""
    ra: list[str] = []
    ngan_xep: list[bool] = []  # True: đang ở phần mã của trường
    for ch in text:
        if ch == "\x13":
            ngan_xep.append(True)
            continue
        if ch == "\x14":
            if ngan_xep:
                ngan_xep[-1] = False
            continue
        if ch == "\x15":
            if ngan_xep:
                ngan_xep.pop()
            continue
        if any(ngan_xep):
            continue
        if ch in "\x01\x08":  # ký tự đối tượng/hình vẽ
            continue
        ra.append(ch)
    return "".join(ra)
