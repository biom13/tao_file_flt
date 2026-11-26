import os
import re
import struct
import sys
import unicodedata
from collections import OrderedDict
import openpyxl

IT_COL_NAMES = {
    "it code",
    "it. no",
    "it no",
    "it_no",
    "it_code",
    "itno",
    "itcode",
}

IT_FALLBACK_COL_NAMES = {
    "no.",
    "no",
}

GATE_COL_NAMES = {
    "cổng được phép quẹt vân tay",
    "máy được phép quẹt bằng vân tay",
    "máy được phép quẹt vân tay",
    "machine no.",
    "machine no",
    "machine_no",
}

INVALID_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def normalize_header(val):
    if val is None:
        return ""
    return re.sub(r"\s+", " ", str(val).strip()).casefold()


def normalize_gate_key(val):
    s = re.sub(r"\s+", " ", str(val).strip()).casefold()
    m = re.match(r"^(cổng|máy)\s+(.+)$", s)
    if m:
        prefix = m.group(1)
        rest = m.group(2)
    else:
        prefix = ""
        rest = s
    parts = []
    for p in rest.split(","):
        p = p.strip()
        if prefix:
            p = re.sub(r"^(cổng|máy)\s+", "", p)
        if p:
            parts.append(p)
    return (prefix, tuple(sorted(parts)))


def remove_vietnamese_accents(text):
    text = text.replace("đ", "d").replace("Đ", "D")
    nfd = unicodedata.normalize("NFD", text)
    return "".join(ch for ch in nfd if unicodedata.category(ch) != "Mn")


def sanitize_filename(name):
    cleaned = remove_vietnamese_accents(name.strip())
    cleaned = re.sub(r"\s*,\s*", "_", cleaned)
    cleaned = re.sub(r"[^A-Za-z0-9 _-]+", "_", cleaned)
    cleaned = re.sub(r"_+", "_", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned)
    cleaned = cleaned.strip(" _.")
    return cleaned or "unknown_gate"


def format_it_code(val):
    if val is None:
        return ""
    if isinstance(val, bool):
        return ""
    if isinstance(val, int):
        return str(val)
    if isinstance(val, float):
        if val.is_integer():
            return str(int(val))
        return str(val).strip()
    return str(val).strip()


def is_valid_it_code(code):
    if not code:
        return False
    if re.fullmatch(r"column\d+", code, flags=re.IGNORECASE):
        return False
    return True


def build_flt_bytes(it_codes):
    """Build DevExpress TcxFilterCriteria binary .flt content."""
    out = bytearray(b"\xff\xff\xff\xff\x01")
    out += struct.pack("<I", len(it_codes))
    for code in it_codes:
        raw = code.encode("utf-8")
        n = len(raw)
        if n > 255:
            raise ValueError(f"IT Code too long (>255 bytes): {code!r}")
        out += b"\x00\x06"
        out += struct.pack("<I", n)
        out += raw
        out += b"\x01\x00\x00\x00\x06"
        out += struct.pack("B", n)
        out += raw
    return bytes(out)


def find_columns(headers):
    it_idx = None
    it_fallback_idx = None
    gate_idx = None
    for idx, h in enumerate(headers):
        nh = normalize_header(h)
        if not nh:
            continue
        if it_idx is None and nh in IT_COL_NAMES:
            it_idx = idx
        elif it_fallback_idx is None and nh in IT_FALLBACK_COL_NAMES:
            it_fallback_idx = idx
        if gate_idx is None and nh in GATE_COL_NAMES:
            gate_idx = idx
    if it_idx is None:
        it_idx = it_fallback_idx
    return it_idx, gate_idx


def collect_from_folder(folder):
    groups = OrderedDict()
    excel_files = [
        fn
        for fn in sorted(os.listdir(folder))
        if fn.lower().endswith((".xlsx", ".xlsm")) and not fn.startswith("~$")
    ]

    if not excel_files:
        print("Khong tim thay file Excel (.xlsx/.xlsm) nao trong thu muc.")
        return groups

    for fn in excel_files:
        path = os.path.join(folder, fn)
        print(f"Dang doc: {fn}")
        wb = openpyxl.load_workbook(path, data_only=True)
        try:
            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                if ws.sheet_state != "visible":
                    continue
                if ws.max_row < 2:
                    continue
                headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
                it_idx, gate_idx = find_columns(headers)
                if it_idx is None or gate_idx is None:
                    continue

                it_col = it_idx + 1
                gate_col = gate_idx + 1
                count_sheet = 0

                for r in range(2, ws.max_row + 1):
                    it_raw = ws.cell(r, it_col).value
                    gate_raw = ws.cell(r, gate_col).value
                    it_code = format_it_code(it_raw)
                    if not is_valid_it_code(it_code):
                        continue
                    if gate_raw is None:
                        continue
                    gate_str = re.sub(r"\s+", " ", str(gate_raw).strip())
                    if not gate_str or re.fullmatch(
                        r"column\d+", gate_str, flags=re.IGNORECASE
                    ):
                        continue

                    key = normalize_gate_key(gate_str)
                    if key not in groups:
                        groups[key] = {"display": gate_str, "codes": OrderedDict()}
                    groups[key]["codes"][it_code] = None
                    count_sheet += 1

                if count_sheet > 0:
                    print(f"  - Sheet '{sheet_name}': {count_sheet} dong hop le")
        finally:
            wb.close()

    return groups


def main():
    folder = os.path.dirname(os.path.abspath(__file__))
    print(f"Thu muc lam viec: {folder}")
    print("-" * 50)

    groups = collect_from_folder(folder)
    if not groups:
        print("Khong co du lieu hop le de tao file .flt.")
        return 1

    out_dir = os.path.join(folder, "result")
    os.makedirs(out_dir, exist_ok=True)

    print("-" * 50)
    generated = 0
    for group in groups.values():
        codes = list(group["codes"].keys())
        if not codes:
            continue
        filename = sanitize_filename(group["display"]) + ".flt"
        out_path = os.path.join(out_dir, filename)
        data = build_flt_bytes(codes)
        with open(out_path, "wb") as f:
            f.write(data)
        print(f"Da tao: result\\{filename} ({len(codes)} IT Code)")
        generated += 1

    print("-" * 50)
    print(f"Hoan tat! Da tao {generated} file .flt.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
