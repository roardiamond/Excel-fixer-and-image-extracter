#!/usr/bin/env python3
"""
extract_excel_images.py

1) Extracts ALL embedded images from an .xlsx file, named by the row of the cell
   they are anchored to (an image at H23 -> 23.png).

2) Creates an EMAIL-SAFE COPY of the workbook. Excel 365 "Place in Cell" pictures
   show as #VALUE! in older Excel, Outlook/Gmail previews, phone apps, etc.
   The copy turns each of them into a normal picture that sits exactly over the
   same cell, so the images display everywhere.

Standard library only (no pip install needed). The original file is opened
read-only and is never modified.
"""

import posixpath
import re
import struct
import sys
import xml.etree.ElementTree as ET
import zipfile
from collections import Counter
from pathlib import Path

# ----------------------------- SETTINGS -----------------------------
EXCEL_FILE = "input.xlsx"                 # place this file beside the script
OUTPUT_FOLDER = "extracted_images"        # extracted images go here
EXTRACT_IMAGES = True                     # save every image as 23.png, 24.jpg ...
MAKE_EMAIL_SAFE_COPY = True               # create the copy that works when emailed
EMAIL_SAFE_FILE = "input_email_safe.xlsx" # name of that copy (send THIS one)
ONLY_COLUMN = None                        # e.g. "H" to extract only column H; None = all
# --------------------------------------------------------------------

REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
XDR_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_ID = "{%s}id" % REL_NS
R_EMBED = "{%s}embed" % REL_NS
CELL_RE = re.compile(r"^([A-Z]+)(\d+)$")
ANCHORS = {"oneCellAnchor", "twoCellAnchor", "absoluteAnchor"}
PX = 9525  # EMUs per pixel


# ============================ generic helpers ============================
def local(tag):
    return tag.rsplit("}", 1)[-1]


def col_to_letters(n):
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def letters_to_col(s):
    n = 0
    for ch in s:
        n = n * 26 + (ord(ch) - 64)
    return n


def rels_path(part):
    d, b = posixpath.split(part)
    return posixpath.join(d, "_rels", b + ".rels")


def resolve(base_part, target):
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(base_part), target))


def read_xml(zf, part):
    try:
        return ET.fromstring(zf.read(part))
    except KeyError:
        return None


def read_rels(zf, part):
    """{rId: (type, resolved_target_part)} for a package part."""
    root = read_xml(zf, rels_path(part))
    out = {}
    if root is None:
        return out
    for r in root:
        if local(r.tag) != "Relationship" or r.get("TargetMode") == "External":
            continue
        out[r.get("Id")] = (r.get("Type", ""), resolve(part, r.get("Target", "")))
    return out


def get_sheets(zf):
    wb = read_xml(zf, "xl/workbook.xml")
    if wb is None:
        raise ValueError("xl/workbook.xml not found - is this a valid .xlsx file?")
    rels = read_rels(zf, "xl/workbook.xml")
    sheets = []
    for el in wb.iter():
        if local(el.tag) == "sheet":
            rid = el.get(R_ID)
            if rid in rels:
                sheets.append((el.get("name"), rels[rid][1]))
    return sheets


def safe(name):
    return re.sub(r'[\\/:*?"<>|]', "_", name)


def unique_path(folder, stem, ext):
    p = folder / (stem + ext)
    n = 1
    while p.exists():
        p = folder / ("%s_dup%d%s" % (stem, n, ext))
        n += 1
    return p


# ================= find images: floating pictures (drawings) =================
def drawing_images(zf, sheet_part, names):
    found = []
    for rtype, dpart in read_rels(zf, sheet_part).values():
        if not rtype.endswith("/drawing"):
            continue
        root = read_xml(zf, dpart)
        if root is None:
            continue
        drels = read_rels(zf, dpart)
        for anchor in root.iter():
            if local(anchor.tag) not in ANCHORS:
                continue
            col = row = None
            for ch in anchor:
                if local(ch.tag) == "from":
                    for g in ch:
                        if local(g.tag) == "col":
                            col = int(g.text) + 1
                        elif local(g.tag) == "row":
                            row = int(g.text) + 1
                    break
            for blip in anchor.iter():
                if local(blip.tag) != "blip":
                    continue
                target = drels.get(blip.get(R_EMBED))
                if target and target[1] in names:
                    found.append((col_to_letters(col) if col else None, row, target[1]))
    return found


# ============ find images: "Place in Cell" pictures (rich data) ============
def load_richdata(zf, names):
    """{vm_index: media_part_path} for in-cell pictures."""
    meta = read_xml(zf, "xl/metadata.xml")
    rv_root = read_xml(zf, "xl/richData/rdrichvalue.xml")
    st_root = read_xml(zf, "xl/richData/rdrichvaluestructure.xml")
    rel_root = read_xml(zf, "xl/richData/richValueRel.xml")
    if any(x is None for x in (meta, rv_root, st_root, rel_root)):
        return {}

    rels = read_rels(zf, "xl/richData/richValueRel.xml")
    rel_parts = []
    for el in rel_root:
        if local(el.tag) == "rel":
            t = rels.get(el.get(R_ID))
            rel_parts.append(t[1] if t else None)

    structs = [[k.get("n") for k in s if local(k.tag) == "k"]
               for s in st_root if local(s.tag) == "s"]

    rich_to_part = []
    for rv in rv_root:
        if local(rv.tag) != "rv":
            continue
        vals = [v.text for v in rv if local(v.tag) == "v"]
        part = None
        try:
            keys = structs[int(rv.get("s", "0"))]
            i = keys.index("_rvRel:LocalImageIdentifier")
            part = rel_parts[int(vals[i])]
        except (ValueError, IndexError, TypeError):
            pass
        rich_to_part.append(part)

    types, futures, vbks = [], {}, []
    for ch in meta:
        n = local(ch.tag)
        if n == "metadataTypes":
            types = [t.get("name") for t in ch if local(t.tag) == "metadataType"]
        elif n == "futureMetadata":
            lst = []
            for bk in ch:
                if local(bk.tag) != "bk":
                    continue
                idx = None
                for e in bk.iter():
                    if local(e.tag) == "rvb":
                        idx = int(e.get("i"))
                        break
                lst.append(idx)
            futures[ch.get("name")] = lst
        elif n == "valueMetadata":
            for bk in ch:
                if local(bk.tag) == "bk":
                    vbks.append(next((e for e in bk if local(e.tag) == "rc"), None))

    vm_map = {}
    for vm, rc in enumerate(vbks, 1):
        if rc is None:
            continue
        try:
            t, v = int(rc.get("t")), int(rc.get("v"))
            name = types[t - 1]
            rich_idx = futures[name][v]
            part = rich_to_part[rich_idx]
        except (ValueError, IndexError, KeyError, TypeError):
            continue
        if part and part in names:
            vm_map[vm] = part
    return vm_map


def richdata_images(zf, sheet_part, vm_map):
    found = []
    if not vm_map:
        return found
    with zf.open(sheet_part) as f:
        for _, el in ET.iterparse(f, events=("end",)):
            tag = local(el.tag)
            if tag == "c":
                vm = el.get("vm")
                if vm and vm.isdigit() and int(vm) in vm_map:
                    m = CELL_RE.match(el.get("r", ""))
                    if m:
                        found.append((m.group(1), int(m.group(2)), vm_map[int(vm)]))
            elif tag == "row":
                el.clear()
    return found


# ============================ PART 1: extraction ============================
def extract_all(zf, names, sheets, vm_map, base):
    out_dir = base / OUTPUT_FOLDER
    out_dir.mkdir(parents=True, exist_ok=True)

    records = []  # (sheet_index, sheet_name, col, row, media_part)
    for si, (sname, spart) in enumerate(sheets):
        if spart not in names:
            continue
        for col, row, part in drawing_images(zf, spart, names):
            records.append((si, sname, col, row, part))
        for col, row, part in richdata_images(zf, spart, vm_map):
            records.append((si, sname, col, row, part))

    skipped = 0
    if ONLY_COLUMN:
        want = ONLY_COLUMN.upper()
        kept = [r for r in records if r[2] == want]
        skipped = len(records) - len(kept)
        records = kept

    records.sort(key=lambda r: (r[0], r[3] if r[3] else 10 ** 9, len(r[2] or ""), r[2] or ""))

    # Rows are numbered across the whole workbook, so if two sheets both have
    # an image on row 23 they become 23_1 and 23_2 (nothing is overwritten).
    row_counts = Counter(r[3] for r in records if r[3])
    row_seen = Counter()
    unplaced = Counter()
    total = 0

    for si, sname, col, row, part in records:
        ext = posixpath.splitext(part)[1].lower() or ".png"
        if row:
            row_seen[row] += 1
            stem = ("%d_%d" % (row, row_seen[row])) if row_counts[row] > 1 else str(row)
            label = "%s%d" % (col, row)
        else:
            unplaced[sname] += 1
            stem = "%s_unplaced_%d" % (safe(sname), unplaced[sname])
            label = "(no cell anchor)"
        target = unique_path(out_dir, stem, ext)
        target.write_bytes(zf.read(part))
        total += 1
        print("Extracted: %s %s -> %s/%s" % (sname, label, OUTPUT_FOLDER, target.name))

    if skipped:
        print("Skipped %d image(s) outside column %s." % (skipped, ONLY_COLUMN.upper()))
    if total == 0:
        print("No embedded images were found in this workbook.")
    print("\nTotal images extracted: %d" % total)
    print("Output folder: %s" % out_dir)
    return total


# ====================== PART 2: email-safe workbook copy ======================
CELL_TAG = re.compile(r"<c\b([^>]*?)(?<!/)>(.*?)</c>", re.S)


def image_size(data):
    """(width, height) in pixels from PNG/JPEG/GIF/BMP headers, else None."""
    try:
        if data[:8] == b"\x89PNG\r\n\x1a\n":
            return struct.unpack(">II", data[16:24])
        if data[:6] in (b"GIF87a", b"GIF89a"):
            return struct.unpack("<HH", data[6:10])
        if data[:2] == b"BM":
            w, h = struct.unpack("<ii", data[18:26])
            return w, abs(h)
        if data[:2] == b"\xff\xd8":
            i = 2
            while i < len(data) - 9:
                if data[i] != 0xFF:
                    i += 1
                    continue
                marker = data[i + 1]
                if marker == 0xFF:
                    i += 1
                    continue
                if marker == 0x01 or 0xD0 <= marker <= 0xD8:
                    i += 2
                    continue
                seg = struct.unpack(">H", data[i + 2:i + 4])[0]
                if 0xC0 <= marker <= 0xCF and marker not in (0xC4, 0xC8, 0xCC):
                    h, w = struct.unpack(">HH", data[i + 5:i + 9])
                    return w, h
                i += 2 + seg
    except Exception:
        pass
    return None


def sheet_geometry(text):
    """Column widths / row heights in pixels, read from the sheet XML."""
    def_col, def_row = 64.0, 15.0 * 96 / 72
    m = re.search(r"<sheetFormatPr\b([^>]*)>", text)
    if m:
        a = m.group(1)
        w = re.search(r'\bdefaultColWidth="([\d.]+)"', a)
        h = re.search(r'\bdefaultRowHeight="([\d.]+)"', a)
        if w:
            def_col = float(w.group(1)) * 7
        if h:
            def_row = float(h.group(1)) * 96 / 72
    cols = {}
    for m in re.finditer(r"<col\b([^>]*?)/?>", text):
        a = m.group(1)
        mn = re.search(r'\bmin="(\d+)"', a)
        mx = re.search(r'\bmax="(\d+)"', a)
        wd = re.search(r'\bwidth="([\d.]+)"', a)
        if mn and mx and wd:
            for c in range(int(mn.group(1)), min(int(mx.group(1)), 16384) + 1):
                cols[c] = float(wd.group(1)) * 7
    rows = {}
    for m in re.finditer(r"<row\b([^>]*)>", text):
        a = m.group(1)
        r = re.search(r'\br="(\d+)"', a)
        ht = re.search(r'\bht="([\d.]+)"', a)
        if r and ht:
            rows[int(r.group(1))] = float(ht.group(1)) * 96 / 72
    return def_col, def_row, cols, rows


def build_anchor(col_letters, row, rid, pic_id, cw, rh, size):
    """A picture centred in the cell, scaled to fit it (aspect ratio kept)."""
    iw, ih = size if size and size[0] and size[1] else (cw, rh)
    scale = min(max(cw - 4, 2) / iw, max(rh - 4, 2) / ih)
    w = max(int(iw * scale), 1)
    h = max(int(ih * scale), 1)
    xoff = max(int((cw - w) / 2 * PX), 0)
    yoff = max(int((rh - h) / 2 * PX), 0)
    return (
        '<xdr:oneCellAnchor xmlns:xdr="%s" xmlns:a="%s" xmlns:r="%s">'
        '<xdr:from><xdr:col>%d</xdr:col><xdr:colOff>%d</xdr:colOff>'
        '<xdr:row>%d</xdr:row><xdr:rowOff>%d</xdr:rowOff></xdr:from>'
        '<xdr:ext cx="%d" cy="%d"/>'
        '<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="%d" name="Picture %d"/>'
        '<xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>'
        '<xdr:blipFill><a:blip r:embed="%s"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>'
        '<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>'
        '<xdr:clientData/></xdr:oneCellAnchor>'
    ) % (XDR_NS, A_NS, REL_NS, letters_to_col(col_letters) - 1, xoff, row - 1, yoff,
         w * PX, h * PX, pic_id, pic_id, rid, w * PX, h * PX)


def add_relationships(rels_text, new_rels):
    if not rels_text:
        rels_text = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                     '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                     '</Relationships>')
    return rels_text.replace("</Relationships>", "".join(new_rels) + "</Relationships>", 1)


def next_rid(rels_text):
    used = [int(x) for x in re.findall(r'Id="rId(\d+)"', rels_text or "")]
    return max(used or [0]) + 1


def make_email_safe(zf, names, sheets, vm_map, out_path):
    if not vm_map:
        print("\nNo in-cell pictures found - nothing to convert.")
        print("(Normal floating pictures already display correctly when emailed.)")
        return 0

    edits, new_parts = {}, {}
    ct = zf.read("[Content_Types].xml").decode("utf-8")
    ct_add = []
    existing_nums = [int(m.group(1)) for n in names
                     for m in [re.match(r"xl/drawings/drawing(\d+)\.xml$", n)] if m]
    next_drawing = max(existing_nums or [0]) + 1
    converted = 0

    for sname, spart in sheets:
        if spart not in names:
            continue
        text = zf.read(spart).decode("utf-8")
        found = []

        def strip_cell(m):
            attrs = m.group(1)
            vm = re.search(r'\bvm="(\d+)"', attrs)
            ref = re.search(r'\br="([A-Z]+)(\d+)"', attrs)
            if not vm or not ref or int(vm.group(1)) not in vm_map:
                return m.group(0)
            found.append((ref.group(1), int(ref.group(2)), vm_map[int(vm.group(1))]))
            attrs = re.sub(r'\s(?:vm|cm|t)="[^"]*"', "", attrs)
            return "<c%s/>" % attrs  # empty cell, keeps its formatting

        new_text = CELL_TAG.sub(strip_cell, text)
        if not found:
            continue

        def_col, def_row, col_px, row_px = sheet_geometry(text)

        # Where do the pictures go: the sheet's existing drawing, or a new one.
        existing = [dp for t, dp in read_rels(zf, spart).values() if t.endswith("/drawing")]
        if existing:
            dpart = existing[0]
            dtext = zf.read(dpart).decode("utf-8")
            new_drawing = False
        else:
            dpart = "xl/drawings/drawing%d.xml" % next_drawing
            next_drawing += 1
            dtext = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                     '<xdr:wsDr xmlns:xdr="%s" xmlns:a="%s"></xdr:wsDr>' % (XDR_NS, A_NS))
            new_drawing = True
            ct_add.append('<Override PartName="/%s" ContentType="application/'
                          'vnd.openxmlformats-officedocument.drawing+xml"/>' % dpart)
        drels_part = rels_path(dpart)
        drels_text = zf.read(drels_part).decode("utf-8") if drels_part in names else ""

        ids = [int(x) for x in re.findall(r'<(?:\w+:)?cNvPr\b[^>]*?\bid="(\d+)"', dtext)]
        pic_id = max(ids or [1])
        rid_n = next_rid(drels_text) - 1
        anchors, rel_items = [], []

        for col, row, media in sorted(found, key=lambda x: (x[1], letters_to_col(x[0]))):
            pic_id += 1
            rid_n += 1
            rid = "rId%d" % rid_n
            rel_items.append('<Relationship Id="%s" Type="%s/image" Target="%s"/>' % (
                rid, REL_NS, posixpath.relpath(media, posixpath.dirname(dpart))))
            cw = col_px.get(letters_to_col(col), def_col)
            rh = row_px.get(row, def_row)
            anchors.append(build_anchor(col, row, rid, pic_id, cw, rh, image_size(zf.read(media))))
            converted += 1
            print("Converted: %s %s%d (in-cell picture -> normal picture)" % (sname, col, row))

        # add the anchors to the drawing xml
        last = None
        for last in re.finditer(r"</(?:\w+:)?wsDr\s*>", dtext):
            pass
        if last:
            dtext = dtext[:last.start()] + "".join(anchors) + dtext[last.start():]
        else:
            dtext = re.sub(r"<((?:\w+:)?wsDr)\b([^>]*?)/>",
                           lambda m: "<%s%s>%s</%s>" % (m.group(1), m.group(2), "".join(anchors), m.group(1)),
                           dtext, count=1)
        edits[dpart] = dtext.encode("utf-8")
        edits_or_new = new_parts if drels_part not in names else edits
        edits_or_new[drels_part] = add_relationships(drels_text, rel_items).encode("utf-8")
        if new_drawing:
            new_parts[dpart] = edits.pop(dpart)

            # link the new drawing from the sheet
            srels_part = rels_path(spart)
            srels_text = zf.read(srels_part).decode("utf-8") if srels_part in names else ""
            srid = "rId%d" % next_rid(srels_text)
            srel = '<Relationship Id="%s" Type="%s/drawing" Target="%s"/>' % (
                srid, REL_NS, posixpath.relpath(dpart, posixpath.dirname(spart)))
            (edits if srels_part in names else new_parts)[srels_part] = \
                add_relationships(srels_text, [srel]).encode("utf-8")

            tag = '<drawing xmlns:r="%s" r:id="%s"/>' % (REL_NS, srid)
            pm = re.search(r"<pageMargins\b", new_text)
            start = pm.start() if pm else max(new_text.find("</sheetData>"), 0)
            m = re.compile(r"<(?:mc:AlternateContent|legacyDrawing|legacyDrawingHF|drawingHF|picture|"
                           r"oleObjects|controls|webPublishItems|tableParts|extLst)\b").search(new_text, start)
            pos = m.start() if m else new_text.rindex("</worksheet>")
            new_text = new_text[:pos] + tag + new_text[pos:]

        edits[spart] = new_text.encode("utf-8")

    # make sure every image extension used has a content type
    ctypes = {"png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg", "gif": "image/gif",
              "bmp": "image/bmp", "tif": "image/tiff", "tiff": "image/tiff",
              "emf": "image/x-emf", "wmf": "image/x-wmf", "svg": "image/svg+xml"}
    for e, t in ctypes.items():
        if any(n.lower().startswith("xl/media/") and n.lower().endswith("." + e) for n in names) \
                and not re.search(r'Extension="%s"' % e, ct, re.I):
            ct_add.append('<Default Extension="%s" ContentType="%s"/>' % (e, t))
    if ct_add:
        edits["[Content_Types].xml"] = ct.replace("</Types>", "".join(ct_add) + "</Types>", 1).encode("utf-8")

    # write the new package (original file untouched)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zf.infolist():
            data = edits.get(info.filename)
            zout.writestr(info.filename, data if data is not None else zf.read(info.filename))
        for name, data in new_parts.items():
            zout.writestr(name, data)

    print("\nEmail-safe copy created (%d picture(s) converted):" % converted)
    print(out_path)
    print("Send THIS file by email; keep the original for yourself.")
    return converted


# ================================= main =================================
def main():
    base = Path(__file__).resolve().parent
    excel_path = Path(EXCEL_FILE)
    if not excel_path.is_absolute():
        excel_path = base / excel_path
        if not excel_path.is_file() and (Path.cwd() / EXCEL_FILE).is_file():
            excel_path = Path.cwd() / EXCEL_FILE
    if not excel_path.is_file():
        print("ERROR: Excel file not found: %s" % excel_path)
        print('Place your file beside this script and name it "%s".' % EXCEL_FILE)
        return 1

    out_path = base / EMAIL_SAFE_FILE
    if out_path.resolve() == excel_path.resolve():
        print("ERROR: EMAIL_SAFE_FILE must be different from EXCEL_FILE.")
        return 1

    try:
        with zipfile.ZipFile(excel_path, "r") as zf:  # read-only
            names = set(zf.namelist())
            sheets = get_sheets(zf)
            vm_map = load_richdata(zf, names)
            if EXTRACT_IMAGES:
                extract_all(zf, names, sheets, vm_map, base)
            if MAKE_EMAIL_SAFE_COPY:
                make_email_safe(zf, names, sheets, vm_map, out_path)
    except zipfile.BadZipFile:
        print("ERROR: %s is not a valid .xlsx file (old .xls files are not supported)." % excel_path.name)
        return 1
    except ValueError as e:
        print("ERROR: %s" % e)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())