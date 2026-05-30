
from __future__ import print_function

import os
import posixpath
import re
import zipfile
import xml.etree.ElementTree as ET

try:
    from io import BytesIO
except ImportError:  # pragma: no cover - Python 2 fallback
    from StringIO import StringIO as BytesIO

try:
    from html import unescape as html_unescape
except ImportError:  # pragma: no cover - Python 2 fallback
    try:
        from HTMLParser import HTMLParser
        html_unescape = HTMLParser().unescape
    except Exception:  # pragma: no cover
        def html_unescape(value):
            return value

try:
    basestring
except NameError:  # pragma: no cover - Python 3
    basestring = str

from logger import log

REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
OD_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _to_text(data):
    """Return *data* decoded as text without throwing on malformed XML."""
    if data is None:
        return ""
    if isinstance(data, bytes):
        return data.decode("utf-8", "ignore")
    return data


def _norm_pkg_path(base_dir, target):
    """Normalize an OOXML relationship target into a package path."""
    if target is None:
        return None
    target = target.replace("\\", "/")
    if target.startswith("/"):
        return target[1:]
    return posixpath.normpath(posixpath.join(base_dir, target))


def _open_zip(data=None, filename=None):
    """Open an OOXML zip package from bytes or a filename."""
    if data is not None and not isinstance(data, Exception):
        if isinstance(data, bytes):
            return zipfile.ZipFile(BytesIO(data))
        if isinstance(data, basestring):
            try:
                raw = data.encode("latin1")
            except Exception:
                raw = data
            return zipfile.ZipFile(BytesIO(raw))
    if filename and os.path.exists(filename):
        return zipfile.ZipFile(filename)
    return None


def _safe_read(zf, name):
    try:
        return _to_text(zf.read(name))
    except Exception:
        return ""


def _parse_xml(data):
    try:
        return ET.fromstring(data)
    except Exception:
        return None


def _local_name(tag):
    if "}" in tag:
        return tag.rsplit("}", 1)[1]
    return tag


def _col_to_index(col):
    """Convert Excel column letters to 1-based index."""
    result = 0
    for ch in col.upper():
        if not ("A" <= ch <= "Z"):
            return None
        result = result * 26 + (ord(ch) - ord("A") + 1)
    return result


def _split_cell_ref(ref):
    m = re.match(r"^\$?([A-Za-z]+)\$?(\d+)$", str(ref).strip())
    if not m:
        return None
    return m.group(1).upper(), int(m.group(2))


def _cell_to_rc(ref):
    parts = _split_cell_ref(ref)
    if parts is None:
        return None
    col, row = parts
    col_index = _col_to_index(col)
    if col_index is None:
        return None
    return row - 1, col_index - 1


def _cell_key_variants(sheet_name, cell_ref):
    cell_ref = str(cell_ref).replace("$", "").upper()
    sheet_name = str(sheet_name).strip("'")
    return [
        "__excel_cell." + sheet_name.lower() + "!" + cell_ref.lower(),
        "__excel_cell." + cell_ref.lower(),
        sheet_name.lower() + "!" + cell_ref.lower(),
        cell_ref.lower(),
    ]


def _add_var(mapping, name, value):
    if name is None or value is None:
        return
    name = str(name)
    mapping[name] = value
    mapping[name.lower()] = value


def _add_doc_var(vm, name, value):
    _add_var(vm.doc_vars, name, value)


def _add_global(vm, name, value):
    _add_var(vm.globals, name, value)


class OOXMLWorkbookContext(object):
    """Read-only resolver for OOXML Excel workbook metadata."""

    def __init__(self, zf):
        self.zf = zf
        self.names = set(zf.namelist())
        self.sheet_paths = []
        self.sheet_names = []
        self.active_sheet_name = None
        self.defined_names = {}
        self.shared_strings = []
        self.cell_values = {}
        self.shape_values = []

    def read(self):
        self._read_shared_strings()
        self._read_workbook()
        self._read_cells()
        self._read_shapes()
        return self

    def _read_shared_strings(self):
        if "xl/sharedStrings.xml" not in self.names:
            return
        root = _parse_xml(_safe_read(self.zf, "xl/sharedStrings.xml"))
        if root is None:
            return
        for si in root.iter():
            if _local_name(si.tag) != "si":
                continue
            self.shared_strings.append("".join(si.itertext()))

    def _read_workbook_relationships(self):
        rels = {}
        root = _parse_xml(_safe_read(self.zf, "xl/_rels/workbook.xml.rels"))
        if root is None:
            return rels
        for rel in root.iter():
            if _local_name(rel.tag) != "Relationship":
                continue
            rel_id = rel.attrib.get("Id")
            target = rel.attrib.get("Target")
            if rel_id and target:
                rels[rel_id] = _norm_pkg_path("xl", target)
        return rels

    def _read_workbook(self):
        if "xl/workbook.xml" not in self.names:
            return
        rels = self._read_workbook_relationships()
        root = _parse_xml(_safe_read(self.zf, "xl/workbook.xml"))
        if root is None:
            return

        active_index = 0
        for elem in root.iter():
            if _local_name(elem.tag) == "workbookView":
                try:
                    active_index = int(elem.attrib.get("activeTab", "0"))
                except Exception:
                    active_index = 0
                break

        for sheet in root.iter():
            if _local_name(sheet.tag) != "sheet":
                continue
            name = sheet.attrib.get("name")
            rid = sheet.attrib.get("{%s}id" % OD_REL_NS)
            path = rels.get(rid)
            if name and path:
                self.sheet_names.append(name)
                self.sheet_paths.append(path)

        if len(self.sheet_names) > 0:
            if active_index < 0 or active_index >= len(self.sheet_names):
                active_index = 0
            self.active_sheet_name = self.sheet_names[active_index]

        for dn in root.iter():
            if _local_name(dn.tag) != "definedName":
                continue
            name = dn.attrib.get("name")
            target = "".join(dn.itertext()).strip()
            if name and target:
                self.defined_names[name.lower()] = target

    def _cell_value_from_elem(self, cell):
        v = None
        inline = None
        for child in cell.iter():
            lname = _local_name(child.tag)
            if lname == "v" and v is None:
                v = "".join(child.itertext())
            elif lname == "is":
                inline = "".join(child.itertext())
        typ = cell.attrib.get("t")
        if typ == "inlineStr" and inline is not None:
            return inline
        if v is None:
            return ""
        if typ == "s":
            try:
                idx = int(v)
                if idx >= 0 and idx < len(self.shared_strings):
                    return self.shared_strings[idx]
            except Exception:
                pass
        return v

    def _read_cells(self):
        for sheet_name, sheet_path in zip(self.sheet_names, self.sheet_paths):
            if sheet_path not in self.names:
                continue
            root = _parse_xml(_safe_read(self.zf, sheet_path))
            if root is None:
                continue
            for cell in root.iter():
                if _local_name(cell.tag) != "c":
                    continue
                ref = cell.attrib.get("r")
                if not ref:
                    continue
                value = self._cell_value_from_elem(cell)
                for key in _cell_key_variants(sheet_name, ref):
                    self.cell_values[key.lower()] = value

    def _resolve_defined_name_value(self, target):
        """Resolve a definedName target such as Sheet1!$K$2 to a cell value."""
        if target is None:
            return None
        target = str(target).strip()
        if "!" not in target:
            return None
        sheet_part, cell_part = target.rsplit("!", 1)
        sheet_part = sheet_part.strip().strip("'")
        if "]" in sheet_part:
            sheet_part = sheet_part[sheet_part.rindex("]") + 1:]
        cell_part = cell_part.strip().replace("$", "")
        if ":" in cell_part:
            cell_part = cell_part.split(":", 1)[0]
        key = "__excel_cell." + sheet_part.lower() + "!" + cell_part.lower()
        return self.cell_values.get(key)

    def _relationship_targets_for_sheet(self, sheet_path):
        """Return drawing package paths associated with a worksheet."""
        base = posixpath.dirname(sheet_path)
        rels_path = posixpath.join(base, "_rels", posixpath.basename(sheet_path) + ".rels")
        if rels_path not in self.names:
            return []
        root = _parse_xml(_safe_read(self.zf, rels_path))
        if root is None:
            return []
        targets = []
        for rel in root.iter():
            if _local_name(rel.tag) != "Relationship":
                continue
            typ = rel.attrib.get("Type", "")
            target = rel.attrib.get("Target")
            if target and typ.endswith("/drawing"):
                targets.append(_norm_pkg_path(base, target))
        return targets

    def _read_shape_info_from_drawing(self, drawing_path):
        text = _safe_read(self.zf, drawing_path)
        if len(text) == 0:
            return []
        root = _parse_xml(text)
        results = []
        if root is None:
            pat = r"<(?:[^:>]+:)?(?:cNvPr|docPr)\b([^>]*)>"
            chunks = re.findall(pat, text, re.I | re.S)
            if len(chunks) == 0:
                pat = r"<(?:[^:>]+:)?(?:cNvPr|docPr)\b([^>]*)/?>"
                chunks = re.findall(pat, text, re.I | re.S)
            for attrs in chunks:
                info = {}
                for key, val in re.findall(r"(\w+)=\"([^\"]*)\"", attrs):
                    info[key.lower()] = html_unescape(val)
                if info:
                    results.append(info)
            return results

        for elem in root.iter():
            lname = _local_name(elem.tag)
            if lname not in ("cNvPr", "docPr"):
                continue
            info = {}
            for key in ("id", "name", "descr", "title"):
                if key in elem.attrib:
                    info[key] = html_unescape(elem.attrib.get(key))
            if info:
                results.append(info)
        return results

    def _read_textbox_text_from_drawing(self, drawing_path):
        """Best effort extraction of text in DrawingML textboxes."""
        root = _parse_xml(_safe_read(self.zf, drawing_path))
        if root is None:
            return []
        texts = []
        curr = []
        for elem in root.iter():
            if _local_name(elem.tag) == "t":
                curr.append("".join(elem.itertext()))
        if curr:
            joined = "".join(curr)
            if joined:
                texts.append(joined)
        return texts

    def _read_shapes(self):
        seen_drawings = []
        for sheet_path in self.sheet_paths:
            for drawing_path in self._relationship_targets_for_sheet(sheet_path):
                if drawing_path not in seen_drawings:
                    seen_drawings.append(drawing_path)
        for name in sorted(self.names):
            if name.startswith("xl/drawings/drawing") and name.endswith(".xml"):
                if name not in seen_drawings:
                    seen_drawings.append(name)

        for drawing_path in seen_drawings:
            for info in self._read_shape_info_from_drawing(drawing_path):
                value = info.get("descr") or info.get("title") or info.get("name")
                if value is None:
                    continue
                self.shape_values.append({
                    "value": value,
                    "descr": info.get("descr"),
                    "title": info.get("title"),
                    "name": info.get("name"),
                    "id": info.get("id"),
                    "path": drawing_path,
                })
            for text_value in self._read_textbox_text_from_drawing(drawing_path):
                self.shape_values.append({
                    "value": text_value,
                    "text": text_value,
                    "path": drawing_path,
                })

    def populate(self, vm):
        """Populate a SimulationVBA instance with resolved workbook values."""
        for obj_name in ("Application", "Excel.Application", "ActiveWorkbook", "ThisWorkbook", "ActiveSheet"):
            _add_global(vm, obj_name, obj_name)

        if self.active_sheet_name:
            _add_global(vm, "ActiveSheet.Name", self.active_sheet_name)
            _add_global(vm, "Application.ActiveSheet.Name", self.active_sheet_name)
            _add_global(vm, "ActiveWorkbook.ActiveSheet.Name", self.active_sheet_name)

        for key, value in self.cell_values.items():
            _add_doc_var(vm, key, value)
            _add_global(vm, key, value)

        for name, target in self.defined_names.items():
            value = self._resolve_defined_name_value(target)
            if value is None:
                value = target
            _add_doc_var(vm, "__excel_defined_name." + name, value)
            _add_global(vm, "__excel_defined_name." + name, value)
            _add_global(vm, name, value)

        pos = 1
        for shape in self.shape_values:
            value = shape.get("value")
            if value is None:
                pos += 1
                continue
            accessors = [
                "AlternativeText", "AlternativeText$", "Title", "Name",
                "TextFrame.TextRange.Text", "TextFrame.ContainingRange",
            ]
            prefixes = [
                "Shapes('%d')" % pos,
                "Shapes(%d)" % pos,
                "ActiveSheet.Shapes('%d')" % pos,
                "ActiveSheet.Shapes(%d)" % pos,
                "Application.ActiveSheet.Shapes('%d')" % pos,
                "Application.ActiveSheet.Shapes(%d)" % pos,
            ]
            for sheet_name in self.sheet_names:
                prefixes.extend([
                    "Sheets('%s').Shapes('%d')" % (sheet_name, pos),
                    "Sheets(\"%s\").Shapes('%d')" % (sheet_name, pos),
                    "Worksheets('%s').Shapes('%d')" % (sheet_name, pos),
                    "Worksheets(\"%s\").Shapes('%d')" % (sheet_name, pos),
                ])
            for prefix in prefixes:
                _add_global(vm, prefix, prefix)
                for accessor in accessors:
                    _add_doc_var(vm, prefix + "." + accessor, value)
                    _add_global(vm, prefix + "." + accessor, value)
            pos += 1

        _best_payload = None
        _best_payload_score = -1
        _payload_prefixes = ("PGh0", "TVq", "UEsDB", "SUV", "SFlG", "UGs")
        for shape in self.shape_values:
            v = shape.get("value", "")
            if not v or len(v) < 16:
                continue
            score = 0
            for pfx in _payload_prefixes:
                if v.startswith(pfx):
                    score += 100
            if len(v) > score:
                score += len(v)
            if score > _best_payload_score:
                _best_payload_score = score
                _best_payload = v

        pos = 1
        for shape in self.shape_values:
            value = shape.get("value")
            if value is None:
                pos += 1
                continue
            if _best_payload is not None and len(value) < 16 and value != _best_payload:
                log.info(
                    "Shape position %d has suspiciously short alt-text (%d chars). "
                    "Falling back to best payload candidate (%d chars, starts with %r).",
                    pos, len(value), len(_best_payload), _best_payload[:20],
                )
                accessors = [
                    "AlternativeText", "AlternativeText$", "Title", "Name",
                    "TextFrame.TextRange.Text", "TextFrame.ContainingRange",
                ]
                prefixes = [
                    "Shapes('%d')" % pos,
                    "Shapes(%d)" % pos,
                    "ActiveSheet.Shapes('%d')" % pos,
                    "ActiveSheet.Shapes(%d)" % pos,
                    "Application.ActiveSheet.Shapes('%d')" % pos,
                    "Application.ActiveSheet.Shapes(%d)" % pos,
                ]
                for sheet_name in self.sheet_names:
                    prefixes.extend([
                        "Sheets('%s').Shapes('%d')" % (sheet_name, pos),
                        "Sheets(\"%s\").Shapes('%d')" % (sheet_name, pos),
                        "Worksheets('%s').Shapes('%d')" % (sheet_name, pos),
                        "Worksheets(\"%s\").Shapes('%d')" % (sheet_name, pos),
                    ])
                for prefix in prefixes:
                    for accessor in accessors:
                        _add_doc_var(vm, prefix + "." + accessor, _best_payload)
                        _add_global(vm, prefix + "." + accessor, _best_payload)
            pos += 1


def read_ooxml_context(data, filename, vm):
    """Populate *vm* with static OOXML values useful for emulation.

    This function is deliberately best-effort. It never raises for malformed or
    unsupported files; failures are logged at debug/warning level and normal
    emulation continues.
    """
    zf = None
    try:
        zf = _open_zip(data, filename)
    except Exception as e:
        if log.getEffectiveLevel() <= 10:
            log.debug("OOXML context: not a zip package or cannot open package: " + str(e))
        return
    if zf is None:
        return
    try:
        names = set(zf.namelist())
        if "xl/workbook.xml" not in names:
            return
        log.info("Reading OOXML workbook runtime context...")
        ctx = OOXMLWorkbookContext(zf).read()
        ctx.populate(vm)
        log.info("OOXML workbook context loaded: %d sheets, %d defined names, %d cell values, %d shape values." %
                 (len(ctx.sheet_names), len(ctx.defined_names), len(ctx.cell_values), len(ctx.shape_values)))
    except Exception as e:
        log.warning("Reading OOXML workbook runtime context failed. " + str(e))
    finally:
        try:
            zf.close()
        except Exception:
            pass
