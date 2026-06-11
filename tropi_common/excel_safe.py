"""SafeWorkbook — OOXML-safe ZIP/XML xlsx editor (fleet union build).

Union of the 8 per-service copies (2026-06-11 diff matrix): keyaccounts base
+ two warehouse-receipts grafts (cell_str xml:space-preserve fix,
shift_formula_refs array-formula shifter). Strict superset of every copy it
replaces. See EXCEL_SAFE_DIFF_MATRIX.md in ~/dev for the per-method lineage.
"""

import io
import posixpath
import re
import zipfile
from datetime import date, datetime
from xml.sax.saxutils import escape as xml_escape


# Files that cause corruption when repacked — always dropped on save
_DROP_FILES = {"xl/calcChain.xml", "docProps/thumbnail.emf"}


class SafeWorkbook:
    """Safe Excel workbook manipulation via ZIP + regex.

    Opens an .xlsx file as a ZIP archive, provides methods to read/modify
    sheet XML and other internal files, and repacks with all OOXML
    corruption prevention applied automatically.
    """

    def __init__(self, xlsx_bytes: bytes):
        self._infos: dict[str, zipfile.ZipInfo] = {}
        self._data: dict[str, bytes] = {}
        self._sheet_paths: dict[str, str] = {}  # sheet name -> ZIP path

        with zipfile.ZipFile(io.BytesIO(xlsx_bytes), "r") as zf:
            for item in zf.infolist():
                self._infos[item.filename] = item
                self._data[item.filename] = zf.read(item.filename)

        self._parse_sheet_paths()

    def _parse_sheet_paths(self):
        """Build sheet name -> ZIP file path mapping from workbook.xml + rels."""
        wb_xml = self._read_internal("xl/workbook.xml")
        rels_xml = self._read_internal("xl/_rels/workbook.xml.rels")

        for m in re.finditer(
            r'<sheet\s+name="([^"]*)"\s+sheetId="\d+"\s+(?:state="[^"]*"\s+)?r:id="(rId\d+)"',
            wb_xml,
        ):
            name, rid = m.group(1), m.group(2)
            # Find Relationship element with this rId (attributes may be in any order)
            rel_el = re.search(rf'<Relationship\b[^>]*\bId="{rid}"[^>]*/?\s*>', rels_xml)
            rel_m = re.search(r'Target="([^"]*)"', rel_el.group()) if rel_el else None
            if rel_m:
                target = rel_m.group(1)
                path = f"xl/{target}" if not target.startswith("/") else target.lstrip("/")
                self._sheet_paths[name] = path

    def _read_internal(self, path: str) -> str:
        """Read an internal ZIP file as UTF-8 string."""
        if path in self._data:
            return self._data[path].decode("utf-8")
        return ""

    def _write_internal(self, path: str, content: str):
        """Write a UTF-8 string to an internal ZIP file."""
        self._data[path] = content.encode("utf-8")

    # ── File-level access (for project-specific needs) ────────────

    def read_file(self, path: str) -> str:
        """Read any internal file as UTF-8 string (e.g., 'xl/styles.xml')."""
        return self._read_internal(path)

    def write_file(self, path: str, content: str):
        """Write any internal file (creates new or overwrites existing)."""
        if path not in self._infos:
            self._infos[path] = zipfile.ZipInfo(path)
        self._write_internal(path, content)

    def file_exists(self, path: str) -> bool:
        """Check if an internal file exists."""
        return path in self._data

    def filenames(self) -> list[str]:
        """List all internal file paths."""
        return list(self._data.keys())

    # ── Sheet access ──────────────────────────────────────────────

    def sheet_names(self) -> list[str]:
        """Return list of sheet names in workbook order."""
        return list(self._sheet_paths.keys())

    def sheet_xml(self, name: str) -> str:
        """Get raw XML string for a sheet by name."""
        path = self._sheet_paths.get(name)
        if not path or path not in self._data:
            raise ValueError(f"Sheet '{name}' not found")
        return self._data[path].decode("utf-8")

    def set_sheet_xml(self, name: str, xml: str):
        """Update the XML for a sheet."""
        path = self._sheet_paths.get(name)
        if not path:
            raise ValueError(f"Sheet '{name}' not found")
        self._data[path] = xml.encode("utf-8")

    def sheet_path(self, name: str) -> str:
        """Get the internal ZIP path for a sheet."""
        path = self._sheet_paths.get(name)
        if not path:
            raise ValueError(f"Sheet '{name}' not found")
        return path

    def rename_sheet(self, old_name: str, new_name: str):
        """Rename a sheet in workbook.xml, definedNames, and app.xml."""
        if old_name not in self._sheet_paths:
            raise ValueError(f"Sheet '{old_name}' not found")

        # Update workbook.xml
        wb_xml = self._read_internal("xl/workbook.xml")
        wb_xml = wb_xml.replace(f'name="{old_name}"', f'name="{new_name}"')
        wb_xml = wb_xml.replace(f"{old_name}!", f"'{new_name}'!")
        self._write_internal("xl/workbook.xml", wb_xml)

        # Update app.xml if it exists
        if self.file_exists("docProps/app.xml"):
            app_xml = self._read_internal("docProps/app.xml")
            app_xml = app_xml.replace(
                f"<vt:lpstr>{xml_escape(old_name)}</vt:lpstr>",
                f"<vt:lpstr>{xml_escape(new_name)}</vt:lpstr>",
            )
            # Also update Print_Area references
            app_xml = app_xml.replace(
                f"<vt:lpstr>{old_name}!Print_Area</vt:lpstr>",
                f"<vt:lpstr>'{new_name}'!Print_Area</vt:lpstr>",
            )
            self._write_internal("docProps/app.xml", app_xml)

        # Update internal mapping
        self._sheet_paths[new_name] = self._sheet_paths.pop(old_name)

    # ── Table access ──────────────────────────────────────────────

    def table_xml(self, sheet_name: str) -> str | None:
        """Read table XML associated with a sheet, if any."""
        table_path = self._find_table_path(sheet_name)
        if table_path and table_path in self._data:
            return self._data[table_path].decode("utf-8")
        return None

    def set_table_xml(self, sheet_name: str, xml: str):
        """Update table XML associated with a sheet."""
        table_path = self._find_table_path(sheet_name)
        if not table_path:
            raise ValueError(f"No table found for sheet '{sheet_name}'")
        self._data[table_path] = xml.encode("utf-8")

    def _find_table_path(self, sheet_name: str) -> str | None:
        """Find the ZIP path of the table associated with a sheet."""
        path = self._sheet_paths.get(sheet_name)
        if not path:
            return None
        sheet_basename = path.split("/")[-1].replace(".xml", "")
        rels_path = f"xl/worksheets/_rels/{sheet_basename}.xml.rels"
        if rels_path not in self._data:
            return None
        rels = self._data[rels_path].decode("utf-8")
        m = re.search(r'Target="([^"]*table[^"]*)"', rels, re.IGNORECASE)
        if not m:
            return None
        target = m.group(1)
        if target.startswith("/"):
            return target.lstrip("/")
        base_dir = posixpath.dirname(path)
        return posixpath.normpath(posixpath.join(base_dir, target))

    # ── Shared strings ────────────────────────────────────────────

    def read_shared_strings(self) -> list[str]:
        """Parse shared strings table into a list of decoded values.

        Useful for reading existing cell values that use t="s" references.
        For writing, prefer inline strings (cell_str) — they always work.
        """
        from xml.sax.saxutils import unescape as xml_unescape

        ss_xml = self._read_internal("xl/sharedStrings.xml")
        if not ss_xml:
            return []

        _UNESCAPE_EXTRA = {"&apos;": "'", "&quot;": '"'}
        values = []
        for m in re.finditer(r'<si[^>]*>(.*?)</si>', ss_xml, re.DOTALL):
            texts = re.findall(r'<t[^>]*>(.*?)</t>', m.group(1), re.DOTALL)
            raw = "".join(texts)
            values.append(xml_unescape(raw, _UNESCAPE_EXTRA))

        return values

    def add_shared_strings(self, new_values: list[str]):
        """Add new entries to the shared strings table and update counts.

        Returns a dict {value: index} for the newly added strings.
        """
        ss_xml = self._read_internal("xl/sharedStrings.xml")
        if not ss_xml:
            return {}

        existing = self.read_shared_strings()
        base_idx = len(existing)
        result = {}

        new_entries = ""
        for i, val in enumerate(new_values):
            new_entries += f"<si><t>{xml_escape(str(val))}</t></si>"
            result[val] = base_idx + i

        # Insert before closing tag
        ss_xml = ss_xml.replace("</sst>", new_entries + "</sst>")

        # Update counts
        total = base_idx + len(new_values)
        ss_xml = re.sub(r'count="\d+"', f'count="{total}"', ss_xml, count=1)
        ss_xml = re.sub(r'uniqueCount="\d+"', f'uniqueCount="{total}"', ss_xml, count=1)

        self._write_internal("xl/sharedStrings.xml", ss_xml)
        return result

    # ── Workbook-level settings ───────────────────────────────────

    def set_full_calc_on_load(self):
        """Force Excel to recalculate all formulas when the file is opened."""
        wb_xml = self._read_internal("xl/workbook.xml")
        if "<calcPr" in wb_xml:
            if 'fullCalcOnLoad="1"' not in wb_xml:
                wb_xml = re.sub(
                    r'<calcPr\b', '<calcPr fullCalcOnLoad="1"', wb_xml
                )
        else:
            wb_xml = wb_xml.replace(
                "</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>'
            )
        self._write_internal("xl/workbook.xml", wb_xml)

    # ── External link removal ────────────────────────────────────

    def strip_external_links(self):
            """Remove all external link artifacts from the workbook.
    
            Conflicted copies and other external references embed
            external link files, relationships, and [N] formula references
            that corrupt the file. This method removes ALL of them:
    
            1. Drops xl/externalLinks/ files and rels
            2. Removes externalLink overrides from [Content_Types].xml
            3. Removes externalLink relationships from workbook.xml.rels
            4. Removes <externalReferences> block from workbook.xml
            5. Removes duplicate definedNames with localSheetId
            6. Strips [N] external book refs from ALL .xml files
               (sheets, tables, workbook — everything)
    
            Returns True if any modifications were made.
            """
            modified = False
    
            # 1. Drop externalLinks files
            to_drop = [f for f in self._data if "externalLink" in f]
            for f in to_drop:
                del self._data[f]
                if f in self._infos:
                    del self._infos[f]
                modified = True
    
            # 2. Clean [Content_Types].xml
            ct_key = "[Content_Types].xml"
            if ct_key in self._data:
                ct = self._data[ct_key].decode("utf-8")
                new_ct = re.sub(
                    r'<Override[^>]*PartName="[^"]*externalLink[^"]*"[^>]*/>\s*',
                    "", ct,
                )
                if new_ct != ct:
                    self._data[ct_key] = new_ct.encode("utf-8")
                    modified = True
    
            # 3. Clean workbook.xml.rels
            wb_rels = "xl/_rels/workbook.xml.rels"
            if wb_rels in self._data:
                rels = self._data[wb_rels].decode("utf-8")
                new_rels = re.sub(
                    r'<Relationship[^>]*externalLink[^>]*/>\s*', "", rels
                )
                if new_rels != rels:
                    self._data[wb_rels] = new_rels.encode("utf-8")
                    modified = True
    
            # 4 & 5. Clean workbook.xml
            wb_key = "xl/workbook.xml"
            if wb_key in self._data:
                wb = self._data[wb_key].decode("utf-8")
                new_wb = re.sub(
                    r'<externalReferences>.*?</externalReferences>\s*',
                    "", wb, flags=re.DOTALL,
                )
                # Remove duplicate definedNames with localSheetId
                new_wb = re.sub(
                    r'<definedName name="(ArtikuliList|KlientiList|ShtampiList)"'
                    r' localSheetId="\d+">[^<]*</definedName>',
                    "", new_wb,
                )
                if new_wb != wb:
                    self._data[wb_key] = new_wb.encode("utf-8")
                    modified = True
    
            # 6. Strip [N] from ALL .xml files (sheets, tables, everything)
            for filename in list(self._data.keys()):
                if not filename.endswith(".xml"):
                    continue
                content = self._data[filename].decode("utf-8")
                if re.search(r'\[\d+\]', content):
                    self._data[filename] = re.sub(
                        r'\[\d+\]', "", content
                    ).encode("utf-8")
                    modified = True
    
            # Rebuild sheet paths since files may have been removed
            if modified:
                self._parse_sheet_paths()
    
            return modified

    # ── Save / repack ─────────────────────────────────────────────

    def save(self) -> bytes:
        """Repack the ZIP with all OOXML corruption prevention applied.

        Automatically:
        - Drops calcChain.xml and thumbnail.emf
        - Cleans all references to dropped files
        - Preserves ZipInfo metadata for unmodified files
        """
        # Clean Content_Types
        ct_key = "[Content_Types].xml"
        if ct_key in self._data:
            ct = self._data[ct_key].decode("utf-8")
            ct = re.sub(
                r'<Override[^>]*PartName="/xl/calcChain\.xml"[^>]*/>\s*',
                "", ct,
            )
            # Only strip the emf Default if no .emf parts remain in the package.
            # Otherwise we orphan xl/media/imageN.emf → Excel "errors detected" warning.
            remaining_emf = [f for f in self._data
                             if f.lower().endswith(".emf")]
            if not remaining_emf:
                ct = re.sub(r'<Default[^>]*Extension="emf"[^>]*/>\s*', "", ct)
            self._data[ct_key] = ct.encode("utf-8")

        # Clean calcChain from workbook rels
        wb_rels = "xl/_rels/workbook.xml.rels"
        if wb_rels in self._data:
            rels = self._data[wb_rels].decode("utf-8")
            rels = re.sub(
                r'<Relationship[^>]*Target="calcChain\.xml"[^>]*/>\s*',
                "", rels,
            )
            self._data[wb_rels] = rels.encode("utf-8")

        # Clean thumbnail from root rels
        root_rels = "_rels/.rels"
        if root_rels in self._data:
            rels = self._data[root_rels].decode("utf-8")
            rels = re.sub(
                r'<Relationship[^>]*Target="docProps/thumbnail\.emf"[^>]*/>\s*',
                "", rels,
            )
            self._data[root_rels] = rels.encode("utf-8")

        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for filename, info in self._infos.items():
                if filename in _DROP_FILES:
                    continue
                zf.writestr(info, self._data[filename])

        return buf.getvalue()

    # ── Static: cell XML builders ─────────────────────────────────

    @staticmethod
    def cell_str(col: str, row: int, value: str, style: str = "") -> str:
            """Inline string cell: <c r="A1" s="4" t="inlineStr"><is><t>...</t></is></c>
    
            If the value has leading/trailing whitespace we add
            `xml:space="preserve"` to the `<t>` tag — otherwise XML parsers
            (including Excel) strip those characters on load. Debugged
            2026-04-21 in `railway-warehouse-receipts`: `"Кg "` (trailing
            space, 4 chars) was being saved as `"Кg"` (3 chars), silently
            breaking formulas that did exact-string match on the unit.
            """
            s = f' s="{style}"' if style else ""
            val_str = str(value)
            t_attrs = ' xml:space="preserve"' if val_str != val_str.strip() else ""
            return (
                f'<c r="{col}{row}"{s} t="inlineStr">'
                f'<is><t{t_attrs}>{xml_escape(val_str)}</t></is></c>'
            )

    @staticmethod
    def cell_num(col: str, row: int, value: float | int, style: str = "") -> str:
        """Numeric cell: <c r="A1" s="4"><v>42.5</v></c>"""
        s = f' s="{style}"' if style else ""
        return f'<c r="{col}{row}"{s}><v>{value}</v></c>'

    @staticmethod
    def cell_formula(col: str, row: int, formula: str, style: str = "") -> str:
        """Formula cell: <c r="A1" s="4"><f>SUM(B1:B10)</f></c>"""
        s = f' s="{style}"' if style else ""
        return f'<c r="{col}{row}"{s}><f>{xml_escape(formula)}</f></c>'

    @staticmethod
    def cell_empty(col: str, row: int, style: str = "") -> str:
        """Empty cell with style preserved: <c r="A1" s="4"/>"""
        s = f' s="{style}"' if style else ""
        return f'<c r="{col}{row}"{s}/>'

    @staticmethod
    def cell_shared_str(col: str, row: int, index: int, style: str = "") -> str:
        """Shared string reference cell: <c r="A1" s="4" t="s"><v>0</v></c>"""
        s = f' s="{style}"' if style else ""
        return f'<c r="{col}{row}"{s} t="s"><v>{index}</v></c>'

    # ── Static: sheet XML manipulation ────────────────────────────

    @staticmethod
    def get_style(sheet_xml: str, col: str, row: int) -> str:
        """Extract style index (s="N") from an existing cell. Returns "" if none."""
        pat = re.compile(
            rf'<c r="{col}{row}"[^>]*/>|<c r="{col}{row}"[^>]*>.*?</c>',
            re.DOTALL,
        )
        m = pat.search(sheet_xml)
        if m:
            s_m = re.search(r's="(\d+)"', m.group())
            if s_m:
                return s_m.group(1)
        return ""

    @staticmethod
    def get_row_styles(sheet_xml: str, row: int) -> dict[str, str]:
        """Extract {column_letter: style_index} from all cells in a row.

        Useful for preserving template row styles when building new rows.
        """
        row_m = re.search(
            rf'<row r="{row}"[^>]*>(.*?)</row>', sheet_xml, re.DOTALL
        )
        if not row_m:
            return {}
        styles = {}
        for cm in re.finditer(
            r'<c r="([A-Z]+)\d+"[^>]*\ss="(\d+)"', row_m.group(1)
        ):
            styles[cm.group(1)] = cm.group(2)
        return styles

    @staticmethod
    def inject_cell(sheet_xml: str, col: str, row: int, cell_xml: str) -> str:
        """Replace an existing cell, insert into existing row, or create a new row.

        Three-tier fallback so empty cells in completely empty rows still
        accept writes (e.g. signature-block date cells in templates).
        """
        # Tier 1: cell already exists — replace
        pat = re.compile(
            rf'<c r="{col}{row}"[^>]*/>|<c r="{col}{row}"[^>]*>.*?</c>',
            re.DOTALL,
        )
        if pat.search(sheet_xml):
            return pat.sub(cell_xml, sheet_xml)
        # Tier 2: row exists, cell missing — insert into row
        row_re = re.compile(
            rf'(<row\b[^>]*r="{row}"[^>]*>)(.*?)(</row>)', re.DOTALL
        )
        m = row_re.search(sheet_xml)
        if m:
            return (
                sheet_xml[: m.start(2)]
                + m.group(2) + cell_xml
                + sheet_xml[m.end(2):]
            )
        # Tier 3: row missing — create <row> and insert it in numerical order
        new_row = f'<row r="{row}">{cell_xml}</row>'
        # Find the next-higher row to insert before; if none, insert before </sheetData>
        all_rows = list(re.finditer(r'<row\b[^>]*r="(\d+)"', sheet_xml))
        for rm in all_rows:
            if int(rm.group(1)) > row:
                # Find the row's <row ...> opening tag start
                row_start = rm.start()
                return sheet_xml[:row_start] + new_row + sheet_xml[row_start:]
        # No higher row — append before </sheetData>
        return re.sub(r'(</sheetData>)', new_row + r'\1', sheet_xml, count=1)

    @staticmethod
    def sort_cells(sheet_xml: str) -> str:
        """Sort cells within each row by column index (OOXML requirement).

        Call this after injecting cells to ensure column order compliance.
        """

        def _col_to_num(col: str) -> int:
            n = 0
            for ch in col:
                n = n * 26 + (ord(ch) - ord("A") + 1)
            return n

        def _sort_row(m: re.Match) -> str:
            tag, content, end = m.group(1), m.group(2), m.group(3)
            cell_re = re.compile(
                r'(<c r="([A-Z]+)\d+"[^>]*/>|<c r="([A-Z]+)\d+"[^>]*>.*?</c>)',
                re.DOTALL,
            )
            cells = []
            for cm in cell_re.finditer(content):
                col = cm.group(2) or cm.group(3)
                cells.append((_col_to_num(col), cm.group(1)))
            if not cells:
                return m.group(0)
            cleaned = cell_re.sub("", content)
            cells.sort(key=lambda x: x[0])
            return tag + cleaned + "".join(c for _, c in cells) + end

        return re.sub(
            r'(<row\b[^>]*>)(.*?)(</row>)',
            _sort_row,
            sheet_xml,
            flags=re.DOTALL,
        )

    @staticmethod
    def update_dimension(sheet_xml: str, max_row: int | None = None) -> str:
        """Update <dimension> to cover all rows.

        If max_row is None, scans the XML for the highest row number.
        """
        if max_row is None:
            rows = [
                int(m.group(1))
                for m in re.finditer(r'<row\b[^>]*r="(\d+)"', sheet_xml)
            ]
            max_row = max(rows) if rows else 1

        return re.sub(
            r'(<dimension ref="[A-Z]+\d+:[A-Z]+)\d+(")',
            rf'\g<1>{max_row}\2',
            sheet_xml,
        )

    @staticmethod
    def bump_table_ref(table_xml: str, delta: int) -> str:
        """Bump max row in <table> and <autoFilter> refs by delta rows."""

        def _bump(m):
            return f"{m.group(1)}{int(m.group(2)) + delta}{m.group(3)}"

        table_xml = re.sub(
            r'(<table[^>]+ref="[A-Z]+\d+:[A-Z]+)(\d+)(")', _bump, table_xml
        )
        table_xml = re.sub(
            r'(<autoFilter[^>]+ref="[A-Z]+\d+:[A-Z]+)(\d+)(")',
            _bump,
            table_xml,
        )
        return table_xml

    @staticmethod
    def build_row(
        row: int,
        cells: list[str],
        spans: str = "",
        height: float = 0,
        extra_attrs: str = "",
    ) -> str:
        """Build a complete <row> element from a list of cell XML strings."""
        attrs = f'r="{row}"'
        if spans:
            attrs += f' spans="{spans}"'
        if height:
            attrs += f' ht="{height}" customHeight="1"'
        if extra_attrs:
            attrs += f" {extra_attrs}"
        return f'<row {attrs}>{"".join(cells)}</row>'

    @staticmethod
    def insert_rows_before(
        sheet_xml: str, before_marker: str, rows_xml: list[str]
    ) -> str:
        """Insert row XML strings before a marker string in sheet XML.

        Falls back to inserting before </sheetData> if marker not found.
        """
        pos = sheet_xml.find(before_marker)
        if pos >= 0:
            return (
                sheet_xml[:pos] + "".join(rows_xml) + sheet_xml[pos:]
            )
        return sheet_xml.replace(
            "</sheetData>", "".join(rows_xml) + "</sheetData>"
        )

    @staticmethod
    def shift_rows(sheet_xml: str, from_row: int, delta: int) -> str:
        """Shift all rows with r >= from_row by delta.

        Renumbers both row r="" attributes and cell r="COL{row}" references.
        """

        def _shift(m):
            old_r = int(m.group(1))
            if old_r >= from_row:
                new_r = old_r + delta
                result = re.sub(
                    rf'\br="{old_r}"', f'r="{new_r}"', m.group(0)
                )
                result = re.sub(
                    rf'r="([A-Z]{{1,3}}){old_r}"',
                    rf'r="\g<1>{new_r}"',
                    result,
                )
                return result
            return m.group(0)

        return re.sub(
            r'<row r="(\d+)"(?:[^>]*?/>|[^>]*?>.*?</row>)',
            _shift,
            sheet_xml,
            flags=re.DOTALL,
        )

    @staticmethod
    def shift_formula_refs(sheet_xml: str, from_row: int, delta: int) -> str:
            """Bump row numbers inside `<f ... ref="...">` attributes.
    
            CALL THIS AFTER `shift_rows()` whenever the sheet may contain array
            formulas (`<f t="array" ref="V425">…</f>`) or shared-formula masters
            (`<f t="shared" ref="A1:A10" si="0">…</f>`).
    
            `shift_rows()` only updates the `<c r="…">` wrapper. If the `ref=`
            inside `<f>` still points at the pre-shift row, Excel on open marks
            the formula as corrupt and SILENTLY STRIPS IT. The repair log reads:
            "Removed Records: Formula from /xl/worksheets/sheet1.xml part".
    
            Debugged 2026-04-20 while building `railway-warehouse-receipts`:
            the receipts workbook had `<f t="array" ref="V425">x</f>` anchored
            at V425; shifting to V475 without updating ref="V425" → V475 caused
            Excel's open-time validator to drop the formula.
            """
            def _shift_one(ref: str) -> str:
                parts = ref.split(":")
                out = []
                for part in parts:
                    m = re.match(r'^(\$?)([A-Z]+)(\$?)(\d+)$', part)
                    if not m:
                        out.append(part)
                        continue
                    dc, col, dr, row = m.group(1), m.group(2), m.group(3), int(m.group(4))
                    if row >= from_row:
                        row += delta
                    out.append(f"{dc}{col}{dr}{row}")
                return ":".join(out)
    
            return re.sub(
                r'(<f\b[^>]*?\bref=")([^"]+)(")',
                lambda m: m.group(1) + _shift_one(m.group(2)) + m.group(3),
                sheet_xml,
            )

    # ── Static: shared formula expansion ─────────────────────────

    @staticmethod
    def expand_shared_formulas(sheet_xml: str) -> str:
        """Expand ALL shared formula groups into individual normal formulas.

        MUST be called BEFORE any inject_cell() on sheets with shared formulas.
        Prevents orphaned slave cells that corrupt the workbook.

        How it works:
        1. Finds every master cell: <f ... ref="A1:A10" t="shared" si="N">FORMULA</f>
        2. Finds every slave cell: <f t="shared" si="N"/>  (self-closing, no formula)
        3. Computes the shifted formula for each slave (row/col offset from master)
        4. Replaces both masters and slaves with plain <f>FORMULA</f>
        """
        def _col_to_num(col: str) -> int:
            n = 0
            for ch in col:
                n = n * 26 + (ord(ch) - 64)
            return n

        def _num_to_col(n: int) -> str:
            r = ""
            while n > 0:
                n, rem = divmod(n - 1, 26)
                r = chr(65 + rem) + r
            return r

        def _shift_formula(formula: str, row_delta: int, col_delta: int) -> str:
            """Shift cell references in a formula by row/col delta.
            Respects $ (absolute) markers."""
            def _shift_ref(m):
                dollar_c = m.group(1)  # '$' or ''
                col = m.group(2)
                dollar_r = m.group(3)  # '$' or ''
                row = int(m.group(4))
                nc = col if dollar_c else _num_to_col(_col_to_num(col) + col_delta)
                nr = row if dollar_r else row + row_delta
                if nr < 1:
                    nr = 1
                return f"{dollar_c}{nc}{dollar_r}{nr}"
            return re.sub(r'(\$?)([A-Z]{1,3})(\$?)(\d+)', _shift_ref, formula)

        # ── Step 1: collect masters ──────────────────────────────
        # Match whole <c>...</c> cells, then check for shared master <f>
        masters = {}  # si -> {col, row, formula}
        for m in re.finditer(
            r'<c r="([A-Z]+)(\d+)"[^>]*>(.*?)</c>',
            sheet_xml, re.DOTALL
        ):
            content = m.group(3)
            f_m = re.search(r'<f\b([^>]*)>(.*?)</f>', content, re.DOTALL)
            if not f_m:
                continue
            attrs = f_m.group(1)
            if 't="shared"' in attrs and 'ref=' in attrs:
                si_m = re.search(r'si="(\d+)"', attrs)
                if si_m:
                    masters[si_m.group(1)] = {
                        "col": m.group(1),
                        "row": int(m.group(2)),
                        "formula": f_m.group(2),
                    }

        if not masters:
            return sheet_xml

        # ── Step 2: convert masters to normal formulas ───────────
        # Remove ref="...", t="shared", si="N" from master <f> tags
        def _clean_master_f(m):
            attrs = m.group(1)
            formula = m.group(2)
            cleaned = re.sub(r'\s*ref="[^"]*"', '', attrs)
            cleaned = re.sub(r'\s*t="shared"', '', cleaned)
            cleaned = re.sub(r'\s*si="\d+"', '', cleaned)
            cleaned = cleaned.strip()
            tag = f"<f{' ' + cleaned if cleaned else ''}>"
            return f"{tag}{formula}</f>"

        sheet_xml = re.sub(
            r'<f\b([^>]*t="shared"[^>]*ref="[^"]*"[^>]*)>(.*?)</f>',
            _clean_master_f, sheet_xml, flags=re.DOTALL)
        sheet_xml = re.sub(
            r'<f\b([^>]*ref="[^"]*"[^>]*t="shared"[^>]*)>(.*?)</f>',
            _clean_master_f, sheet_xml, flags=re.DOTALL)

        # ── Step 3: expand slaves to normal formulas ─────────────
        # Slaves: <f t="shared" si="N"/> (self-closing, inside a <c> element)
        def _expand_slave(m):
            cell_tag = m.group(1)    # <c r="XX99" ...>
            f_attrs = m.group(2)     # attributes of <f>
            rest = m.group(3)        # </c> or <v>...</v></c>

            si_m = re.search(r'si="(\d+)"', f_attrs)
            if not si_m or si_m.group(1) not in masters:
                return m.group(0)    # unknown si, leave as-is

            info = masters[si_m.group(1)]
            ref_m = re.search(r'r="([A-Z]+)(\d+)"', cell_tag)
            if not ref_m:
                return m.group(0)

            slave_col, slave_row = ref_m.group(1), int(ref_m.group(2))
            rd = slave_row - info["row"]
            cd = _col_to_num(slave_col) - _col_to_num(info["col"])
            shifted = _shift_formula(info["formula"], rd, cd)

            return f"{cell_tag}<f>{shifted}</f>{rest}"

        sheet_xml = re.sub(
            r'(<c r="[A-Z]+\d+"[^>]*>)\s*<f\b([^>]*t="shared"[^>]*)/>\s*((?:<v(?:\s*/>|\s*>[^<]*</v>)\s*)?</c>)',
            _expand_slave, sheet_xml, flags=re.DOTALL)

        # ── Step 4: fix orphaned slaves (master already gone) ────
        # Find remaining slaves, infer formula from nearest normal
        # formula in the same column
        orphan_re = re.compile(
            r'(<c r="([A-Z]+)(\d+)"[^>]*>)\s*'
            r'<f\b[^>]*t="shared"[^>]*/>\s*'
            r'((?:<v(?:\s*/>|\s*>[^<]*</v>)\s*)?</c>)',
            re.DOTALL,
        )
        if orphan_re.search(sheet_xml):
            # Build column→[(row, formula)] index from normal formulas
            col_formulas: dict[str, list[tuple[int, str]]] = {}
            for cm in re.finditer(
                r'<c r="([A-Z]+)(\d+)"[^>]*>(?:(?!</c>).)*?'
                r'<f>([^<]+)</f>.*?</c>',
                sheet_xml, re.DOTALL,
            ):
                col_formulas.setdefault(cm.group(1), []).append(
                    (int(cm.group(2)), cm.group(3))
                )
            for v in col_formulas.values():
                v.sort()

            def _fix_orphan(m):
                cell_tag = m.group(1)
                col = m.group(2)
                row = int(m.group(3))
                rest = m.group(4)

                nearby = col_formulas.get(col, [])
                if not nearby:
                    return m.group(0)  # no reference formula, leave
                # Find closest formula above (or below)
                best = min(nearby, key=lambda x: abs(x[0] - row))
                base_row, base_formula = best
                rd = row - base_row
                shifted = _shift_formula(base_formula, rd, 0)
                return f"{cell_tag}<f>{shifted}</f>{rest}"

            sheet_xml = orphan_re.sub(_fix_orphan, sheet_xml)

        return sheet_xml

    # ── Static: conversions ───────────────────────────────────────

    @staticmethod
    def date_to_serial(d: date | datetime) -> int:
        """Convert Python date/datetime to Excel serial number (integer days)."""
        if isinstance(d, datetime):
            d = d.date()
        return (d - date(1899, 12, 30)).days

    @staticmethod
    def col_to_num(col: str) -> int:
        """Column letters to number: A=1, Z=26, AA=27, AZ=52, etc."""
        n = 0
        for ch in col:
            n = n * 26 + (ord(ch) - ord("A") + 1)
        return n
