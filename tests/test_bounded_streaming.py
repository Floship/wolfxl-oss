"""Opt-in string storage through the public API and native streaming adapters."""

from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from io import BytesIO
from zipfile import ZIP_STORED, ZipFile

import openpyxl
import pytest

import wolfxl
from wolfxl import _rust
from wolfxl._worksheet_write_only import WriteOnlyCell
from wolfxl.styles import Font


def test_inline_mode_is_per_sheet_and_keeps_values_and_styles(tmp_path, monkeypatch):
    monkeypatch.setenv("WOLFXL_PROBE_INLINE", "1")
    text = "  001 <&> é 😀\r\n\t "
    stamp = datetime(2026, 11, 24, 13, 45)
    with wolfxl.Workbook(write_only=True, inline_strings=True) as wb:
        inline = wb.create_sheet("Inline")
        shared = wb.create_sheet("Shared", inline_strings=False)
        inline.append([WriteOnlyCell(inline, text, font=Font(bold=True)), stamp, "=1+2"])
        inline.append([])
        inline.append([None, "inline only"])
        shared.append(["shared only"])
        out = tmp_path / "mixed.xlsx"
        wb.save(out)
    with ZipFile(out) as archive:
        assert b't="inlineStr"' in archive.read("xl/worksheets/sheet1.xml")
        assert b't="s"' in archive.read("xl/worksheets/sheet2.xml")
        sst = archive.read("xl/sharedStrings.xml")
        assert b"shared only" in sst
        assert b"inline only" not in sst

    with wolfxl.load_workbook(out, read_only=True, disk_shared_strings=True) as wb:
        sheet = wb["Inline"]
        assert list(sheet.values) == [
            (text, stamp, "=1+2"),
            (None, None, None),
            (None, "inline only", None),
        ]
        assert next(sheet.iter_rows())[0].font.bold
    with wolfxl.load_workbook(out, read_only=True, disk_shared_strings=True, data_only=True) as wb:
        assert next(wb["Inline"].values) == (text, stamp, None)
    reference = openpyxl.load_workbook(out, read_only=True)
    try:
        assert next(reference["Inline"].values) == (text, stamp, "=1+2")
    finally:
        reference.close()

    # The old process flag changes neither default streaming nor eager styles.
    for write_only in (False, True):
        with wolfxl.Workbook(write_only=write_only) as wb:
            sheet = wb.create_sheet() if write_only else wb.active
            sheet.append(["default"])
            wb.save(out)
        with ZipFile(out) as archive:
            assert b"default" in archive.read("xl/sharedStrings.xml")
    eager = wolfxl.Workbook()
    with pytest.raises(ValueError, match="write.only"):
        eager.create_sheet(inline_strings=True)
    with pytest.raises(ValueError, match="write_only"):
        wolfxl.Workbook(inline_strings=True)


def package(tmp_path, *, sst=None, rows=None, dimension="A1:A5"):
    """Build a normal workbook with controlled SST and sparse sheet XML."""
    base = BytesIO()
    wb = wolfxl.Workbook()
    wb.active.append(["base"])
    wb.save(base)
    wb.close()
    path = tmp_path / "strings.xlsx"
    if sst is None:
        sst = (
            "<sst><si/><si><r><t> A&amp;</t></r><rPh><t>annotation</t></rPh>"
            "<r><t><![CDATA[<B>\r\n]]></t></r><r><t>&#13;😀 </t></r></si></sst>"
        )
    if rows is None:
        rows = (
            '<row r="1"><c r="A1" t="s"><v>1</v></c></row>'
            '<row r="3"><c r="A3" t="s"><v>0</v></c></row>'
            '<row r="5"><c r="A5" t="s"><v>1</v></c></row>'
        )
    with ZipFile(base) as source, ZipFile(path, "w", ZIP_STORED) as result:
        for name in source.namelist():
            if name == "xl/sharedStrings.xml":
                body = sst
            elif name == "xl/worksheets/sheet1.xml":
                body = (
                    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                    f'<dimension ref="{dimension}"/><sheetData>{rows}</sheetData></worksheet>'
                )
            else:
                body = source.read(name)
            result.writestr(name, body)
    return path


@pytest.mark.parametrize("kind", ["path", "bytes", "buffer"])
def test_disk_strings_read_sparse_rows_with_decoded_text(tmp_path, kind):
    path = package(tmp_path)
    source = {"path": path, "bytes": path.read_bytes(), "buffer": BytesIO(path.read_bytes())}[kind]
    expected = [(" A&<B>\n\r😀 ",), (None,), ("",), (None,), (" A&<B>\n\r😀 ",)]
    with wolfxl.load_workbook(source, read_only=True, disk_shared_strings=True) as wb:
        assert list(wb.active.values) == expected
        assert [
            row for chunk in wb.active.iter_value_chunks(chunk_size=2) for row in chunk
        ] == expected
        assert [row[0].value for row in wb.active.iter_rows()] == [row[0] for row in expected]


def test_admitted_path_is_held_across_replacement_and_streaming_close(tmp_path):
    # Force the native reader to inflate a large part directly from the file.
    rows = '<row r="1"><c r="A1" t="s"><v>1</v></c></row>' + "".join(
        f'<row r="{i}"><c r="A{i}" t="s"><v>1</v></c></row>' for i in range(2, 30_001)
    )
    path = package(tmp_path, rows=rows, dimension="A1:A30000")
    wb = wolfxl.load_workbook(path, read_only=True, disk_shared_strings=True)
    replacement = path.with_suffix(".new")
    replacement.write_bytes(b"replaced")
    replacement.replace(path)
    rows = wb.active.values
    assert next(rows) == (" A&<B>\n\r😀 ",)
    wb.close()
    assert sum(1 for _ in rows) == 29_999


def test_native_sparse_inline_round_trip(tmp_path):
    writer = _rust.NativeWorkbook()
    writer.add_sheet("Data")
    writer.enable_streaming_sheet("Data", inline_strings=True)
    with pytest.raises(ValueError, match="string mode"):
        writer.enable_streaming_sheet("Data", inline_strings=False)
    writer.append_streaming_row("Data", 3, [{"type": "string", "value": "=literal"}])
    writer.append_streaming_row("Data", 5, [None, {"type": "string", "value": "tail"}])
    path = tmp_path / "native.xlsx"
    writer.save(str(path))
    with wolfxl.load_workbook(path, read_only=True, disk_shared_strings=True) as wb:
        assert list(wb.active.values) == [
            (None, None),
            (None, None),
            ("=literal", None),
            (None, None),
            (None, "tail"),
        ]


def test_disk_table_can_be_shared_across_independent_native_readers(tmp_path):
    path = package(tmp_path)
    table = _rust.StreamingSharedStrings.load_disk(str(path))
    assert len(table) == 2

    def read():
        reader = _rust.StreamingSheetReader.open(str(path), "Sheet", shared_strings=table)
        try:
            return reader.read_next_values_compact_chunk()
        finally:
            reader.close()

    with ThreadPoolExecutor(max_workers=4) as pool:
        results = list(pool.map(lambda _: read(), range(16)))
    expected = ([(1, (" A&<B>\n\r😀 ",)), (3, ("",)), (5, (" A&<B>\n\r😀 ",))], 1, 5, False)
    assert results == [expected] * 16


def test_disk_mode_keeps_existing_budgets_and_rejects_unsupported_modes(tmp_path):
    path = package(tmp_path)
    with pytest.raises(wolfxl.LimitExceeded, match="max_cells"):
        wolfxl.load_workbook(
            path, read_only=True, disk_shared_strings=True, limits=wolfxl.Limits(max_cells=1)
        )
    for kwargs in ({}, {"read_only": True, "modify": True}, {"read_only": True, "rich_text": True}):
        with pytest.raises(ValueError, match="requires read_only"):
            wolfxl.load_workbook(path, disk_shared_strings=True, **kwargs)


def test_disk_mode_bounds_xml_tokens_and_split_shared_text(tmp_path):
    path = package(tmp_path, sst="<sst><!--" + "X" * (1024 * 1024) + "--></sst>")
    with pytest.raises(OSError, match="XML token"):
        wolfxl.load_workbook(path, read_only=True, disk_shared_strings=True)
    with pytest.raises(OSError, match="XML token"):
        _rust.StreamingSharedStrings.load_disk(str(path))

    run = "<r><t>" + "é" * 32768 + "</t></r>"
    sst = "<sst><si/>" + "<si>" + run * 2 + "</si></sst>"
    path = package(tmp_path, sst=sst)
    with wolfxl.load_workbook(path, read_only=True, disk_shared_strings=True) as wb:
        assert next(wb.active.values) == ("é" * 65536,)
    path = package(tmp_path, sst=sst.replace("</si>", "<r><t>x</t></r></si>"))
    with pytest.raises(OSError, match="128 KiB"):
        wolfxl.load_workbook(path, read_only=True, disk_shared_strings=True)


def test_disk_mode_bounds_decoded_references_and_sparse_dimensions(tmp_path):
    sst = "<sst><si><t>" + "x" * (128 * 1024) + "</t></si></sst>"
    cells = '<c t="s"><v>0</v></c>' * 65
    path = package(tmp_path, sst=sst, rows=f'<row r="1">{cells}</row>')
    with wolfxl.load_workbook(path, read_only=True, disk_shared_strings=True) as wb:
        with pytest.raises(OSError, match="8 MiB"):
            next(wb.active.values)
    with wolfxl.load_workbook(path, read_only=True, disk_shared_strings=True) as wb:
        with pytest.raises(ValueError, match="Excel worksheet bounds"):
            next(wb.active.iter_rows(values_only=True, max_col=16_385))


def test_inline_save_to_binary_stream_handles_short_writes():
    class PartialWriter(BytesIO):
        def write(self, data):
            return super().write(data[:64])

    target = PartialWriter()
    with wolfxl.Workbook(write_only=True, inline_strings=True) as wb:
        wb.create_sheet().append(["inline"])
        wb.save(target)

    with wolfxl.load_workbook(target.getvalue(), read_only=True) as wb:
        assert list(wb.active.values) == [("inline",)]


def test_inline_save_qualifies_xlookup_without_rescanning_sheet(tmp_path, monkeypatch):
    from wolfxl import _formula_cache

    def unexpected_scan(*args, **kwargs):
        pytest.fail("save read worksheet XML after native streaming output")

    monkeypatch.setattr(_formula_cache, "_mentions_xlookup", unexpected_scan)
    path = tmp_path / "formulas.xlsx"
    with wolfxl.Workbook(write_only=True, inline_strings=True) as wb:
        sheet = wb.create_sheet()
        sheet.append(["XLOOKUP", "=IF(TRUE,XLOOKUP(1,A1:A2,B1:B2),0)"])
        wb.save(path)

    with wolfxl.load_workbook(path, read_only=True, disk_shared_strings=True) as wb:
        assert next(wb.active.values) == ("XLOOKUP", "=IF(TRUE,_xlfn.XLOOKUP(1,A1:A2,B1:B2),0)")


def test_disk_read_preserves_cached_values_of_formulas(tmp_path):
    rows = (
        '<row r="1"><c r="A1" t="s"><f>Lookup()</f><v>1</v></c></row>'
        '<row r="3"><c r="A3"><f>1+2</f><v>3</v></c></row>'
    )
    path = package(tmp_path, rows=rows)
    with wolfxl.load_workbook(path, read_only=True, disk_shared_strings=True) as wb:
        assert list(wb.active.values) == [("=Lookup()",), (None,), ("=1+2",)]
    with wolfxl.load_workbook(path, read_only=True, disk_shared_strings=True, data_only=True) as wb:
        assert list(wb.active.values) == [(" A&<B>\n\r😀 ",), (None,), (3,)]
