"""Preparation preserves unchanged markup while validating/transcoding XML."""

from io import BytesIO
from zipfile import ZipFile

import pytest
from lxml import etree

from carbotek import Engine
from carbotek.errors import CorruptedTemplateError
from carbotek.package import TemplatePackage, TemplatePart
from carbotek.xlsx.prepare import prepare_workbook


def prepare(content: bytes) -> bytes:
    name = "xl/workbook.xml"
    package = TemplatePackage.create({name: TemplatePart(name, content, True)}, (name,), "xlsx")
    prepared, _ = prepare_workbook(package)
    return prepared.parts[name].content


def test_utf8_xml_keeps_original_declaration_and_empty_element_spelling() -> None:
    content = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook><sheets></sheets></workbook>'
    assert prepare(content) == content


def test_non_utf8_xml_is_transcoded_before_marker_preparation() -> None:
    content = '<?xml version="1.0" encoding="ISO-8859-1"?><workbook><name>caf\u00e9</name></workbook>'.encode("latin-1")
    assert "caf\u00e9" in prepare(content).decode("utf-8")
    assert b"ISO-8859-1" not in prepare(content)


def test_unchanged_utf8_xml_still_rejects_dtd() -> None:
    with pytest.raises(CorruptedTemplateError):
        prepare(b'<!DOCTYPE workbook [<!ENTITY x "bad">]><workbook/>')


@pytest.mark.parametrize("uppercase_removed_parts", [False, True])
def test_render_cleans_up_references_to_removed_workbook_parts(uppercase_removed_parts: bool) -> None:
    content_types = "http://schemas.openxmlformats.org/package/2006/content-types"
    sheet_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    rel_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    package_rel = "http://schemas.openxmlformats.org/package/2006/relationships"
    mime = "application/vnd.openxmlformats-officedocument.spreadsheetml."
    worksheet = "xl/worksheets/sheet1.xml"
    overrides = {
        "xl/workbook.xml": mime + "sheet.main+xml",
        worksheet: mime + "worksheet+xml",
        "xl/sharedStrings.xml": mime + "sharedStrings+xml",
        "xl/calcChain.xml": mime + "calcChain+xml",
    }
    defaults = {"xml": "application/xml", "rels": "application/vnd.openxmlformats-package.relationships+xml"}
    manifest = "".join(f"<ct:Default Extension='{ext}' ContentType='{kind}'/>" for ext, kind in defaults.items())
    for part, kind in overrides.items():
        if uppercase_removed_parts and part in {"xl/sharedStrings.xml", "xl/calcChain.xml"}:
            part = part.upper()
        manifest += f"<ct:Override ContentType='{kind}' PartName='/{part}'/>"
    parts = {
        "[Content_Types].xml": f"<ct:Types xmlns:ct='{content_types}'>{manifest}</ct:Types>",
        "_rels/.rels": (
            f'<Relationships xmlns="{package_rel}">'
            f'<Relationship Id="r1" Target="xl/workbook.xml" Type="{rel_ns}/officeDocument"/></Relationships>'
        ),
        "xl/workbook.xml": (
            f'<workbook xmlns="{sheet_ns}" xmlns:r="{rel_ns}"><sheets>'
            '<sheet name="Sheet1" sheetId="1" r:id="r1"/></sheets></workbook>'
        ),
        "xl/_rels/workbook.xml.rels": (
            f"<p:Relationships xmlns:p='{package_rel}'>"
            f"<p:Relationship Id='r1' Target='worksheets/sheet1.xml' Type='{rel_ns}/worksheet'/>"
            f"<p:Relationship Id='r2' Target='/XL/SHAREDSTRINGS.XML' Type='{rel_ns}/sharedStrings'/>"
            f"<p:Relationship Id='r3' Target='../xl/calcChain.xml' Type='{rel_ns}/calcChain'/>"
            f"<p:Relationship Id='r4' Target='sharedStrings.xml' TargetMode='External' Type='{rel_ns}/hyperlink'/>"
            "</p:Relationships>"
        ),
        worksheet: (
            f'<worksheet xmlns="{sheet_ns}"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c>'
            '<c r="B1"><f>1+1</f><v>2</v></c></row></sheetData></worksheet>'
        ),
        "xl/sharedStrings.xml": f'<sst xmlns="{sheet_ns}"><si><t>{{d.name}}</t></si></sst>',
        "xl/calcChain.xml": f'<calcChain xmlns="{sheet_ns}"><c r="B1" i="1"/></calcChain>',
    }
    template = BytesIO()
    with ZipFile(template, "w") as archive:
        for name, content in parts.items():
            archive.writestr(name, content)

    rendered = Engine().render(template.getvalue(), {"name": "Alice"})
    with ZipFile(BytesIO(rendered)) as archive:
        assert {"xl/sharedStrings.xml", "xl/calcChain.xml"}.isdisjoint(archive.namelist())
        rels = etree.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        assert {node.get("Id"): node.get("Target") for node in rels.findall(f"{{{package_rel}}}Relationship")} == {
            "r1": "worksheets/sheet1.xml",
            "r4": "sharedStrings.xml",
        }
        types = etree.fromstring(archive.read("[Content_Types].xml"))
        assert {
            node.get("PartName"): node.get("ContentType") for node in types.findall(f"{{{content_types}}}Override")
        } == {
            "/xl/workbook.xml": overrides["xl/workbook.xml"],
            f"/{worksheet}": overrides[worksheet],
        }
        assert {
            node.get("Extension"): node.get("ContentType") for node in types.findall(f"{{{content_types}}}Default")
        } == defaults
        sheet = etree.fromstring(archive.read(worksheet))
        assert sheet.findtext(f".//{{{sheet_ns}}}t") == "Alice"
