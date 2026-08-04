# -*- coding: utf-8 -*-

from hashlib import sha256
from io import BytesIO
import json
from pathlib import PurePosixPath
from zipfile import ZipFile

from docx import Document
from lxml import etree


NAMESPACES = {
    "o": "urn:schemas-microsoft-com:office:office",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "v": "urn:schemas-microsoft-com:vml",
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
}
PACKAGE_RELATIONSHIP = (
    "http://schemas.openxmlformats.org/officeDocument/2006/"
    "relationships/package"
)
W14_ANCHOR_ID = "{%s}anchorId" % NAMESPACES["w14"]


def assert_file_identity(path, expected_size, expected_sha256):
    data = path.read_bytes()
    assert len(data) == expected_size, (
        "unexpected fixture size: %s" % path.name
    )
    assert sha256(data).hexdigest() == expected_sha256, (
        "unexpected fixture SHA-256: %s" % path.name
    )


def _inspect_ooxml_package(data, required_main_part=None):
    with ZipFile(BytesIO(data)) as package:
        assert package.testzip() is None, (
            "embedded OOXML package has corrupt data"
        )
        names = set(package.namelist())
        assert "[Content_Types].xml" in names, (
            "embedded target has no content types"
        )
        etree.fromstring(package.read("[Content_Types].xml"))

        if required_main_part is None:
            main_parts = sorted(
                names
                & {
                    "ppt/presentation.xml",
                    "word/document.xml",
                    "xl/workbook.xml",
                }
            )
            assert len(main_parts) == 1, (
                "embedded package has no unique OOXML main part"
            )
            main_part = main_parts[0]
        else:
            assert required_main_part in names, (
                "embedded package lacks %s" % required_main_part
            )
            main_part = required_main_part

        main_xml = etree.fromstring(package.read(main_part))
        sheet_names = (
            main_xml.xpath(".//s:sheet/@name", namespaces=NAMESPACES)
            if main_part == "xl/workbook.xml"
            else []
        )
    return main_part, sheet_names


def assert_xlsx_fixture(path, expected_size, expected_sha256, expected_sheet):
    assert_file_identity(path, expected_size, expected_sha256)
    main_part, sheet_names = _inspect_ooxml_package(
        path.read_bytes(), required_main_part="xl/workbook.xml"
    )
    assert main_part == "xl/workbook.xml"
    assert sheet_names == [expected_sheet], (
        "unexpected workbook sheets in %s: %r" % (path.name, sheet_names)
    )


def add_ole_paragraph_anchor_ids(source, target):
    anchor_count = 0
    with ZipFile(source) as source_archive, ZipFile(
        target, "w"
    ) as target_archive:
        for item in source_archive.infolist():
            data = source_archive.read(item.filename)
            if item.filename == "word/document.xml":
                document_xml = etree.fromstring(data)
                paragraphs = document_xml.xpath(
                    ".//w:p[.//o:OLEObject]", namespaces=NAMESPACES
                )
                anchor_count = len(paragraphs)
                for index, paragraph in enumerate(paragraphs, start=1):
                    paragraph.set(W14_ANCHOR_ID, "%08X" % index)
                data = etree.tostring(
                    document_xml,
                    xml_declaration=True,
                    encoding="UTF-8",
                    standalone=True,
                )
            target_archive.writestr(item, data)
    return anchor_count


def add_non_ole_vml_shape(source, target, shape_id):
    with ZipFile(source) as source_archive, ZipFile(
        target, "w"
    ) as target_archive:
        for item in source_archive.infolist():
            data = source_archive.read(item.filename)
            if item.filename == "word/document.xml":
                document_xml = etree.fromstring(data)
                ole_objects = document_xml.xpath(
                    ".//o:OLEObject", namespaces=NAMESPACES
                )
                for index, ole_object in enumerate(
                    ole_objects, start=2048
                ):
                    old_shape_id = ole_object.get("ShapeID")
                    matching_shapes = ole_object.getparent().xpath(
                        "./v:shape[@id=$shape_id]",
                        namespaces=NAMESPACES,
                        shape_id=old_shape_id,
                    )
                    assert len(matching_shapes) == 1
                    new_shape_id = "_x0000_i%d" % index
                    matching_shapes[0].set("id", new_shape_id)
                    ole_object.set("ShapeID", new_shape_id)
                body = document_xml.xpath(
                    ".//w:body", namespaces=NAMESPACES
                )[0]
                paragraph = etree.Element("{%s}p" % NAMESPACES["w"])
                run = etree.SubElement(
                    paragraph, "{%s}r" % NAMESPACES["w"]
                )
                pict = etree.SubElement(
                    run, "{%s}pict" % NAMESPACES["w"]
                )
                shape = etree.SubElement(
                    pict, "{%s}shape" % NAMESPACES["v"]
                )
                shape.set("id", shape_id)
                section = body.find("{%s}sectPr" % NAMESPACES["w"])
                if section is None:
                    body.append(paragraph)
                else:
                    body.insert(body.index(section), paragraph)
                data = etree.tostring(
                    document_xml,
                    xml_declaration=True,
                    encoding="UTF-8",
                    standalone=True,
                )
            target_archive.writestr(item, data)


def _assert_unique(values, label, normalizer=None):
    normalized = (
        [normalizer(value) for value in values]
        if normalizer
        else values
    )
    assert len(set(normalized)) == len(normalized), "duplicate %s: %r" % (
        label,
        values,
    )


def _embedding_target(relationship):
    target = relationship.get("Target")
    assert target, "OLE package relationship has no target"
    relative = PurePosixPath(target)
    assert not relative.is_absolute(), (
        "external OLE package target: %s" % target
    )
    assert ".." not in relative.parts, (
        "traversing OLE package target: %s" % target
    )
    package_path = PurePosixPath("word") / relative
    assert package_path.parts[:2] == ("word", "embeddings"), (
        "OLE target is outside word/embeddings: %s" % target
    )
    return package_path.as_posix()


def assert_ole_integrity(
    result,
    expected_count,
    expected_anchor_count,
    required_embedding_member=None,
    expected_embedding_sha256=None,
    expected_workbook_sheets=None,
):
    Document(result)
    with ZipFile(result) as archive:
        assert archive.testzip() is None, "result DOCX has corrupt ZIP data"
        document_xml = etree.fromstring(archive.read("word/document.xml"))
        relationships_xml = etree.fromstring(
            archive.read("word/_rels/document.xml.rels")
        )
        archive_names = set(archive.namelist())
        relationships = {
            relationship.get("Id"): relationship
            for relationship in relationships_xml.xpath(
                ".//pr:Relationship", namespaces=NAMESPACES
            )
        }
        ole_objects = document_xml.xpath(
            ".//o:OLEObject", namespaces=NAMESPACES
        )
        assert len(ole_objects) == expected_count, (
            "expected %d OLE objects, found %d"
            % (expected_count, len(ole_objects))
        )

        shape_ids = []
        object_ids = []
        relationship_ids = []
        embedding_targets = []
        embedding_sha256 = []
        embedded_main_parts = []
        workbook_sheet_names = []
        for ole_object in ole_objects:
            shape_id = ole_object.get("ShapeID")
            object_id = ole_object.get("ObjectID")
            relationship_id = ole_object.get("{%s}id" % NAMESPACES["r"])
            assert shape_id, "OLE object has no ShapeID"
            assert object_id, "OLE object has no ObjectID"
            assert relationship_id, "OLE object has no relationship ID"

            matching_shapes = ole_object.getparent().xpath(
                "./v:shape[@id=$shape_id]",
                namespaces=NAMESPACES,
                shape_id=shape_id,
            )
            assert len(matching_shapes) == 1, (
                "OLE ShapeID does not resolve to one sibling shape: %s"
                % shape_id
            )

            relationship = relationships.get(relationship_id)
            assert relationship is not None, (
                "missing OLE relationship: %s" % relationship_id
            )
            assert relationship.get("TargetMode") is None, (
                "OLE package relationship must be internal: %s"
                % relationship_id
            )
            assert relationship.get("Type") == PACKAGE_RELATIONSHIP, (
                "unexpected OLE relationship type: %s" % relationship_id
            )
            embedding_target = _embedding_target(relationship)
            assert embedding_target in archive_names, (
                "missing embedded package part: %s" % embedding_target
            )

            embedded_data = archive.read(embedding_target)
            main_part, sheet_names = _inspect_ooxml_package(
                embedded_data, required_embedding_member
            )
            shape_ids.append(shape_id)
            object_ids.append(object_id)
            relationship_ids.append(relationship_id)
            embedding_targets.append(embedding_target)
            embedding_sha256.append(sha256(embedded_data).hexdigest())
            embedded_main_parts.append(main_part)
            workbook_sheet_names.extend(sheet_names)

    assert len(shape_ids) == expected_count
    assert len(object_ids) == expected_count
    assert len(relationship_ids) == expected_count
    assert len(embedding_targets) == expected_count
    _assert_unique(shape_ids, "OLE ShapeID values")
    _assert_unique(object_ids, "OLE ObjectID values")
    _assert_unique(relationship_ids, "OLE relationship IDs")
    _assert_unique(embedding_targets, "OLE embedding targets")

    all_shape_ids = document_xml.xpath(".//v:shape/@id", namespaces=NAMESPACES)
    assert all(all_shape_ids), "empty VML shape ID"
    _assert_unique(all_shape_ids, "VML shape ID values")

    anchor_ids = document_xml.xpath(".//@w14:anchorId", namespaces=NAMESPACES)
    assert len(anchor_ids) == expected_anchor_count, (
        "expected %d w14:anchorId values, found %d"
        % (expected_anchor_count, len(anchor_ids))
    )
    assert all(anchor_ids), "empty w14:anchorId value"
    _assert_unique(
        anchor_ids,
        "w14:anchorId values after hexadecimal normalization",
        normalizer=lambda value: value.upper(),
    )

    if expected_embedding_sha256 is not None:
        assert embedding_sha256 == list(expected_embedding_sha256), (
            "embedded payload SHA-256 values changed: %r" % embedding_sha256
        )
    if expected_workbook_sheets is not None:
        assert workbook_sheet_names == list(expected_workbook_sheets), (
            "embedded workbook sheets changed: %r" % workbook_sheet_names
        )

    summary = {
        "anchor_count": len(anchor_ids),
        "embedded_main_parts": sorted(embedded_main_parts),
        "embedding_sha256": sorted(embedding_sha256),
        "embedding_target_count": len(embedding_targets),
        "object_id_count": len(object_ids),
        "ole_count": len(ole_objects),
        "relationship_id_count": len(relationship_ids),
        "shape_id_count": len(shape_ids),
        "workbook_sheet_names": sorted(workbook_sheet_names),
    }
    print(json.dumps(summary, sort_keys=True))
    return summary
