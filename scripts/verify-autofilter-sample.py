"""Independently inspect the XLSX sample emitted by office-worker-smoke.spec.ts.

Uses only Python's ZIP/XML readers. This checks the actual Office parts and ZIP
integrity; it does not claim that an Office application's refiltering is tested.
"""

import hashlib
import json
import sys
import xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import ZipFile


sample = Path(sys.argv[1])
namespace = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
with ZipFile(sample) as package:
    assert package.testzip() is None, "ZIP checksum validation failed"
    for entry in package.infolist():
        assert entry.date_time == (1980, 1, 1, 0, 0, 0), entry.filename
        if entry.filename.endswith((".xml", ".rels")):
            ET.fromstring(package.read(entry.filename))
    sheet = ET.fromstring(package.read("xl/worksheets/sheet1.xml"))
    auto_filter = sheet.find("s:autoFilter", namespace)
    assert auto_filter is not None and auto_filter.attrib["ref"] == "B3:D7"
    columns = auto_filter.findall("s:filterColumn", namespace)
    assert [column.attrib["colId"] for column in columns] == ["0", "2"]
    members = columns[0].findall("s:filters/s:filter", namespace)
    assert [member.attrib["val"] for member in members] == ["East"]
    conditions = columns[1].find("s:customFilters", namespace)
    assert conditions is not None and conditions.attrib["and"] == "1"
    assert [condition.attrib for condition in conditions] == [
        {"operator": "greaterThanOrEqual", "val": "100"},
        {"operator": "lessThanOrEqual", "val": "200"},
    ]
    hidden = [int(row.attrib["r"]) for row in sheet.findall("s:sheetData/s:row", namespace)
              if row.attrib.get("hidden") == "1"]
    assert hidden == [5, 6, 7], hidden
    properties = sheet.find("s:sheetPr", namespace)
    assert properties is not None and properties.attrib.get("filterMode") == "1"

print(json.dumps({
    "sample": str(sample),
    "sha256": hashlib.sha256(sample.read_bytes()).hexdigest(),
    "zip_and_xml_valid": True,
    "auto_filter_ref": auto_filter.attrib["ref"],
    "relative_column_ids": [column.attrib["colId"] for column in columns],
    "hidden_rows_one_based": hidden,
    "portable_zip_timestamps": True,
}, indent=2))
