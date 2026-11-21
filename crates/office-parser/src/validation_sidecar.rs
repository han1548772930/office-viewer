use std::collections::{HashMap, VecDeque};
use std::io::Read;

use quick_xml::events::{BytesStart, Event};
use quick_xml::{Reader, XmlVersion};
use zip::ZipArchive;

type RangeKey = (u32, u16, u32, u16);

#[derive(Debug, Clone, Default)]
pub(crate) struct ValidationSupplement {
    pub(crate) show_drop_down: Option<bool>,
    pub(crate) error_style: Option<String>,
}

#[derive(Debug, Default)]
pub(crate) struct ValidationSidecar {
    enabled: bool,
    by_sheet: HashMap<String, HashMap<RangeKey, VecDeque<ValidationSupplement>>>,
    pub(crate) warnings: Vec<String>,
}

impl ValidationSidecar {
    pub(crate) fn read(bytes: &[u8]) -> Self {
        match read_sidecar(bytes) {
            Ok(sidecar) => sidecar,
            Err(warning) => Self {
                warnings: vec![warning],
                ..Self::default()
            },
        }
    }

    pub(crate) fn take(
        &mut self,
        sheet_name: &str,
        range: RangeKey,
    ) -> Option<ValidationSupplement> {
        let supplement = self
            .by_sheet
            .get_mut(sheet_name)
            .and_then(|ranges| ranges.get_mut(&range))
            .and_then(VecDeque::pop_front);
        if self.enabled && supplement.is_none() {
            self.warnings.push(format!(
                "Could not match imported data validation range {range:?} to its XLSX metadata on worksheet {sheet_name}"
            ));
        }
        supplement
    }

    pub(crate) fn warn_unmatched_rules(&mut self) {
        let unmatched = self
            .by_sheet
            .values()
            .flat_map(HashMap::values)
            .map(VecDeque::len)
            .sum::<usize>();
        if unmatched > 0 {
            self.warnings.push(format!(
                "rxls did not retain {unmatched} data validation range(s), usually due to a missing formula or malformed rule"
            ));
        }
    }
}

fn read_sidecar(bytes: &[u8]) -> Result<ValidationSidecar, String> {
    let mut archive = ZipArchive::new(std::io::Cursor::new(bytes))
        .map_err(|_| "Could not read XLSX validation metadata: invalid ZIP package".to_owned())?;
    let root_rels = read_xml_part(&mut archive, "_rels/.rels")?.ok_or_else(|| {
        "Could not read XLSX validation metadata: missing package relationships".to_owned()
    })?;
    let workbook_path = relationships(&root_rels)
        .into_iter()
        .find_map(|relationship| {
            (relationship.kind == "officeDocument" && !relationship.external)
                .then(|| resolve_target("", &relationship.target))
                .flatten()
        })
        .ok_or_else(|| {
            "Could not locate the XLSX workbook part for validation metadata".to_owned()
        })?;
    let workbook_xml = read_xml_part(&mut archive, &workbook_path)?
        .ok_or_else(|| format!("Missing XLSX workbook part {workbook_path}"))?;
    let rels_path = relationship_part_path(&workbook_path);
    let workbook_rels_xml = read_xml_part(&mut archive, &rels_path)?
        .ok_or_else(|| format!("Missing XLSX workbook relationships {rels_path}"))?;
    let sheet_relationships = relationships(&workbook_rels_xml)
        .into_iter()
        .filter(|relationship| !relationship.external)
        .filter_map(|relationship| {
            let path = resolve_target(&workbook_path, &relationship.target)?;
            Some((relationship.id, (relationship.kind, path)))
        })
        .collect::<HashMap<_, _>>();
    let sheet_refs = workbook_sheets(&workbook_xml);
    let mut sidecar = ValidationSidecar {
        enabled: true,
        ..ValidationSidecar::default()
    };
    for (sheet_name, relationship_id) in sheet_refs {
        let Some((kind, sheet_path)) = sheet_relationships
            .get(&relationship_id)
            .map(|(kind, path)| (kind, path))
        else {
            sidecar.warnings.push(format!(
                "Could not resolve validation metadata worksheet relationship for {sheet_name}"
            ));
            continue;
        };
        if kind != "worksheet" {
            continue;
        }
        let Some(sheet_xml) = read_xml_part(&mut archive, sheet_path)? else {
            sidecar.warnings.push(format!(
                "Missing worksheet XML {sheet_path} while reading validation metadata"
            ));
            continue;
        };
        parse_worksheet_validations(&sheet_name, &sheet_xml, &mut sidecar);
    }
    Ok(sidecar)
}

#[derive(Debug)]
struct Relationship {
    id: String,
    target: String,
    kind: String,
    external: bool,
}

fn relationships(xml: &str) -> Vec<Relationship> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut relationships = Vec::new();
    loop {
        match reader.read_event() {
            Ok(Event::Start(element) | Event::Empty(element))
                if local_name(element.name().as_ref()) == b"Relationship" =>
            {
                let attrs = attributes(&reader, &element);
                let id = attrs.get("Id").cloned();
                let target = attrs.get("Target").cloned();
                let kind = attrs
                    .get("Type")
                    .and_then(|value| value.rsplit('/').next())
                    .map(str::to_owned);
                let external = attrs
                    .get("TargetMode")
                    .is_some_and(|value| value == "External");
                if let (Some(id), Some(target), Some(kind)) = (id, target, kind) {
                    relationships.push(Relationship {
                        id,
                        target,
                        kind,
                        external,
                    });
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    relationships
}

fn workbook_sheets(xml: &str) -> Vec<(String, String)> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut sheets = Vec::new();
    loop {
        match reader.read_event() {
            Ok(Event::Start(element) | Event::Empty(element))
                if local_name(element.name().as_ref()) == b"sheet" =>
            {
                let attrs = attributes(&reader, &element);
                if let (Some(name), Some(relationship_id)) =
                    (attrs.get("name").cloned(), attrs.get("id").cloned())
                {
                    sheets.push((name, relationship_id));
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    sheets
}

fn parse_worksheet_validations(sheet_name: &str, xml: &str, sidecar: &mut ValidationSidecar) {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    loop {
        match reader.read_event() {
            Ok(Event::Start(element) | Event::Empty(element))
                if local_name(element.name().as_ref()) == b"dataValidation" =>
            {
                let attrs = attributes(&reader, &element);
                let kind = attrs.get("type").map(String::as_str).unwrap_or_default();
                if !matches!(
                    kind,
                    "list" | "whole" | "decimal" | "date" | "time" | "textLength" | "custom"
                ) {
                    sidecar.warnings.push(format!(
                        "Skipped unsupported or missing data validation type {kind:?} on worksheet {sheet_name}"
                    ));
                    continue;
                }
                if let Some(operator) = attrs.get("operator")
                    && !matches!(
                        operator.as_str(),
                        "between"
                            | "notBetween"
                            | "equal"
                            | "notEqual"
                            | "greaterThan"
                            | "lessThan"
                            | "greaterThanOrEqual"
                            | "lessThanOrEqual"
                    )
                {
                    sidecar.warnings.push(format!(
                        "Unsupported data validation operator {operator:?} on worksheet {sheet_name}"
                    ));
                }
                for attribute in [
                    "allowBlank",
                    "showDropDown",
                    "showInputMessage",
                    "showErrorMessage",
                ] {
                    if attrs
                        .get(attribute)
                        .is_some_and(|value| parse_ooxml_bool(value).is_none())
                    {
                        sidecar.warnings.push(format!(
                            "Invalid {attribute} boolean {value:?} on worksheet {sheet_name}",
                            value = attrs.get(attribute).expect("attribute was present")
                        ));
                    }
                }
                let Some(sqref) = attrs.get("sqref") else {
                    sidecar.warnings.push(format!(
                        "Skipped data validation without sqref on worksheet {sheet_name}"
                    ));
                    continue;
                };
                if sqref.trim().is_empty() {
                    sidecar.warnings.push(format!(
                        "Skipped data validation with empty sqref on worksheet {sheet_name}"
                    ));
                    continue;
                }
                let ranges = sqref.split_whitespace();
                let supplement = ValidationSupplement {
                    // OOXML showDropDown=true suppresses the in-cell arrow.
                    show_drop_down: attrs
                        .get("showDropDown")
                        .and_then(|value| parse_ooxml_bool(value).map(|hidden| !hidden)),
                    error_style: attrs.get("errorStyle").map(|value| match value.as_str() {
                        "information" => "info".to_owned(),
                        other => other.to_owned(),
                    }),
                };
                if supplement
                    .error_style
                    .as_deref()
                    .is_some_and(|style| !matches!(style, "stop" | "warning" | "info"))
                {
                    sidecar.warnings.push(format!(
                        "Unsupported data validation error style {:?} on worksheet {sheet_name}",
                        supplement.error_style
                    ));
                }
                for range_text in ranges {
                    match parse_range(range_text) {
                        Some(range) => sidecar
                            .by_sheet
                            .entry(sheet_name.to_owned())
                            .or_default()
                            .entry(range)
                            .or_default()
                            .push_back(supplement.clone()),
                        None => sidecar.warnings.push(format!(
                            "Skipped malformed data validation range {range_text:?} on worksheet {sheet_name}"
                        )),
                    }
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
}

fn parse_ooxml_bool(value: &str) -> Option<bool> {
    match value {
        "1" | "true" => Some(true),
        "0" | "false" => Some(false),
        _ => None,
    }
}

fn attributes(reader: &Reader<&[u8]>, element: &BytesStart<'_>) -> HashMap<String, String> {
    element
        .attributes()
        .with_checks(false)
        .filter_map(Result::ok)
        .filter_map(|attribute| {
            let name = local_name(attribute.key.as_ref());
            let name = std::str::from_utf8(name).ok()?.to_owned();
            let value = attribute
                .decoded_and_normalized_value_with(
                    XmlVersion::Implicit1_0,
                    reader.decoder(),
                    1,
                    quick_xml::escape::resolve_xml_entity,
                )
                .ok()?
                .into_owned();
            Some((name, value))
        })
        .collect()
}

fn read_xml_part(
    archive: &mut ZipArchive<std::io::Cursor<&[u8]>>,
    path: &str,
) -> Result<Option<String>, String> {
    const MAX_METADATA_XML_BYTES: u64 = 16 * 1024 * 1024;
    let mut entry = match archive.by_name(path) {
        Ok(entry) => entry,
        Err(zip::result::ZipError::FileNotFound) => return Ok(None),
        Err(error) => return Err(format!("Could not read XLSX package part {path}: {error}")),
    };
    if entry.size() > MAX_METADATA_XML_BYTES {
        return Err(format!(
            "XLSX metadata XML part {path} exceeds the parser limit"
        ));
    }
    let mut xml = String::with_capacity(entry.size() as usize);
    entry
        .read_to_string(&mut xml)
        .map_err(|error| format!("Could not decode XLSX package part {path}: {error}"))?;
    Ok(Some(xml))
}

fn relationship_part_path(source_part: &str) -> String {
    match source_part.rsplit_once('/') {
        Some((directory, file)) => format!("{directory}/_rels/{file}.rels"),
        None => format!("_rels/{source_part}.rels"),
    }
}

fn resolve_target(source_part: &str, target: &str) -> Option<String> {
    if target.contains(':') || target.starts_with("//") {
        return None;
    }
    let mut parts = if target.starts_with('/') {
        Vec::new()
    } else {
        source_part
            .rsplit_once('/')
            .map(|(directory, _)| directory.split('/').collect())
            .unwrap_or_default()
    };
    for part in target.trim_start_matches('/').split('/') {
        match part {
            "" | "." => {}
            ".." => {
                parts.pop()?;
            }
            segment => parts.push(segment),
        }
    }
    (!parts.is_empty()).then(|| parts.join("/"))
}

fn parse_range(value: &str) -> Option<RangeKey> {
    let (first, last) = match value.split_once(':') {
        Some((first, last)) if !last.contains(':') => (parse_cell(first)?, parse_cell(last)?),
        Some(_) => return None,
        None => {
            let cell = parse_cell(value)?;
            (cell, cell)
        }
    };
    (first.0 <= last.0 && first.1 <= last.1).then_some((first.0, first.1, last.0, last.1))
}

fn parse_cell(value: &str) -> Option<(u32, u16)> {
    let value = value.replace('$', "");
    let split = value.find(|character: char| character.is_ascii_digit())?;
    let (column, row) = value.split_at(split);
    if column.is_empty() || row.is_empty() || !column.bytes().all(|byte| byte.is_ascii_alphabetic())
    {
        return None;
    }
    let column = column.bytes().try_fold(0_u32, |index, byte| {
        index
            .checked_mul(26)?
            .checked_add(u32::from(byte.to_ascii_uppercase() - b'A' + 1))
    })?;
    let row = row.parse::<u32>().ok()?;
    if column == 0 || column > 16_384 || row == 0 || row > 1_048_576 {
        return None;
    }
    Some((row - 1, u16::try_from(column - 1).ok()?))
}

fn local_name(qualified_name: &[u8]) -> &[u8] {
    qualified_name
        .rsplit(|byte| *byte == b':')
        .next()
        .unwrap_or(qualified_name)
}

#[cfg(test)]
mod tests {
    use super::{ValidationSidecar, parse_range, parse_worksheet_validations};

    #[test]
    fn parses_absolute_and_multi_cell_validation_ranges() {
        assert_eq!(parse_range("$A$2:C4"), Some((1, 0, 3, 2)));
        assert_eq!(
            parse_range("XFD1048576"),
            Some((1_048_575, 16_383, 1_048_575, 16_383))
        );
        assert_eq!(parse_range("C4:A2"), None);
        assert_eq!(parse_range("A0"), None);
    }

    #[test]
    fn warns_for_unsupported_types_and_malformed_ranges() {
        let xml = r#"<worksheet><dataValidations>
            <dataValidation type="extensionRule" sqref="A1"/>
            <dataValidation type="whole" sqref="A0"/>
            <dataValidation type="whole" sqref="A1" operator="invalid" showDropDown="yes"/>
            <dataValidation type="whole"/>
        </dataValidations></worksheet>"#;
        let mut sidecar = ValidationSidecar::default();

        parse_worksheet_validations("Input", xml, &mut sidecar);

        assert_eq!(sidecar.warnings.len(), 5);
        assert!(sidecar.warnings[0].contains("unsupported or missing data validation type"));
        assert!(sidecar.warnings[1].contains("malformed data validation range"));
        assert!(sidecar.warnings[2].contains("Unsupported data validation operator"));
        assert!(sidecar.warnings[3].contains("Invalid showDropDown boolean"));
        assert!(sidecar.warnings[4].contains("without sqref"));
    }
}
