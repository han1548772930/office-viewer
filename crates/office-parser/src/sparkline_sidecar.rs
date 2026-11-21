use std::collections::HashMap;
use std::io::Read;

use quick_xml::events::{BytesStart, Event};
use quick_xml::{Reader, XmlVersion};
use zip::ZipArchive;

use crate::model::SparklinePayload;

#[derive(Debug, Default)]
pub(crate) struct SparklineSidecar {
    by_sheet: HashMap<String, HashMap<String, SparklinePayload>>,
    pub(crate) warnings: Vec<String>,
}

impl SparklineSidecar {
    pub(crate) fn read(bytes: &[u8]) -> Self {
        match read_sidecar(bytes) {
            Ok(sidecar) => sidecar,
            Err(warning) => Self { warnings: vec![warning], ..Self::default() },
        }
    }

    pub(crate) fn take(&mut self, sheet: &str, target: &str) -> Option<SparklinePayload> {
        self.by_sheet.get_mut(sheet)?.remove(&target.to_ascii_uppercase())
    }

    pub(crate) fn take_sheet(&mut self, sheet: &str) -> HashMap<String, SparklinePayload> {
        self.by_sheet.remove(sheet).unwrap_or_default()
    }
}

#[derive(Debug)]
struct Relationship { id: String, target: String, kind: String, external: bool }

#[derive(Debug, Default)]
struct Group {
    kind: String,
    empty_cells: String,
    series_color: Option<String>,
    negative_color: Option<String>,
    high_color: Option<String>,
    low_color: Option<String>,
    axis: String,
    min: Option<f64>,
    max: Option<f64>,
}

#[derive(Debug, Default)]
struct Sparkline { formula: String, target: String }

fn read_sidecar(bytes: &[u8]) -> Result<SparklineSidecar, String> {
    let mut archive = ZipArchive::new(std::io::Cursor::new(bytes))
        .map_err(|_| "Could not read XLSX sparkline metadata: invalid ZIP package".to_owned())?;
    let root_rels = read_xml_part(&mut archive, "_rels/.rels")?.ok_or_else(|| "Could not read XLSX sparkline metadata: missing package relationships".to_owned())?;
    let workbook_path = relationships(&root_rels).into_iter().find_map(|rel| {
        (rel.kind == "officeDocument" && !rel.external).then(|| resolve_target("", &rel.target)).flatten()
    }).ok_or_else(|| "Could not locate XLSX workbook part for sparkline metadata".to_owned())?;
    let workbook = read_xml_part(&mut archive, &workbook_path)?.ok_or_else(|| format!("Missing XLSX workbook part {workbook_path}"))?;
    let rels_path = relationship_part_path(&workbook_path);
    let rels_xml = read_xml_part(&mut archive, &rels_path)?.ok_or_else(|| format!("Missing XLSX workbook relationships {rels_path}"))?;
    let sheet_relationships = relationships(&rels_xml).into_iter().filter(|rel| !rel.external).filter_map(|rel| {
        resolve_target(&workbook_path, &rel.target).map(|path| (rel.id, (rel.kind, path)))
    }).collect::<HashMap<_, _>>();
    let mut sidecar = SparklineSidecar::default();
    for (sheet_name, relationship_id) in workbook_sheets(&workbook) {
        let Some((kind, path)) = sheet_relationships.get(&relationship_id) else {
            sidecar.warnings.push(format!("Could not resolve sparkline worksheet relationship for {sheet_name}"));
            continue;
        };
        if kind != "worksheet" { continue }
        let Some(xml) = read_xml_part(&mut archive, path)? else {
            sidecar.warnings.push(format!("Missing worksheet XML {path} while reading sparkline metadata"));
            continue;
        };
        let controls = parse_sparklines(&xml, &mut sidecar.warnings);
        if !controls.is_empty() { sidecar.by_sheet.insert(sheet_name, controls); }
    }
    Ok(sidecar)
}

fn parse_sparklines(xml: &str, warnings: &mut Vec<String>) -> HashMap<String, SparklinePayload> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut in_group = false;
    let mut in_sparkline = false;
    let mut group = Group::default();
    let mut sparkline = Sparkline::default();
    let mut current_value: Option<&'static [u8]> = None;
    let mut result = HashMap::new();
    loop {
        match reader.read_event() {
            Ok(Event::Start(element) | Event::Empty(element)) => {
                let qualified_name = element.name();
                let name = local_name(qualified_name.as_ref());
                let attrs = attributes(&reader, &element);
                if name == b"sparklineGroup" {
                    in_group = true;
                    group = Group {
                        kind: attrs.get("type").cloned().unwrap_or_else(|| "line".into()),
                        empty_cells: attrs.get("displayEmptyCellsAs").cloned().unwrap_or_else(|| "gap".into()),
                        axis: if attrs.get("minAxisType").is_some_and(|value| value == "custom") || attrs.get("maxAxisType").is_some_and(|value| value == "custom") { "custom".into() } else { "individual".into() },
                        min: attrs.get("manualMin").and_then(|value| value.parse().ok()),
                        max: attrs.get("manualMax").and_then(|value| value.parse().ok()),
                        ..Group::default()
                    };
                } else if in_group && name == b"sparkline" {
                    in_sparkline = true;
                    sparkline = Sparkline::default();
                } else if in_group && name.starts_with(b"color") {
                    let color = attrs.get("rgb").and_then(|value| normalize_color(value));
                    match name {
                        b"colorSeries" => group.series_color = color.clone(),
                        b"colorNegative" => group.negative_color = color.clone(),
                        b"colorHigh" => group.high_color = color.clone(),
                        b"colorLow" => group.low_color = color.clone(),
                        _ => {}
                    }
                }
                if in_sparkline && !element.is_empty() {
                    current_value = match name { b"f" => Some(b"f"), b"sqref" => Some(b"sqref"), _ => None };
                }
            }
            Ok(Event::Text(text)) if in_sparkline => {
                let decoded = text.decode().unwrap_or_default();
                let value = quick_xml::escape::unescape(&decoded).unwrap_or_else(|_| decoded.clone()).into_owned();
                match current_value {
                    Some(b"f") => sparkline.formula.push_str(&value),
                    Some(b"sqref") => sparkline.target.push_str(&value),
                    _ => {}
                }
            }
            Ok(Event::End(element)) => {
                match local_name(element.name().as_ref()) {
                    b"f" if in_sparkline => { /* populated by the text handler below */ }
                    b"sparkline" if in_group => {
                        if let Some(target) = normalize_cell_reference(&sparkline.target) {
                            if let Some(source) = normalize_formula(&sparkline.formula) {
                                let kind = match group.kind.as_str() { "column" => "column", "stacked" => "winloss", _ => "line" };
                                let empty_cells = match group.empty_cells.as_str() { "zero" => "zero", "span" => "connect", _ => "gap" };
                                result.insert(target, SparklinePayload {
                                    sparkline_type: kind.into(), source_range: source,
                                    color: group.series_color.clone().unwrap_or_else(|| "#4472C4".into()),
                                    negative_color: group.negative_color.clone(), high_point_color: group.high_color.clone(),
                                    low_point_color: group.low_color.clone(), empty_cells: empty_cells.into(), axis: group.axis.clone(),
                                    min: group.min, max: group.max,
                                });
                            } else { warnings.push("Skipped XLSX sparkline with invalid source formula".into()); }
                        } else { warnings.push("Skipped XLSX sparkline with invalid target cell".into()); }
                        in_sparkline = false;
                    }
                    b"sparklineGroup" => in_group = false,
                    _ => {}
                }
                if matches!(local_name(element.name().as_ref()), b"f" | b"sqref") { current_value = None; }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    result
}

fn normalize_formula(value: &str) -> Option<String> {
    let reference = value.rsplit_once('!').map_or(value, |(_, reference)| reference).replace('$', "");
    let mut parts = reference.split(':');
    let first = normalize_cell_reference(parts.next()?)?;
    let last = normalize_cell_reference(parts.next().unwrap_or(&first))?;
    if parts.next().is_some() { return None }
    Some(if first == last { first } else { format!("{first}:{last}") })
}

fn normalize_cell_reference(value: &str) -> Option<String> {
    let value = value.trim().replace('$', "");
    let split = value.find(|ch: char| ch.is_ascii_digit())?;
    let (letters, digits) = value.split_at(split);
    if letters.is_empty() || !letters.bytes().all(|ch| ch.is_ascii_alphabetic()) { return None }
    let row = digits.parse::<u32>().ok()?;
    if row == 0 || row > 1_048_576 { return None }
    Some(format!("{}{}", letters.to_ascii_uppercase(), row))
}

fn normalize_color(value: &str) -> Option<String> {
    let hex = value.strip_prefix('#').unwrap_or(value);
    if hex.len() != 6 && hex.len() != 8 || !hex.bytes().all(|ch| ch.is_ascii_hexdigit()) { return None }
    Some(format!("#{}", if hex.len() == 8 { &hex[2..] } else { hex }.to_ascii_uppercase()))
}

fn relationships(xml: &str) -> Vec<Relationship> {
    let mut reader = Reader::from_str(xml); reader.config_mut().trim_text(true);
    let mut result = Vec::new();
    loop {
        match reader.read_event() {
            Ok(Event::Start(element) | Event::Empty(element)) if local_name(element.name().as_ref()) == b"Relationship" => {
                let attrs = attributes(&reader, &element);
                if let (Some(id), Some(target), Some(kind)) = (attrs.get("Id"), attrs.get("Target"), attrs.get("Type").and_then(|value| value.rsplit('/').next())) {
                    result.push(Relationship { id: id.clone(), target: target.clone(), kind: kind.into(), external: attrs.get("TargetMode").is_some_and(|value| value == "External") });
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    result
}

fn workbook_sheets(xml: &str) -> Vec<(String, String)> {
    let mut reader = Reader::from_str(xml); reader.config_mut().trim_text(true);
    let mut result = Vec::new();
    loop {
        match reader.read_event() {
            Ok(Event::Start(element) | Event::Empty(element)) if local_name(element.name().as_ref()) == b"sheet" => {
                let attrs = attributes(&reader, &element);
                if let (Some(name), Some(id)) = (attrs.get("name"), attrs.get("id")) { result.push((name.clone(), id.clone())); }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    result
}

fn attributes(reader: &Reader<&[u8]>, element: &BytesStart<'_>) -> HashMap<String, String> {
    element.attributes().with_checks(false).filter_map(Result::ok).filter_map(|attribute| {
        let name = std::str::from_utf8(local_name(attribute.key.as_ref())).ok()?.to_owned();
        let value = attribute.decoded_and_normalized_value_with(XmlVersion::Implicit1_0, reader.decoder(), 1, quick_xml::escape::resolve_xml_entity).ok()?.into_owned();
        Some((name, value))
    }).collect()
}

fn read_xml_part(archive: &mut ZipArchive<std::io::Cursor<&[u8]>>, path: &str) -> Result<Option<String>, String> {
    let mut entry = match archive.by_name(path) { Ok(entry) => entry, Err(zip::result::ZipError::FileNotFound) => return Ok(None), Err(error) => return Err(format!("Could not read XLSX package part {path}: {error}")) };
    if entry.size() > 16 * 1024 * 1024 { return Err(format!("XLSX metadata XML part {path} exceeds the parser limit")) }
    let mut value = String::with_capacity(entry.size() as usize);
    entry.read_to_string(&mut value).map_err(|error| format!("Could not decode XLSX package part {path}: {error}"))?;
    Ok(Some(value))
}

fn relationship_part_path(source: &str) -> String {
    match source.rsplit_once('/') { Some((directory, file)) => format!("{directory}/_rels/{file}.rels"), None => format!("_rels/{source}.rels") }
}

fn resolve_target(source: &str, target: &str) -> Option<String> {
    if target.contains(':') || target.starts_with("//") { return None }
    let mut parts = if target.starts_with('/') { Vec::new() } else { source.rsplit_once('/').map(|(directory, _)| directory.split('/').collect()).unwrap_or_default() };
    for part in target.trim_start_matches('/').split('/') { match part { "" | "." => {}, ".." => { parts.pop()?; }, value => parts.push(value) } }
    (!parts.is_empty()).then(|| parts.join("/"))
}

fn local_name(value: &[u8]) -> &[u8] { value.rsplit(|byte| *byte == b':').next().unwrap_or(value) }
