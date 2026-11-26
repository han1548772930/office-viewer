//! Small OPC color resources for x14 sparklines, including relocated and Strict parts.

use std::collections::BTreeMap;
use std::io::{BufReader, Cursor, Read};
use zip::ZipArchive;
use office_io_common::conditional_format_xml::{self as xml, Element, XmlSelectionLimits};

const PACKAGE: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const STRICT_PACKAGE: &str = "http://purl.oclc.org/ooxml/package/relationships";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
const STRICT_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/";
const MAX_METADATA_BYTES: u64 = 16 * 1024 * 1024;

#[derive(Default)]
pub(super) struct ColorResources {
    pub(super) theme_xml: Option<String>,
    pub(super) theme: Option<Element>,
    pub(super) indexed_colors: Option<Vec<String>>,
}
struct Relationship { kind: String, target: String, external: bool }

pub(super) fn read_color_resources(archive: &mut ZipArchive<Cursor<&[u8]>>) -> Result<ColorResources, String> {
    let root = relationships(&read_xml(archive, "_rels/.rels")?)?;
    let mut books = root.values().filter(|rel| rel.kind == "officeDocument" && !rel.external);
    let book = books.next().ok_or("SPARKLINE_OPC_INVALID: missing internal workbook relationship")?;
    if books.next().is_some() { return Err("SPARKLINE_OPC_INVALID: multiple workbook relationships".into()); }
    let book_path = resolve_target("", &book.target).ok_or("SPARKLINE_OPC_INVALID: invalid workbook target")?;
    let rels = relationships(&read_xml(archive, &relationship_path(&book_path))?)?;
    let mut resources = ColorResources::default();
    if let Some(theme) = one_resource(&rels, "theme")? {
        let path = resolve_target(&book_path, &theme.target).ok_or("SPARKLINE_OPC_INVALID: invalid theme target")?;
        let source = read_xml(archive, &path)?;
        if source.name != "theme" || !matches!(source.namespace.as_str(), "http://schemas.openxmlformats.org/drawingml/2006/main" | "http://purl.oclc.org/ooxml/drawingml/main") {
            return Err("SPARKLINE_COLOR_UNRESOLVED: invalid workbook theme root".into());
        }
        resources.theme_xml = Some(source.render()); resources.theme = Some(source);
    }
    if let Some(styles) = one_resource(&rels, "styles")? {
        let path = resolve_target(&book_path, &styles.target).ok_or("SPARKLINE_OPC_INVALID: invalid styles target")?;
        let entry = archive.by_name(&path).map_err(|error| format!("SPARKLINE_COLOR_UNRESOLVED: unavailable styles part {path}: {error}"))?;
        let selected = xml::extract(BufReader::new(entry), XmlSelectionLimits::default(), |path, _| {
            Ok(path.len() == 3 && sheet(&path[0].namespace) && path[0].name == "styleSheet"
                && sheet(&path[1].namespace) && path[1].name == "colors"
                && sheet(&path[2].namespace) && path[2].name == "indexedColors")
        })?;
        if selected.root.name != "styleSheet" || !sheet(&selected.root.namespace) { return Err("SPARKLINE_COLOR_UNRESOLVED: invalid styles root".into()); }
        if selected.fragments.len() > 1 { return Err("SPARKLINE_COLOR_UNRESOLVED: multiple indexed color palettes".into()); }
        if let Some(palette) = selected.fragments.first() {
            if palette.children.len() > 256 { return Err("SPARKLINE_SOURCE_LIMIT: indexed palette exceeds 256 entries".into()); }
            let mut colors = Vec::new();
            for color in &palette.children {
                if color.name != "rgbColor" || !sheet(&color.namespace) { return Err("SPARKLINE_COLOR_UNRESOLVED: invalid indexed color entry".into()); }
                let raw = color.attribute_ns("", "rgb").ok_or("SPARKLINE_COLOR_UNRESOLVED: indexed entry has no ARGB")?;
                if office_io_common::sparkline_io::argb(raw).is_none() { return Err("SPARKLINE_COLOR_UNRESOLVED: invalid indexed ARGB".into()); }
                colors.push(if raw.len() == 6 { format!("FF{}", raw.to_ascii_uppercase()) } else { raw.to_ascii_uppercase() });
            }
            resources.indexed_colors = Some(colors);
        }
    }
    Ok(resources)
}

fn sheet(namespace: &str) -> bool { matches!(namespace, xml::MAIN_NAMESPACE | xml::STRICT_MAIN_NAMESPACE) }
fn one_resource<'a>(rels: &'a BTreeMap<String, Relationship>, kind: &str) -> Result<Option<&'a Relationship>, String> {
    let mut matching = rels.values().filter(|rel| rel.kind == kind);
    let first = matching.next();
    if first.is_some_and(|rel| rel.external) || matching.next().is_some() { return Err(format!("SPARKLINE_OPC_INVALID: invalid/multiple {kind} resources")); }
    Ok(first)
}
fn read_xml(archive: &mut ZipArchive<Cursor<&[u8]>>, path: &str) -> Result<Element, String> {
    let entry = archive.by_name(path).map_err(|error| format!("SPARKLINE_OPC_INVALID: unavailable part {path}: {error}"))?;
    if entry.size() > MAX_METADATA_BYTES { return Err(format!("SPARKLINE_SOURCE_LIMIT: metadata part {path} exceeds 16 MiB")); }
    let mut text = String::new(); entry.take(MAX_METADATA_BYTES + 1).read_to_string(&mut text).map_err(|error| error.to_string())?;
    if text.len() as u64 > MAX_METADATA_BYTES { return Err(format!("SPARKLINE_SOURCE_LIMIT: metadata part {path} exceeds 16 MiB")); }
    xml::parse(&text)
}
fn relationships(root: &Element) -> Result<BTreeMap<String, Relationship>, String> {
    if root.name != "Relationships" || !matches!(root.namespace.as_str(), PACKAGE | STRICT_PACKAGE) { return Err("SPARKLINE_OPC_INVALID: invalid relationships root".into()); }
    let mut result = BTreeMap::new();
    for node in &root.children {
        if node.name != "Relationship" || !matches!(node.namespace.as_str(), PACKAGE | STRICT_PACKAGE) { continue; }
        let id = node.attribute_ns("", "Id").filter(|value| !value.is_empty()).ok_or("SPARKLINE_OPC_INVALID: missing relationship ID")?;
        let kind = node.attribute_ns("", "Type").ok_or("SPARKLINE_OPC_INVALID: missing relationship type")?;
        let kind = kind.strip_prefix(REL).or_else(|| kind.strip_prefix(STRICT_REL)).unwrap_or("");
        let target = node.attribute_ns("", "Target").ok_or("SPARKLINE_OPC_INVALID: missing relationship target")?;
        let external = match node.attribute_ns("", "TargetMode") { None | Some("Internal") => false, Some("External") => true, _ => return Err("SPARKLINE_OPC_INVALID: invalid relationship target mode".into()) };
        if result.insert(id.into(), Relationship { kind: kind.into(), target: target.into(), external }).is_some() { return Err("SPARKLINE_OPC_INVALID: duplicate relationship ID".into()); }
    }
    Ok(result)
}
fn relationship_path(source: &str) -> String {
    match source.rsplit_once('/') { Some((directory, file)) => format!("{directory}/_rels/{file}.rels"), None => format!("_rels/{source}.rels") }
}
fn resolve_target(source: &str, target: &str) -> Option<String> {
    let mut decoded = Vec::with_capacity(target.len()); let mut input = target.bytes();
    while let Some(byte) = input.next() {
        decoded.push(if byte == b'%' { (char::from(input.next()?).to_digit(16)? * 16 + char::from(input.next()?).to_digit(16)?) as u8 } else { byte });
    }
    let target = std::str::from_utf8(&decoded).ok()?;
    if target.contains([':', '\\', '?', '#']) || target.starts_with("//") || target.chars().any(char::is_control) { return None; }
    let mut parts = if target.starts_with('/') { Vec::new() } else { source.rsplit_once('/').map(|(directory, _)| directory.split('/').collect()).unwrap_or_default() };
    for part in target.trim_start_matches('/').split('/') { match part { "" | "." => {}, ".." => { parts.pop()?; }, value => parts.push(value) } }
    (!parts.is_empty()).then(|| parts.join("/"))
}
