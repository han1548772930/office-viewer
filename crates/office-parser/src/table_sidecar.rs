use std::{collections::HashMap, io::Read};

use quick_xml::events::Event;
use quick_xml::Reader;
use serde_json::{Value, json};
use zip::ZipArchive;

#[derive(Debug, Default)]
pub(crate) struct TableSidecar {
    pub(crate) tables: Vec<Value>,
    pub(crate) warnings: Vec<String>,
}

#[derive(Debug, Default)]
struct Node { name: String, attrs: HashMap<String, String>, children: Vec<Node> }

#[derive(Debug)]
struct Rel { id: String, target: String, kind: String }

impl TableSidecar {
    pub(crate) fn read(bytes: &[u8]) -> Self {
        match read_tables(bytes) {
            Ok(tables) => Self { tables, warnings: Vec::new() },
            Err(error) => Self { tables: Vec::new(), warnings: vec![error] },
        }
    }
}

fn read_tables(bytes: &[u8]) -> Result<Vec<Value>, String> {
    let mut zip = ZipArchive::new(std::io::Cursor::new(bytes)).map_err(|_| "Invalid XLSX while reading tables".to_owned())?;
    let workbook = xml_part(&mut zip, "xl/workbook.xml")?.ok_or_else(|| "Missing workbook part while reading tables".to_owned())?;
    let workbook_rels = xml_part(&mut zip, "xl/_rels/workbook.xml.rels")?.unwrap_or_default();
    let workbook_rels_map = rels(&workbook_rels).into_iter().filter_map(|rel| resolve("xl/workbook.xml", &rel.target).map(|path| (rel.id, (rel.kind, path)))).collect::<HashMap<_, _>>();
    let mut tables = Vec::new();
    for (sheet_index, sheet) in parse(&workbook).descendants("sheet").iter().enumerate() {
        let Some(sheet_name) = sheet.attrs.get("name") else { continue };
        let Some(rid) = sheet.attrs.get("id") else { continue };
        let Some((kind, sheet_path)) = workbook_rels_map.get(rid) else { continue };
        if kind != "worksheet" { continue }
        let Some(sheet_xml) = xml_part(&mut zip, sheet_path)? else { continue };
        let sheet_rels_path = relationship_path(sheet_path);
        let sheet_rels_xml = xml_part(&mut zip, &sheet_rels_path)?.unwrap_or_default();
        let sheet_rels = rels(&sheet_rels_xml).into_iter().filter_map(|rel| resolve(sheet_path, &rel.target).map(|path| (rel.id, path))).collect::<HashMap<_, _>>();
        for table_part in parse(&sheet_xml).descendants("tablePart") {
            let Some(table_rid) = table_part.attrs.get("id") else { continue };
            let Some(table_path) = sheet_rels.get(table_rid) else { continue };
            let Some(table_xml) = xml_part(&mut zip, table_path)? else { continue };
            let root = parse(&table_xml);
            let Some(table) = root.children.iter().find(|node| node.name == "table") else { continue };
            let Some(name) = table.attrs.get("displayName").or_else(|| table.attrs.get("name")) else { continue };
            let Some(reference) = table.attrs.get("ref") else { continue };
            let Some(range) = parse_range(reference) else { continue };
            let columns = table.descendants("tableColumn").iter().enumerate().map(|(index, column)| json!({"id":format!("{name}-column-{}",index+1),"name":column.attrs.get("name").cloned().unwrap_or_else(||format!("Column{}",index+1))})).collect::<Vec<_>>();
            if columns.len() != (range[3] - range[1] + 1) as usize {
                continue;
            }
            let style_name = table.descendants("tableStyleInfo").first().and_then(|style| style.attrs.get("name")).cloned().unwrap_or_else(|| "TableStyleMedium2".to_owned());
            tables.push(json!({
                "id":name,"name":name,"sheetName":sheet_name,"sheetIndex":sheet_index,"range":{
                    "startRow":range[0],"startCol":range[1],"endRow":range[2],"endCol":range[3]
                },"columns":columns,"showHeaderRow":table.attrs.get("headerRowCount").is_none_or(|value| value != "0"),
                "showTotalsRow":table.attrs.get("totalsRowShown").is_some_and(|value| value == "1"),"styleName":style_name
            }));
        }
    }
    Ok(tables)
}

fn parse_range(reference: &str) -> Option<[u32; 4]> {
    let (start, end) = reference.split_once(':').unwrap_or((reference, reference));
    let (start_row, start_col) = cell(start)?;
    let (end_row, end_col) = cell(end)?;
    Some([start_row, start_col, end_row, end_col])
}

fn cell(value: &str) -> Option<(u32, u32)> {
    let split = value.find(|char: char| char.is_ascii_digit())?;
    let (letters, digits) = value.split_at(split);
    let mut column = 0u32;
    for byte in letters.bytes() { if !byte.is_ascii_alphabetic() { return None } column = column.checked_mul(26)?.checked_add((byte.to_ascii_uppercase() - b'A' + 1) as u32)?; }
    Some((digits.parse::<u32>().ok()?.checked_sub(1)?, column.checked_sub(1)?))
}

fn xml_part(zip: &mut ZipArchive<std::io::Cursor<&[u8]>>, path: &str) -> Result<Option<String>, String> {
    let Ok(mut file) = zip.by_name(path) else { return Ok(None) };
    let mut xml = String::new(); file.read_to_string(&mut xml).map_err(|_| format!("Cannot read XLSX part {path}"))?; Ok(Some(xml))
}

fn rels(xml: &str) -> Vec<Rel> {
    parse(xml).descendants("Relationship").into_iter().filter_map(|node| Some(Rel { id: node.attrs.get("Id")?.clone(), target: node.attrs.get("Target")?.clone(), kind: node.attrs.get("Type")?.rsplit('/').next()?.to_owned() })).collect()
}

fn relationship_path(part: &str) -> String {
    let (directory, file) = part.rsplit_once('/').unwrap_or(("", part));
    if directory.is_empty() { format!("_rels/{part}.rels") } else { format!("{directory}/_rels/{file}.rels") }
}

fn resolve(source: &str, target: &str) -> Option<String> {
    let mut parts = Vec::<&str>::new();
    if !target.starts_with('/') { parts.extend(source.split('/').take(source.split('/').count().saturating_sub(1))); }
    for part in target.trim_start_matches('/').split('/') { match part { "" | "." => (), ".." => { parts.pop()?; }, other => parts.push(other) } }
    Some(parts.join("/"))
}

fn parse(xml: &str) -> Node {
    let mut reader = Reader::from_str(xml); reader.config_mut().trim_text(true);
    let mut stack = vec![Node::default()];
    loop { match reader.read_event() {
        Ok(Event::Start(event)) => stack.push(Node { name: local(event.name().as_ref()), attrs: attributes(&event), ..Node::default() }),
        Ok(Event::Empty(event)) => stack.last_mut().unwrap().children.push(Node { name: local(event.name().as_ref()), attrs: attributes(&event), ..Node::default() }),
        Ok(Event::End(_)) => { if stack.len() > 1 { let node = stack.pop().unwrap(); stack.last_mut().unwrap().children.push(node); } },
        Ok(Event::Eof) | Err(_) => break, _ => (),
    }}
    stack.pop().unwrap_or_default()
}

fn attributes(event: &quick_xml::events::BytesStart<'_>) -> HashMap<String, String> {
    event.attributes().flatten().filter_map(|attribute| Some((local(attribute.key.as_ref()), attribute.unescape_value().ok()?.into_owned()))).collect()
}
fn local(name: &[u8]) -> String { String::from_utf8_lossy(name).rsplit(':').next().unwrap_or_default().to_owned() }
impl Node { fn descendants(&self, name: &str) -> Vec<&Node> { let mut out = Vec::new(); if self.name == name { out.push(self); } for child in &self.children { out.extend(child.descendants(name)); } out } }
