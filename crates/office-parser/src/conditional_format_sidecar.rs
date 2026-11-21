use std::collections::HashMap;
use std::io::Read;

use quick_xml::events::Event;
use quick_xml::{Reader, XmlVersion};
use serde_json::{Value, json};
use zip::ZipArchive;

#[derive(Debug, Default)]
pub(crate) struct ConditionalFormatSidecar {
    by_sheet: HashMap<usize, Vec<Value>>,
    pub(crate) warnings: Vec<String>,
}

#[derive(Debug, Default, Clone)]
struct Node {
    name: String,
    attrs: HashMap<String, String>,
    text: String,
    children: Vec<Node>,
}

#[derive(Debug)]
struct Relationship {
    id: String,
    target: String,
    kind: String,
    external: bool,
}

impl ConditionalFormatSidecar {
    pub(crate) fn read(bytes: &[u8]) -> Self {
        match read_sidecar(bytes) {
            Ok(value) => value,
            Err(error) => Self {
                warnings: vec![error],
                ..Self::default()
            },
        }
    }

    pub(crate) fn take(&mut self, sheet_index: usize) -> Vec<Value> {
        self.by_sheet.remove(&sheet_index).unwrap_or_default()
    }
}

fn read_sidecar(bytes: &[u8]) -> Result<ConditionalFormatSidecar, String> {
    let mut archive = ZipArchive::new(std::io::Cursor::new(bytes))
        .map_err(|_| "Invalid XLSX ZIP while reading conditional formatting".to_owned())?;
    let workbook = parse_part(&mut archive, "xl/workbook.xml")?
        .ok_or_else(|| "Missing XLSX workbook part while reading conditional formatting".to_owned())?;
    let workbook_rels = parse_part(&mut archive, "xl/_rels/workbook.xml.rels")?.unwrap_or_default();
    let relation_map = relationships(&workbook_rels)
        .into_iter()
        .filter(|relationship| !relationship.external)
        .filter_map(|relationship| {
            resolve("xl/workbook.xml", &relationship.target)
                .map(|path| (relationship.id, (relationship.kind, path)))
        })
        .collect::<HashMap<_, _>>();
    let style_path = relation_map
        .values()
        .find_map(|(kind, path)| (kind == "styles").then_some(path.as_str()));
    let dxfs = match style_path {
        Some(path) => parse_part(&mut archive, path)?
            .map(|xml| parse_xml(&xml).descendants("dxf"))
            .unwrap_or_default(),
        None => Vec::new(),
    };
    let sheets = parse_xml(&workbook).descendants("sheet");
    let mut sidecar = ConditionalFormatSidecar::default();
    for (index, sheet) in sheets.iter().enumerate() {
        let Some(relationship_id) = sheet.attrs.get("id") else { continue };
        let Some((kind, path)) = relation_map.get(relationship_id) else { continue };
        if kind != "worksheet" { continue }
        let Some(xml) = parse_part(&mut archive, path)? else { continue };
        let rules = parse_rules(&xml, &dxfs, &mut sidecar.warnings);
        if !rules.is_empty() {
            sidecar.by_sheet.insert(index, rules);
        }
    }
    Ok(sidecar)
}

fn parse_rules(xml: &str, dxfs: &[Node], warnings: &mut Vec<String>) -> Vec<Value> {
    let root = parse_xml(xml);
    let mut rules = Vec::new();
    for group in root.descendants("conditionalFormatting") {
        let ranges = group
            .attrs
            .get("sqref")
            .into_iter()
            .flat_map(|value| value.split_whitespace())
            .filter_map(parse_range)
            .collect::<Vec<_>>();
        if ranges.is_empty() {
            warnings.push("Skipped XLSX conditional formatting with invalid sqref".into());
            continue;
        }
        for rule in group.children.iter().filter(|node| node.name == "cfRule") {
            let xml_type = rule.attrs.get("type").map(String::as_str).unwrap_or_default();
            let (rule_type, sub_type) = match xml_type {
                "cellIs" => ("cellIs", Some("number")), "expression" => ("expression", Some("formula")),
                "containsText" => match rule.attrs.get("operator").map(String::as_str).unwrap_or("containsText") {
                    "notContains" => ("notContainsText", Some("text")), "beginsWith" => ("beginsWith", Some("text")),
                    "endsWith" => ("endsWith", Some("text")), _ => ("containsText", Some("text")),
                },
                "duplicateValues" => ("duplicateValues", None), "uniqueValues" => ("uniqueValues", None),
                "timePeriod" => ("timePeriod", Some("timePeriod")), "top10" => ("top10", Some("rank")),
                "aboveAverage" => ("average", Some("average")), "dataBar" => ("dataBar", None),
                "colorScale" => ("colorScale", None), "iconSet" => ("iconSet", None),
                _ => {
                    warnings.push(format!("Skipped unsupported XLSX conditional formatting rule type: {xml_type}"));
                    continue;
                }
            };
            let operator = rule.attrs.get("operator").or_else(|| rule.attrs.get("timePeriod")).map(String::as_str).unwrap_or_else(|| {
                if xml_type == "aboveAverage" {
                    if rule.attrs.get("equalAverage").is_some_and(|value| value == "1") { "equal" }
                    else if rule.attrs.get("aboveAverage").is_some_and(|value| value == "0") { "lessThan" } else { "greaterThan" }
                } else { "" }
            });
            if xml_type == "cellIs" && !matches!(operator, "equal" | "notEqual" | "greaterThan" | "greaterThanOrEqual" | "lessThan" | "lessThanOrEqual" | "between" | "notBetween") {
                warnings.push("Skipped XLSX conditional formatting with unsupported operator".into());
                continue;
            }
            let formulas = rule.children.iter().filter(|node| node.name == "formula").map(|node| node.text.trim().to_owned()).collect::<Vec<_>>();
            if xml_type == "cellIs" && (formulas.is_empty() || matches!(operator, "between" | "notBetween") && formulas.len() < 2) {
                warnings.push("Skipped XLSX conditional formatting with missing formula operands".into());
                continue;
            }
            let style = rule.attrs.get("dxfId")
                .and_then(|value| value.parse::<usize>().ok())
                .and_then(|index| dxfs.get(index))
                .map(differential_style)
                .unwrap_or_else(|| json!({}));
            if rule.attrs.contains_key("dxfId") && rule.attrs.get("dxfId").and_then(|value| value.parse::<usize>().ok()).is_none_or(|index| index >= dxfs.len()) {
                warnings.push("XLSX conditional formatting references an unavailable differential style".into());
            }
            let priority = rule.attrs.get("priority").and_then(|value| value.parse::<u32>().ok()).unwrap_or((rules.len() + 1) as u32);
            let config = parse_visual_config(rule);
            let formula1 = formulas.first().cloned().or_else(|| rule.attrs.get("text").cloned()).unwrap_or_default();
            let mut output = json!({
                "id": format!("conditional-{priority}-{}", rules.len() + 1), "type": rule_type,
                "ranges": ranges, "operator": operator, "formula1": formula1,
                "style": style, "priority": priority,
                "stopIfTrue": rule.attrs.get("stopIfTrue").is_some_and(|value| value == "1" || value.eq_ignore_ascii_case("true"))
            });
            if let Some(sub_type) = sub_type { output["subType"] = json!(sub_type); }
            if let Some(formula) = formulas.get(1) { output["formula2"] = json!(formula); }
            if let Some(rank) = rule.attrs.get("rank").and_then(|value| value.parse::<u32>().ok()) { output["value"] = json!(rank); }
            if rule.attrs.get("bottom").is_some_and(|value| value == "1") { output["isBottom"] = json!(true); }
            if rule.attrs.get("percent").is_some_and(|value| value == "1") { output["isPercent"] = json!(true); }
            if let Some(config) = config { output["config"] = config; }
            rules.push(output);
        }
    }
    rules
}

fn parse_visual_config(rule: &Node) -> Option<Value> {
    let data_bar = rule.descendants("dataBar").first().cloned();
    let color_scale = rule.descendants("colorScale").first().cloned();
    let icon_set = rule.descendants("iconSet").first().cloned();
    if let Some(node) = data_bar {
        let color = node.descendants("color").first().and_then(color_value).unwrap_or_else(|| "#638EC6".into());
        let values = node.children.iter().filter(|child| child.name == "cfvo").collect::<Vec<_>>();
        let min = values.first().map(|value| json!({ "type": value.attrs.get("type"), "value": value.attrs.get("val") }));
        let max = values.get(1).map(|value| json!({ "type": value.attrs.get("type"), "value": value.attrs.get("val") }));
        return Some(json!({ "positiveColor": color, "isShowValue": node.attrs.get("showValue").is_none_or(|value| value != "0"), "min": min, "max": max }));
    }
    if let Some(node) = color_scale {
        let cfvo = node.children.iter().filter(|child| child.name == "cfvo").collect::<Vec<_>>();
        let colors = node.children.iter().filter(|child| child.name == "color").collect::<Vec<_>>();
        let thresholds = cfvo.iter().zip(colors).enumerate().filter_map(|(index, (value, color))| Some(json!({
            "value": { "type": value.attrs.get("type")?, "value": value.attrs.get("val") },
            "color": color_value(color)? , "index": index
        }))).collect::<Vec<_>>();
        return Some(json!({ "thresholds": thresholds }));
    }
    if let Some(node) = icon_set {
        let values = node.children.iter().filter(|child| child.name == "cfvo").collect::<Vec<_>>();
        let icons = values.iter().enumerate().rev().map(|(index, value)| json!({
            "value": { "type": value.attrs.get("type"), "value": value.attrs.get("val") },
            "iconType": node.attrs.get("iconSet"), "iconId": values.len() - index - 1,
            "operator": "greaterThanOrEqual"
        })).collect::<Vec<_>>();
        return Some(json!({ "icons": icons, "iconSet": node.attrs.get("iconSet"), "reverse": node.attrs.get("reverse"), "isShowValue": node.attrs.get("showValue").is_none_or(|value| value != "0") }));
    }
    None
}

fn differential_style(node: &Node) -> Value {
    let font = node.children.iter().find(|child| child.name == "font");
    let fill = node.children.iter().find(|child| child.name == "fill");
    let mut style = serde_json::Map::new();
    if let Some(color) = font.and_then(|font| font.descendants("color").first().and_then(color_value)) {
        style.insert("fontColor".into(), Value::String(color));
    }
    if let Some(bold) = font.and_then(|font| font.children.iter().find(|child| child.name == "b")) {
        style.insert("bold".into(), Value::Bool(bold.attrs.get("val").is_none_or(|value| value != "0" && !value.eq_ignore_ascii_case("false"))));
    }
    for (xml_name, field) in [("i", "italic"), ("u", "underline"), ("strike", "strike")] {
        if let Some(node) = font.and_then(|font| font.children.iter().find(|child| child.name == xml_name)) {
            let enabled = node.attrs.get("val").is_none_or(|value| value != "0" && !value.eq_ignore_ascii_case("false"));
            style.insert(field.into(), Value::Bool(enabled));
        }
    }
    if let Some(color) = fill.and_then(|fill| fill.descendants("fgColor").first().and_then(color_value)) {
        style.insert("fill".into(), Value::String(color));
    }
    Value::Object(style)
}

fn color_value(node: &Node) -> Option<String> {
    let value = node.attrs.get("rgb")?;
    let value = value.strip_prefix('#').unwrap_or(value);
    if value.len() != 6 && value.len() != 8 || !value.bytes().all(|byte| byte.is_ascii_hexdigit()) { return None }
    Some(format!("#{}", if value.len() == 8 { &value[2..] } else { value }.to_ascii_uppercase()))
}

fn parse_range(value: &str) -> Option<Value> {
    let mut parts = value.split(':');
    let (start_row, start_col) = parse_cell(parts.next()?)?;
    let (end_row, end_col) = parse_cell(parts.next().unwrap_or(value))?;
    if parts.next().is_some() || end_row < start_row || end_col < start_col { return None }
    Some(json!({ "startRow": start_row, "startCol": start_col, "endRow": end_row, "endCol": end_col }))
}

fn parse_cell(value: &str) -> Option<(u32, u32)> {
    let value = value.trim().replace('$', "");
    let split = value.find(|character: char| character.is_ascii_digit())?;
    let (letters, digits) = value.split_at(split);
    if letters.is_empty() || !letters.bytes().all(|byte| byte.is_ascii_alphabetic()) { return None }
    let mut column = 0u32;
    for byte in letters.bytes() {
        column = column.checked_mul(26)?.checked_add(u32::from(byte.to_ascii_uppercase() - b'A' + 1))?;
    }
    let row = digits.parse::<u32>().ok()?;
    if row == 0 || row > 1_048_576 || column == 0 || column > 16_384 { return None }
    Some((row - 1, column - 1))
}

fn parse_xml(xml: &str) -> Node {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut stack = vec![Node::default()];
    loop {
        match reader.read_event() {
            Ok(Event::Start(element)) => {
                let node = Node { name: local_name(element.name().as_ref()).to_owned(), attrs: attributes(&reader, &element), ..Node::default() };
                stack.push(node);
            }
            Ok(Event::Empty(element)) => {
                let node = Node { name: local_name(element.name().as_ref()).to_owned(), attrs: attributes(&reader, &element), ..Node::default() };
                stack.last_mut().unwrap().children.push(node);
            }
            Ok(Event::Text(text)) => {
                if let Some(node) = stack.last_mut() {
                    let decoded = text.decode().unwrap_or_default();
                    node.text.push_str(&quick_xml::escape::unescape(&decoded).unwrap_or_else(|_| decoded.clone()));
                }
            }
            Ok(Event::End(_)) if stack.len() > 1 => {
                let node = stack.pop().unwrap();
                stack.last_mut().unwrap().children.push(node);
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    stack.pop().unwrap_or_default()
}

impl Node {
    fn descendants(&self, name: &str) -> Vec<Node> {
        let mut result = Vec::new();
        for child in &self.children {
            if child.name == name { result.push(child.clone()); }
            result.extend(child.descendants(name));
        }
        result
    }
}

fn relationships(xml: &str) -> Vec<Relationship> {
    parse_xml(xml).descendants("Relationship").into_iter().filter_map(|node| Some(Relationship {
        id: node.attrs.get("Id")?.clone(), target: node.attrs.get("Target")?.clone(),
        kind: node.attrs.get("Type")?.rsplit('/').next()?.to_owned(),
        external: node.attrs.get("TargetMode").is_some_and(|value| value == "External"),
    })).collect()
}

fn attributes(reader: &Reader<&[u8]>, element: &quick_xml::events::BytesStart<'_>) -> HashMap<String, String> {
    element.attributes().with_checks(false).filter_map(Result::ok).filter_map(|attribute| {
        let name = local_name(attribute.key.as_ref()).to_owned();
        let value = attribute.decoded_and_normalized_value_with(XmlVersion::Implicit1_0, reader.decoder(), 1, quick_xml::escape::resolve_xml_entity).ok()?.into_owned();
        Some((name, value))
    }).collect()
}

fn parse_part(archive: &mut ZipArchive<std::io::Cursor<&[u8]>>, path: &str) -> Result<Option<String>, String> {
    let mut entry = match archive.by_name(path) { Ok(entry) => entry, Err(zip::result::ZipError::FileNotFound) => return Ok(None), Err(error) => return Err(format!("Could not read XLSX package part {path}: {error}")) };
    if entry.size() > 16 * 1024 * 1024 { return Err(format!("XLSX metadata XML part {path} exceeds the parser limit")) }
    let mut xml = String::with_capacity(entry.size() as usize);
    entry.read_to_string(&mut xml).map_err(|error| format!("Could not decode XLSX package part {path}: {error}"))?;
    Ok(Some(xml))
}

fn resolve(source: &str, target: &str) -> Option<String> {
    if target.contains(':') || target.starts_with("//") { return None }
    let mut parts = if target.starts_with('/') { Vec::new() } else { source.rsplit_once('/').map(|(directory, _)| directory.split('/').collect()).unwrap_or_default() };
    for part in target.trim_start_matches('/').split('/') { match part { "" | "." => {}, ".." => { parts.pop()?; }, value => parts.push(value) } }
    (!parts.is_empty()).then(|| parts.join("/"))
}

fn local_name(value: &[u8]) -> &str {
    std::str::from_utf8(value.rsplit(|byte| *byte == b':').next().unwrap_or(value)).unwrap_or_default()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_conditional_format_ranges_and_differential_style() {
        let style = parse_xml("<styleSheet><dxfs count=\"1\"><dxf><font><b/><color rgb=\"FFFFFFFF\"/></font><fill><patternFill patternType=\"solid\"><fgColor rgb=\"FF00AA55\"/></patternFill></fill></dxf></dxfs></styleSheet>");
        let rule_xml = "<worksheet><conditionalFormatting sqref=\"$A$1:$A$5 $C$1:$C$5\"><cfRule type=\"cellIs\" dxfId=\"0\" priority=\"3\" operator=\"between\" stopIfTrue=\"1\"><formula>10</formula><formula>20</formula></cfRule></conditionalFormatting></worksheet>";
        let mut warnings = Vec::new();
        let styles = style.descendants("dxf");
        let rules = parse_rules(rule_xml, &styles, &mut warnings);
        assert!(warnings.is_empty());
        assert_eq!(rules[0]["ranges"], json!([
            { "startRow": 0, "startCol": 0, "endRow": 4, "endCol": 0 },
            { "startRow": 0, "startCol": 2, "endRow": 4, "endCol": 2 }
        ]));
        assert_eq!(rules[0]["formula1"], "10");
        assert_eq!(rules[0]["formula2"], "20");
        assert_eq!(rules[0]["priority"], 3);
        assert_eq!(rules[0]["stopIfTrue"], true);
        assert_eq!(rules[0]["style"], json!({ "fill": "#00AA55", "fontColor": "#FFFFFF", "bold": true }));
    }
}
