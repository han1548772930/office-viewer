use std::collections::HashMap;
use std::io::Read;

use quick_xml::events::Event;
use quick_xml::{Reader, XmlVersion};
use serde_json::{Value, json};
use zip::ZipArchive;
use crate::model::ChartAxisPayload;

#[derive(Debug, Default)]
pub(crate) struct PivotSidecar {
    pub(crate) pivot_tables: Vec<Value>,
    pub(crate) chart_pivot_ids: HashMap<(String, usize), String>,
    pub(crate) chart_settings: HashMap<(String, usize), ChartSettings>,
    pub(crate) warnings: Vec<String>,
}

#[derive(Debug)]
pub(crate) struct ChartSettings {
    pub(crate) kind: Option<String>,
    pub(crate) x_axis: ChartAxisPayload,
    pub(crate) y_axis: ChartAxisPayload,
    pub(crate) legend_position: Option<String>,
}

#[derive(Debug, Default)]
struct Node {
    name: String,
    attrs: HashMap<String, String>,
    text: String,
    children: Vec<Node>,
}

#[derive(Debug)]
struct Relationship {
    id: String,
    kind: String,
    target: String,
    external: bool,
}

impl PivotSidecar {
    pub(crate) fn read(bytes: &[u8]) -> Self {
        match read_sidecar(bytes) {
            Ok(value) => value,
            Err(error) => Self {
                warnings: vec![error],
                ..Self::default()
            },
        }
    }
}

fn read_sidecar(bytes: &[u8]) -> Result<PivotSidecar, String> {
    let mut archive = ZipArchive::new(std::io::Cursor::new(bytes))
        .map_err(|_| "Invalid XLSX ZIP while reading pivot parts".to_owned())?;
    let workbook = parse_part(&mut archive, "xl/workbook.xml")?
        .ok_or_else(|| "Missing XLSX workbook part while reading pivots".to_owned())?;
    let workbook_rels = parse_part(&mut archive, "xl/_rels/workbook.xml.rels")?.unwrap_or_default();
    let relation_map = relationships(&workbook_rels)
        .into_iter()
        .filter(|rel| !rel.external)
        .filter_map(|rel| {
            resolve("xl/workbook.xml", &rel.target).map(|path| (rel.id, (rel.kind, path)))
        })
        .collect::<HashMap<_, _>>();
    let mut sidecar = PivotSidecar::default();
    let mut table_names = HashMap::<String, String>::new();
    let mut pending_chart_links = Vec::<(String, usize, String)>::new();
    let sheets = workbook
        .descendants("sheet")
        .into_iter()
        .filter_map(|node| {
            Some((
                node.attrs.get("name")?.clone(),
                node.attrs.get("id")?.clone(),
            ))
        })
        .collect::<Vec<_>>();
    for (sheet_name, rel_id) in sheets {
        let Some((kind, sheet_path)) = relation_map.get(&rel_id) else {
            continue;
        };
        if kind != "worksheet" {
            continue;
        }
        let Some(sheet_xml) = parse_part(&mut archive, sheet_path)? else {
            continue;
        };
        let sheet_rels_path = relationship_part_path(sheet_path);
        let sheet_rels = parse_part(&mut archive, &sheet_rels_path)?.unwrap_or_default();
        let sheet_rel_map = relationships(&sheet_rels)
            .into_iter()
            .filter(|rel| !rel.external)
            .filter_map(|rel| {
                resolve(sheet_path, &rel.target).map(|path| (rel.id, (rel.kind, path)))
            })
            .collect::<HashMap<_, _>>();
        if let Some((_, drawing_path)) = sheet_rel_map.values().find(|(kind, _)| kind == "drawing")
        {
            if let Some(drawing_xml) = parse_part(&mut archive, drawing_path)? {
                let drawing_rels = parse_part(&mut archive, &relationship_part_path(drawing_path))?
                    .unwrap_or_default();
                let drawing_rel_map = relationships(&drawing_rels)
                    .into_iter()
                    .filter(|rel| !rel.external)
                    .filter_map(|rel| resolve(drawing_path, &rel.target).map(|path| (rel.id, path)))
                    .collect::<HashMap<_, _>>();
                for (index, chart_link) in drawing_xml.descendants("chart").into_iter().enumerate()
                {
                    let Some(rel_id) = chart_link.attrs.get("id") else {
                        continue;
                    };
                    let Some(chart_path) = drawing_rel_map.get(rel_id) else {
                        continue;
                    };
                    let Some(chart_xml) = parse_part(&mut archive, chart_path)? else {
                        continue;
                    };
                    let axis = |positions: &[&str]| {
                        chart_xml.descendants("valAx").into_iter().find(|node| {
                            node.descendants("axPos").into_iter().any(|position| position.attrs.get("val").is_some_and(|value| positions.contains(&value.as_str())))
                        }).map(|node| {
                            let log_base = node.descendants("logBase").into_iter().find_map(|base| base.attrs.get("val").and_then(|value| value.parse::<f64>().ok())).filter(|value| value.is_finite() && (2.0..=1000.0).contains(value));
                            ChartAxisPayload { scale_type: if log_base.is_some() { "log" } else { "linear" }.into(), log_base }
                        }).unwrap_or_else(|| ChartAxisPayload { scale_type: "linear".into(), log_base: None })
                    };
                    let legend_position = chart_xml
                        .descendants("legend")
                        .into_iter()
                        .next()
                        .and_then(|legend| legend.children.iter().find(|child| child.name == "legendPos"))
                        .and_then(|position| match position.attrs.get("val")?.as_str() {
                            "t" => Some("top"),
                            "r" => Some("right"),
                            "b" => Some("bottom"),
                            "l" => Some("left"),
                            _ => None,
                        })
                        .map(str::to_owned);
                    sidecar.chart_settings.insert(
                        (sheet_name.clone(), index),
                        ChartSettings {
                            kind: chart_kind(&chart_xml),
                            x_axis: axis(&["b", "t"]),
                            y_axis: axis(&["l", "r"]),
                            legend_position,
                        },
                    );
                    let Some(source_name) = chart_xml
                        .descendants("name")
                        .into_iter()
                        .find(|node| node.text.contains('!'))
                        .map(|node| node.text.clone())
                    else {
                        continue;
                    };
                    let table_name = source_name
                        .rsplit('!')
                        .next()
                        .unwrap_or_default()
                        .trim_start_matches('\'')
                        .trim_end_matches('\'')
                        .to_owned();
                    pending_chart_links.push((sheet_name.clone(), index, table_name));
                }
            }
        }
        for part in sheet_xml.descendants("pivotTablePart") {
            let Some(rid) = part.attrs.get("id") else {
                continue;
            };
            let Some((part_kind, table_path)) = sheet_rel_map.get(rid) else {
                continue;
            };
            if part_kind != "pivotTable" {
                continue;
            }
            let Some(table_xml) = parse_part(&mut archive, table_path)? else {
                continue;
            };
            let table_rels_path = relationship_part_path(table_path);
            let table_rels = parse_part(&mut archive, &table_rels_path)?.unwrap_or_default();
            let cache_rel = relationships(&table_rels)
                .into_iter()
                .find(|rel| rel.kind == "pivotCacheDefinition" && !rel.external)
                .and_then(|rel| resolve(table_path, &rel.target));
            let Some(cache_path) = cache_rel else {
                sidecar.warnings.push(format!(
                    "Pivot table part {table_path} has no pivot cache definition relation"
                ));
                continue;
            };
            let Some(cache_xml) = parse_part(&mut archive, &cache_path)? else {
                continue;
            };
            let source = cache_xml.descendants("worksheetSource").into_iter().next();
            let source_range = source
                .and_then(|node| node.attrs.get("ref"))
                .and_then(|value| parse_range(value));
            let source_sheet = source.and_then(|node| node.attrs.get("sheet")).cloned();
            let Some(source_range) = source_range else {
                sidecar.warnings.push(format!(
                    "Pivot cache {cache_path} has no valid worksheet source range"
                ));
                continue;
            };
            let Some(target_range) = table_xml
                .descendants("location")
                .into_iter()
                .next()
                .and_then(|node| node.attrs.get("ref"))
                .and_then(|value| parse_range(value))
            else {
                sidecar.warnings.push(format!(
                    "Pivot table {table_path} has no valid output range"
                ));
                continue;
            };
            let fields = canonical_fields(&table_xml, &cache_xml, source_range[1]);
            let id = format!("pivot-{}", sidecar.pivot_tables.len() + 1);
            let name = table_xml
                .attrs
                .get("name")
                .cloned()
                .unwrap_or_else(|| id.clone());
            table_names.insert(name, id.clone());
            sidecar.pivot_tables.push(json!({
                "id": id,
                "source": {"unitId":"", "subUnitId": source_sheet.as_deref().unwrap_or(""), "sheetName": source_sheet, "range": {"startRow":source_range[0],"startCol":source_range[1],"endRow":source_range[2],"endCol":source_range[3]}},
                "anchor": {"unitId":"", "subUnitId":sheet_name, "row":target_range[0], "col":target_range[1]},
                "outputRange":{"startRow":target_range[0],"startCol":target_range[1],"endRow":target_range[2],"endCol":target_range[3]},
                "fields": fields,
                "options": {"repeatRowLabels":table_xml.attrs.get("compact").is_some_and(|value| value == "0"), "showRowGrandTotal":bool_attr(&table_xml,"rowGrandTotals",true), "showColGrandTotal":bool_attr(&table_xml,"colGrandTotals",true), "showRowSubtotals":subtotals_enabled(&table_xml,"axisRow"), "showColSubtotals":subtotals_enabled(&table_xml,"axisCol"), "valuePosition":"column"},
                "revision": 0
            }));
        }
    }
    for (sheet_name, index, table_name) in pending_chart_links {
        if let Some(id) = table_names.get(&table_name) {
            sidecar
                .chart_pivot_ids
                .insert((sheet_name, index), id.clone());
        }
    }
    Ok(sidecar)
}

fn chart_kind(chart: &Node) -> Option<String> {
    let plots = ["barChart", "lineChart", "areaChart", "pieChart", "doughnutChart", "scatterChart", "radarChart"]
        .into_iter().flat_map(|kind| chart.descendants(kind)).collect::<Vec<_>>();
    if plots.len() != 1 { return None; }
    let plot = plots[0];
    Some(match plot.name.as_str() {
        "barChart" if plot.children.iter().any(|child| child.name == "grouping" && child.attrs.get("val").is_some_and(|value| value == "stacked")) => "stacked-column",
        "barChart" => "bar",
        "lineChart" => "line",
        "areaChart" => "area",
        "pieChart" => "pie",
        "doughnutChart" => "doughnut",
        "scatterChart" => "scatter",
        "radarChart" => "radar",
        _ => return None,
    }.to_owned())
}

fn canonical_fields(table: &Node, cache: &Node, source_start_col: u32) -> Vec<Value> {
    let cache_fields = cache.descendants("cacheField");
    let areas = [
        ("rowFields", "row"),
        ("colFields", "column"),
        ("pageFields", "filter"),
        ("dataFields", "value"),
    ];
    let mut placements = HashMap::<usize, (String, usize, Option<String>, Option<String>)>::new();
    for (container_name, area) in areas {
        let Some(container) = table.descendants(container_name).into_iter().next() else {
            continue;
        };
        for (position, item) in container
            .children
            .iter()
            .filter(|node| {
                node.name == "field" || node.name == "pageField" || node.name == "dataField"
            })
            .enumerate()
        {
            let field = item
                .attrs
                .get("x")
                .or_else(|| item.attrs.get("fld"))
                .and_then(|value| value.parse::<usize>().ok());
            let Some(field) = field else { continue };
            placements.insert(
                field,
                (
                    area.to_owned(),
                    position,
                    item.attrs.get("name").cloned(),
                    item.attrs.get("subtotal").cloned(),
                ),
            );
        }
    }
    cache_fields.iter().enumerate().map(|(index, field)| {
        let placement = placements.get(&index);
        let name = field.attrs.get("name").cloned().unwrap_or_else(|| format!("Column {}", index + 1));
        let mut result = json!({"id":index.to_string(),"sourceColumn":source_start_col + index as u32,"name":name,"area":placement.map_or("unused", |value| value.0.as_str()),"position":placement.map_or(0, |value| value.1)});
        if let Some((_, _, Some(display_name), aggregation)) = placement.filter(|placement| placement.0 == "value") { result["displayName"] = json!(display_name); result["aggregation"] = json!(normalize_aggregation(aggregation.as_deref())); }
        result
    }).collect()
}

fn normalize_aggregation(value: Option<&str>) -> &'static str {
    match value {
        Some("count") => "count",
        Some("average") => "average",
        Some("min") => "min",
        Some("max") => "max",
        _ => "sum",
    }
}
fn subtotals_enabled(table: &Node, axis: &str) -> bool {
    table
        .descendants("pivotField")
        .into_iter()
        .filter(|field| field.attrs.get("axis").is_some_and(|value| value == axis))
        .all(|field| !field.attrs.get("defaultSubtotal").is_some_and(|value| value == "0"))
}
fn bool_attr(node: &Node, key: &str, default: bool) -> bool {
    node.attrs.get(key).map_or(default, |value| {
        value == "1" || value.eq_ignore_ascii_case("true")
    })
}

fn parse_range(value: &str) -> Option<[u32; 4]> {
    let (start, end) = value.split_once(':').unwrap_or((value, value));
    let (start_row, start_col) = parse_cell(start)?;
    let (end_row, end_col) = parse_cell(end)?;
    Some([start_row, start_col, end_row, end_col])
}

fn parse_cell(value: &str) -> Option<(u32, u32)> {
    let value = value.trim_matches('\'');
    let split = value.find(|character: char| character.is_ascii_digit())?;
    let (letters, digits) = value.split_at(split);
    let col = letters
        .bytes()
        .try_fold(0u32, |acc, byte| {
            acc.checked_mul(26)?
                .checked_add(u32::from(byte.to_ascii_uppercase() - b'A' + 1))
        })?
        .checked_sub(1)?;
    Some((digits.parse::<u32>().ok()?.checked_sub(1)?, col))
}

fn parse_part(
    archive: &mut ZipArchive<std::io::Cursor<&[u8]>>,
    path: &str,
) -> Result<Option<Node>, String> {
    let Ok(mut file) = archive.by_name(path) else {
        return Ok(None);
    };
    let mut source = String::new();
    file.read_to_string(&mut source)
        .map_err(|error| error.to_string())?;
    parse_xml(&source).map(Some)
}

fn parse_xml(source: &str) -> Result<Node, String> {
    let mut reader = Reader::from_str(source);
    reader.config_mut().trim_text(true);
    let mut stack = vec![Node::default()];
    loop {
        match reader.read_event() {
            Ok(Event::Start(element)) => {
                let node = node_from(&reader, &element)?;
                stack.push(node);
            }
            Ok(Event::Empty(element)) => {
                let node = node_from(&reader, &element)?;
                stack.last_mut().unwrap().children.push(node);
            }
            Ok(Event::Text(text)) => {
                let decoded = text.decode().map_err(|error| error.to_string())?;
                if let Some(node) = stack.last_mut() {
                    node.text.push_str(&decoded);
                }
            }
            Ok(Event::CData(text)) => {
                let decoded = text.decode().map_err(|error| error.to_string())?;
                if let Some(node) = stack.last_mut() {
                    node.text.push_str(&decoded);
                }
            }
            Ok(Event::End(_)) => {
                if stack.len() > 1 {
                    let node = stack.pop().unwrap();
                    stack.last_mut().unwrap().children.push(node);
                }
            }
            Ok(Event::Eof) => break,
            Err(error) => return Err(error.to_string()),
            _ => {}
        }
    }
    Ok(stack
        .pop()
        .and_then(|node| node.children.into_iter().next())
        .unwrap_or_default())
}

fn node_from(
    reader: &Reader<&[u8]>,
    element: &quick_xml::events::BytesStart<'_>,
) -> Result<Node, String> {
    let name = local_name(element.name().as_ref());
    let attrs = element
        .attributes()
        .with_checks(false)
        .filter_map(Result::ok)
        .map(|attribute| {
            let key = local_name(attribute.key.as_ref());
            let value = attribute
                .decoded_and_normalized_value_with(
                    XmlVersion::Implicit1_0,
                    reader.decoder(),
                    1,
                    quick_xml::escape::resolve_xml_entity,
                )
                .ok()?
                .into_owned();
            Some((key, value))
        })
        .flatten()
        .collect();
    Ok(Node {
        name,
        attrs,
        ..Node::default()
    })
}

fn local_name(value: &[u8]) -> String {
    String::from_utf8_lossy(value.rsplit(|byte| *byte == b':').next().unwrap_or(value)).into_owned()
}
impl Node {
    fn descendants(&self, name: &str) -> Vec<&Node> {
        let mut result = Vec::new();
        if self.name == name {
            result.push(self);
        }
        for child in &self.children {
            result.extend(child.descendants(name));
        }
        result
    }
}

fn relationships(root: &Node) -> Vec<Relationship> {
    root.descendants("Relationship")
        .into_iter()
        .filter_map(|node| {
            Some(Relationship {
                id: node.attrs.get("Id")?.clone(),
                kind: node.attrs.get("Type")?.rsplit('/').next()?.to_owned(),
                target: node.attrs.get("Target")?.clone(),
                external: node
                    .attrs
                    .get("TargetMode")
                    .is_some_and(|value| value == "External"),
            })
        })
        .collect()
}
fn relationship_part_path(source: &str) -> String {
    let (folder, file) = source.rsplit_once('/').unwrap_or(("", source));
    format!("{folder}/_rels/{file}.rels")
}
fn resolve(source: &str, target: &str) -> Option<String> {
    if target.starts_with('/') {
        return Some(target.trim_start_matches('/').to_owned());
    }
    let source_dir = source.rsplit_once('/').map_or("", |(folder, _)| folder);
    let combined = format!("{source_dir}/{target}");
    let mut components = Vec::<String>::new();
    for component in combined.split('/') {
        match component {
            "" | "." => {}
            ".." => {
                components.pop()?;
            }
            value => components.push(value.to_owned()),
        }
    }
    Some(components.join("/"))
}
