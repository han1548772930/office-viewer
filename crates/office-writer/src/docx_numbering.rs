use std::collections::HashMap;

use serde_json::Value;

use super::{
    decode_data_image, get, integer, text_color, xml_attribute,
    ExportError, ExportWarning, Package, REL_IMAGE,
};
use crate::xml;

#[path = "docx_numbering_sequence.rs"]
mod sequence;
pub(super) use sequence::{prepare_sequence_groups, SequencePlan};

#[derive(Clone, Copy, PartialEq, Eq, Hash)]
enum ListKind {
    Preset(&'static str),
    ChecklistUnchecked,
    ChecklistChecked,
}

pub(super) struct NumberingInstance {
    id: u32,
    owner: usize,
    kind: ListKind,
    definition: Option<Value>,
    name: Option<String>,
}

#[derive(serde::Deserialize)]
struct ListDefinition {
    symbols: Option<Vec<String>>,
    formats: Option<Vec<String>>,
    suffix: Option<String>,
    compound: Option<bool>,
    levels: Option<usize>,
}

fn list_definitions() -> &'static std::collections::BTreeMap<String, ListDefinition> {
    static DEFINITIONS: std::sync::OnceLock<std::collections::BTreeMap<String, ListDefinition>> = std::sync::OnceLock::new();
    DEFINITIONS.get_or_init(|| serde_json::from_str(include_str!("../../../packages/office-studio/src/domain/document-list-presets.json")).expect("bundled list presets"))
}

fn array(value: &Value) -> &[Value] {
    value.as_array().map_or(&[], Vec::as_slice)
}

fn definition(block: &Value) -> Option<&Value> {
    let value = get(get(block, "metadata"), "numberingDefinition");
    (!array(get(value, "levels")).is_empty()).then_some(value)
}

fn abstract_definition(value: &Value) -> Value {
    let mut base = get(value, "abstract").clone();
    if !base.is_object() { base = serde_json::json!({}); }
    if array(get(&base, "levels")).is_empty() { base["levels"] = get(value, "levels").clone(); }
    base
}

fn kind_key(kind: ListKind) -> String {
    match kind {
        ListKind::Preset(key) => format!("preset:{key}"),
        ListKind::ChecklistUnchecked => "checklist:unchecked".to_owned(),
        ListKind::ChecklistChecked => "checklist:checked".to_owned(),
    }
}

fn abstract_key(kind: ListKind, definition: Option<&Value>) -> String {
    definition.map_or_else(|| kind_key(kind), |value| format!("custom:{}", abstract_definition(value)))
}

pub(super) fn numbering_level(block: &Value) -> i64 {
    let value = match get(block, "type").as_str() {
        Some("bullet" | "numbered" | "checklist") => get(block, "level"),
        _ => {
            let metadata = get(block, "metadata");
            let level = get(metadata, "numberingLevel");
            if level.is_number() { level } else { get(get(metadata, "numbering"), "ilvl") }
        }
    };
    integer(value).unwrap_or(0).clamp(0, 8)
}

fn merged_level(definition: &Value, level: &Value) -> Value {
    let base = abstract_definition(definition);
    let index = integer(get(level, "ilvl"));
    let mut merged = array(get(&base, "levels")).iter().find(|candidate| integer(get(candidate, "ilvl")) == index)
        .cloned().unwrap_or_else(|| serde_json::json!({}));
    if let Some(fields) = level.as_object() {
        for (key, value) in fields { merged[key] = value.clone(); }
    }
    merged
}

/// IDs are package-wide; named list counters belong to each independent story.
pub(super) fn assign_list_groups(groups: &[&[Value]]) -> Result<(HashMap<usize, u32>, Vec<NumberingInstance>), ExportError> {
    struct Assignments {
        lookup: HashMap<usize, u32>,
        numbering: Vec<NumberingInstance>,
        named: HashMap<(String, String), u32>,
        next_id: u32,
        owner: usize,
    }
    fn visit(blocks: &[Value], state: &mut Assignments) -> Result<(), ExportError> {
        let mut active: Option<(String, u32)> = None;
        for block in blocks {
            let block_type = get(block, "type").as_str().unwrap_or_default();
            if block_type == "table" {
                active = None;
                for group in super::table_writer::content_groups(get(get(block, "metadata"), "table"))? { visit(group.blocks, state)?; }
                continue;
            }
            let list = get(block, "list");
            let preset = get(list, "preset").as_str().and_then(|key| list_definitions().get_key_value(key)).map(|(key, _)| key.as_str());
            let kind = match block_type {
                "bullet" => ListKind::Preset(preset.filter(|key| key.starts_with("BULLET_")).unwrap_or("BULLET_LIST")),
                "numbered" => ListKind::Preset(preset.filter(|key| key.starts_with("ORDER_")).unwrap_or("ORDER_LIST")),
                "checklist" if get(get(block, "checklist"), "checked").as_bool() == Some(true) => ListKind::ChecklistChecked,
                "checklist" => ListKind::ChecklistUnchecked,
                "heading" | "paragraph" | "quote" if definition(block).is_some() || preset.is_some() || integer(get(get(get(block, "metadata"), "numbering"), "numId")).is_some_and(|id| id > 0) =>
                    ListKind::Preset(preset.unwrap_or("ORDER_LIST")),
                _ => { active = None; continue; }
            };
            let source = if block_type == "checklist" { None } else { definition(block) };
            let signature = abstract_key(kind, source);
            let metadata = get(block, "metadata");
            let name = get(list, "id").as_str().filter(|value| !value.is_empty())
                .or_else(|| get(metadata, "numberingListId").as_str().filter(|value| !value.is_empty()))
                .map(str::to_owned)
                .or_else(|| source.and_then(|value| integer(get(get(value, "instance"), "numId"))).map(|id| format!("num-{id}")))
                .or_else(|| integer(get(get(metadata, "numbering"), "numId")).filter(|id| *id > 0).map(|id| format!("num-{id}")))
                .or_else(|| get(get(block, "checklist"), "listId").as_str().filter(|value| !value.is_empty()).map(str::to_owned));
            let existing = name.as_ref().and_then(|name| state.named.get(&(signature.clone(), name.clone())).copied())
                .or_else(|| name.is_none().then(|| active.as_ref().filter(|(key, _)| key == &signature).map(|(_, id)| *id)).flatten());
            let id = existing.unwrap_or_else(|| {
                let id = state.next_id;
                state.next_id += 1;
                state.numbering.push(NumberingInstance { id, owner: state.owner, kind, definition: source.cloned(), name: name.clone() });
                id
            });
            if let Some(name) = name {
                state.named.insert((signature, name), id);
                active = None;
            } else { active = Some((signature, id)); }
            state.lookup.insert(block as *const Value as usize, id);
        }
        Ok(())
    }
    let mut state = Assignments { lookup: HashMap::new(), numbering: Vec::new(), named: HashMap::new(), next_id: 1, owner: 0 };
    for (owner, blocks) in groups.iter().enumerate() {
        state.owner = owner;
        state.named.clear();
        visit(blocks, &mut state)?;
    }
    Ok((state.lookup, state.numbering))
}

pub(super) fn numbering_left(instances: &[NumberingInstance], id: u32, level: i64) -> f64 {
    let Some(instance) = instances.iter().find(|instance| instance.id == id) else { return 0.0 };
    if let Some(definition) = &instance.definition {
        return array(get(definition, "levels")).iter().find(|item| integer(get(item, "ilvl")) == Some(level))
            .and_then(|item| get(get(item, "indent"), "leftTwips").as_f64()).unwrap_or(0.0) / 15.0;
    }
    match instance.kind {
        ListKind::Preset(_) => 28.0 * (level as f64 + 1.0),
        _ => 48.0 + level as f64 * 24.0,
    }
}

struct NumberingPicture {
    id: u32,
    relationship_id: String,
    width: f64,
    height: f64,
}

fn picture_key(value: &Value) -> Option<String> {
    let picture = get(value, "pictureBullet");
    get(picture, "src").as_str().filter(|src| !src.is_empty()).map(|_| picture.to_string())
}

pub(super) fn write_numbering_xml(instances: &[NumberingInstance], package: &mut Package, _warnings: &mut Vec<ExportWarning>) -> Result<String, ExportError> {
    let mut pictures = HashMap::new();
    for instance in instances {
        let Some(definition) = &instance.definition else { continue };
        let base = abstract_definition(definition);
        let overrides = array(get(get(definition, "instance"), "levelOverrides")).iter().map(|level| merged_level(definition, level)).collect::<Vec<_>>();
        for level in array(get(&base, "levels")).iter()
            .chain(&overrides)
            .chain(array(get(definition, "levels"))) {
            let format = get(level, "format").as_str().unwrap_or("decimal");
            let custom_format = get(level, "customFormat");
            if !custom_format.is_null() && custom_format.as_str().is_none_or(|pattern| pattern.len() > 4096) {
                return Err(ExportError::invalid("自定义编号格式必须是至多 4096 字节的字符串。"));
            }
            if format == "custom" && custom_format.as_str().is_none_or(str::is_empty) {
                return Err(ExportError::invalid("自定义编号必须提供非空的编号格式，不能生成缺少格式的 OOXML 定义。"));
            }
            let digits = get(level, "minimumDigits");
            if !digits.is_null() {
                let Some(digits) = digits.as_f64().filter(|digits| digits.is_finite() && digits.fract() == 0.0 && (0.0..=f64::from(u8::MAX)).contains(digits)) else {
                    return Err(ExportError::invalid("编号的最小数字位数必须是 0 到 255 之间的整数。"));
                };
                if digits > 0.0 && !matches!(format, "decimal" | "decimalZero" | "custom") && custom_format.is_null() {
                    return Err(ExportError::unsupported("当前 DOCX 写出尚不能保留非十进制编号格式上的补零位数。"));
                }
            }
            let Some(key) = picture_key(level) else {
                if integer(get(level, "pictureBulletId")).is_some() {
                    return Err(ExportError::unsupported("A picture bullet has no embedded image data."));
                }
                continue;
            };
            if pictures.contains_key(&key) { continue; }
            let picture = get(level, "pictureBullet");
            let src = get(picture, "src").as_str().unwrap_or_default();
            let (content_type, extension, bytes) = decode_data_image(src).map_err(ExportError::unsupported)?;
            let id = pictures.len() as u32 + 1;
            let relationship_id = format!("rIdNumberingPicture{id}");
            let target = format!("media/numbering-picture-{id}.{extension}");
            package.add_binary(&format!("word/{target}"), content_type, bytes)?;
            package.add_relationship(Some("word/numbering.xml"), &relationship_id, REL_IMAGE, &target, false)?;
            let dimension = |name: &str, fallback: f64| get(picture, name).as_f64().filter(|value| value.is_finite() && *value > 0.0).unwrap_or(fallback);
            let height = dimension("heightPx", 14.67);
            pictures.insert(key, NumberingPicture { id, relationship_id, width: dimension("widthPx", height), height });
        }
    }
    Ok(build_with_pictures(instances, &pictures))
}

pub(super) fn picture_shape_type_xml() -> &'static str { PICTURE_SHAPE_TYPE }

#[cfg(test)]
pub(super) fn build_numbering_xml(instances: &[NumberingInstance]) -> String {
    build_with_pictures(instances, &HashMap::new())
}

fn build_with_pictures(instances: &[NumberingInstance], pictures: &HashMap<String, NumberingPicture>) -> String {
    let defaults = [ListKind::Preset("BULLET_LIST"), ListKind::Preset("ORDER_LIST"), ListKind::ChecklistUnchecked, ListKind::ChecklistChecked];
    let mut abstract_ids = HashMap::new();
    let mut definitions = String::new();
    let mut next_abstract_id = 4u32;
    for (id, kind) in defaults.into_iter().enumerate() {
        abstract_ids.insert(kind_key(kind), id as u32);
        definitions.push_str(&preset_abstract(kind, id as u32));
    }
    for instance in instances {
        let signature = abstract_key(instance.kind, instance.definition.as_ref());
        if abstract_ids.contains_key(&signature) { continue; }
        let id = next_abstract_id;
        next_abstract_id += 1;
        abstract_ids.insert(signature, id);
        if let Some(definition) = &instance.definition {
            let base = abstract_definition(definition);
            let multi_level_type = get(&base, "multiLevelType").as_str().unwrap_or("multilevel");
            let levels = array(get(&base, "levels")).iter().map(|level| level_xml(level, pictures)).collect::<String>();
            let name = get(&base, "name").as_str().map(|name| format!("<w:name w:val=\"{}\"/>", xml_attribute(name))).unwrap_or_default();
            definitions.push_str(&format!("<w:abstractNum w:abstractNumId=\"{id}\"><w:multiLevelType w:val=\"{}\"/>{name}{levels}</w:abstractNum>", xml_attribute(multi_level_type)));
        } else { definitions.push_str(&preset_abstract(instance.kind, id)); }
    }
    let mut numbers = String::new();
    for instance in instances {
        let abstract_id = abstract_ids[&abstract_key(instance.kind, instance.definition.as_ref())];
        let mut overrides = std::collections::BTreeMap::<i64, (Option<i64>, Option<Value>)>::new();
        if let Some(definition) = &instance.definition {
            let source = get(definition, "instance");
            for level in array(get(source, "levelOverrides")) {
                if let Some(index) = integer(get(level, "ilvl")).filter(|index| (0..=8).contains(index)) { overrides.entry(index).or_default().1 = Some(merged_level(definition, level)); }
            }
            if let Some(starts) = get(source, "levelStartOverrides").as_object() {
                for (index, value) in starts {
                    if let (Ok(index), Some(start)) = (index.parse::<i64>(), integer(value).filter(|start| (0..=i64::from(i32::MAX)).contains(start))) {
                        if (0..=8).contains(&index) { overrides.entry(index).or_default().0 = Some(start); }
                    }
                }
            }
        }
        let overrides = overrides.into_iter().map(|(index, (start, level))| {
            let start = start.map(|start| format!("<w:startOverride w:val=\"{start}\"/>")).unwrap_or_default();
            let level = level.as_ref().map(|level| level_xml(level, pictures)).unwrap_or_default();
            format!("<w:lvlOverride w:ilvl=\"{index}\">{start}{level}</w:lvlOverride>")
        }).collect::<String>();
        numbers.push_str(&format!("<w:num w:numId=\"{}\"><w:abstractNumId w:val=\"{abstract_id}\"/>{overrides}</w:num>", instance.id));
    }
    let mut picture_entries = pictures.values().collect::<Vec<_>>();
    picture_entries.sort_by_key(|picture| picture.id);
    let pictures = picture_entries.into_iter().enumerate().map(|(index, picture)| format!(
        "<w:numPicBullet w:numPicBulletId=\"{}\"><w:pict>{}<v:shape id=\"_x0000_i{}\" type=\"#_x0000_t75\" style=\"width:{}px;height:{}px\"><v:imagedata r:id=\"{}\" o:title=\"\"/></v:shape></w:pict></w:numPicBullet>",
        picture.id, if index == 0 { PICTURE_SHAPE_TYPE } else { "" }, picture.id + 1000, picture.width, picture.height, picture.relationship_id,
    )).collect::<String>();
    format!("{}<w:numbering xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">{pictures}{definitions}{numbers}</w:numbering>", xml::DECL)
}

const PICTURE_SHAPE_TYPE: &str = "<v:shapetype id=\"_x0000_t75\" coordsize=\"21600,21600\" o:spt=\"75\" o:preferrelative=\"t\" path=\"m@4@5l@4@11@9@11@9@5xe\" filled=\"f\" stroked=\"f\"><v:stroke joinstyle=\"miter\"/><v:formulas><v:f eqn=\"if lineDrawn pixelLineWidth 0\"/><v:f eqn=\"sum @0 1 0\"/><v:f eqn=\"sum 0 0 @1\"/><v:f eqn=\"prod @2 1 2\"/><v:f eqn=\"prod @3 21600 pixelWidth\"/><v:f eqn=\"prod @3 21600 pixelHeight\"/><v:f eqn=\"sum @0 0 1\"/><v:f eqn=\"prod @6 1 2\"/><v:f eqn=\"prod @7 21600 pixelWidth\"/><v:f eqn=\"sum @8 21600 0\"/><v:f eqn=\"prod @7 21600 pixelHeight\"/><v:f eqn=\"sum @10 21600 0\"/></v:formulas><v:path o:extrusionok=\"f\" gradientshapeok=\"t\" o:connecttype=\"rect\"/><o:lock v:ext=\"edit\" aspectratio=\"t\"/></v:shapetype>";

fn preset_abstract(kind: ListKind, id: u32) -> String {
    let definition = match kind { ListKind::Preset(key) => list_definitions().get(key), _ => None };
    let mut levels = String::new();
    for index in 0..definition.and_then(|definition| definition.levels).unwrap_or(9) {
        let (format, text, left, hanging, font) = if let Some(definition) = definition {
            let (format, text) = if let Some(symbols) = &definition.symbols {
                ("bullet", symbols[index % symbols.len()].clone())
            } else {
                let formats = definition.formats.as_ref().expect("ordered preset formats");
                let slots = if definition.compound == Some(true) { 1..=index + 1 } else { index + 1..=index + 1 };
                let pattern = slots.map(|slot| format!("%{slot}")).collect::<Vec<_>>().join(".") + definition.suffix.as_deref().unwrap_or(".");
                (formats[index % formats.len()].as_str(), pattern)
            };
            (format, text, 420 * (index + 1), 420, "Arial")
        } else {
            ("bullet", if kind == ListKind::ChecklistChecked { "&#x2611;" } else { "&#x2610;" }.to_owned(), 720 + index * 360, 360, "Segoe UI Symbol")
        };
        let text = if definition.is_some() { xml_attribute(&text) } else { text };
        levels.push_str(&format!("<w:lvl w:ilvl=\"{index}\"><w:start w:val=\"1\"/><w:numFmt w:val=\"{format}\"/><w:suff w:val=\"tab\"/><w:lvlText w:val=\"{text}\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:tabs><w:tab w:val=\"num\" w:pos=\"{left}\"/></w:tabs><w:ind w:left=\"{left}\" w:hanging=\"{hanging}\"/></w:pPr><w:rPr><w:rFonts w:ascii=\"{font}\" w:hAnsi=\"{font}\"/></w:rPr></w:lvl>"));
    }
    format!("<w:abstractNum w:abstractNumId=\"{id}\"><w:multiLevelType w:val=\"multilevel\"/>{levels}</w:abstractNum>")
}

fn level_xml(level: &Value, pictures: &HashMap<String, NumberingPicture>) -> String {
    let Some(index) = integer(get(level, "ilvl")).filter(|index| (0..=8).contains(index)) else { return String::new() };
    let mut properties = String::new();
    if let Some(start) = integer(get(level, "start")).filter(|start| (0..=i64::from(i32::MAX)).contains(start)) { properties.push_str(&format!("<w:start w:val=\"{start}\"/>")); }
    let format = get(level, "format").as_str().unwrap_or("decimal");
    let minimum_digits = integer(get(level, "minimumDigits")).unwrap_or(0);
    let custom = if let Some(pattern) = get(level, "customFormat").as_str() {
        format!(" w:format=\"{}\"", xml_attribute(pattern))
    } else if minimum_digits > 0 && matches!(format, "decimal" | "decimalZero") {
        let width = minimum_digits.max(if format == "decimalZero" { 2 } else { 1 }).min(i64::from(u8::MAX)) as usize;
        format!(" w:format=\"{}1\"", "0".repeat(width - 1))
    } else { String::new() };
    properties.push_str(&format!("<w:numFmt w:val=\"{}\"{custom}/>", xml_attribute(format)));
    if let Some(restart) = integer(get(level, "lvlRestart")).filter(|restart| (0..=8).contains(restart)) { properties.push_str(&format!("<w:lvlRestart w:val=\"{restart}\"/>")); }
    if let Some(legal) = get(level, "isLgl").as_bool() { properties.push_str(&format!("<w:isLgl w:val=\"{}\"/>", u8::from(legal))); }
    if let Some(suffix) = get(level, "suffix").as_str() { properties.push_str(&format!("<w:suff w:val=\"{}\"/>", xml_attribute(suffix))); }
    if let Some(text) = get(level, "text").as_str() { properties.push_str(&format!("<w:lvlText w:val=\"{}\"/>", xml_attribute(text))); }
    if let Some(picture) = picture_key(level).and_then(|key| pictures.get(&key)) { properties.push_str(&format!("<w:lvlPicBulletId w:val=\"{}\"/>", picture.id)); }
    let alignment = get(level, "alignment").as_str().filter(|value| matches!(*value, "left" | "center" | "right" | "start" | "end")).unwrap_or("left");
    properties.push_str(&format!("<w:lvlJc w:val=\"{alignment}\"/>"));
    let indent = get(level, "indent");
    let mut indentation = String::new();
    for (field, attribute) in [("leftTwips", "left"), ("rightTwips", "right")] {
        if let Some(value) = integer(get(indent, field)) { indentation.push_str(&format!(" w:{attribute}=\"{value}\"")); }
    }
    let first = integer(get(indent, "firstLineTwips"));
    let hanging = integer(get(indent, "hangingTwips"));
    if first.is_some() || hanging.is_some() {
        let first = first.unwrap_or(0).saturating_sub(hanging.unwrap_or(0));
        let (attribute, value) = if first < 0 { ("hanging", first.saturating_neg()) } else { ("firstLine", first) };
        indentation.push_str(&format!(" w:{attribute}=\"{value}\""));
    }
    if !indentation.is_empty() {
        let tabs = integer(get(indent, "leftTwips")).filter(|_| get(level, "suffix").as_str().unwrap_or("tab") == "tab")
            .map(|left| format!("<w:tabs><w:tab w:val=\"num\" w:pos=\"{left}\"/></w:tabs>")).unwrap_or_default();
        properties.push_str(&format!("<w:pPr>{tabs}<w:ind{indentation}/></w:pPr>"));
    }
    let style = get(level, "runStyle");
    let mut run = String::new();
    let font = get(level, "bulletFontFamily").as_str().or_else(|| get(style, "fontFamilyAscii").as_str()).or_else(|| get(style, "resolvedFontFamilyAscii").as_str()).or_else(|| get(style, "fontFamily").as_str());
    if let Some(font) = font {
        let east_asia = get(style, "fontFamilyEastAsia").as_str().or_else(|| get(style, "resolvedFontFamilyEastAsia").as_str()).unwrap_or(font);
        run.push_str(&format!("<w:rFonts w:ascii=\"{}\" w:hAnsi=\"{}\" w:eastAsia=\"{}\"/>", xml_attribute(font), xml_attribute(font), xml_attribute(east_asia)));
    }
    for (field, attribute) in [("bold", "b"), ("italic", "i"), ("strike", "strike")] {
        if let Some(value) = get(style, field).as_bool() { run.push_str(&format!("<w:{attribute} w:val=\"{}\"/>", u8::from(value))); }
    }
    if let Some(color) = get(level, "bulletColor").as_str().or_else(|| get(style, "color").as_str()).and_then(|value| text_color(value.to_owned())) { run.push_str(&format!("<w:color w:val=\"{color}\"/>")); }
    if let Some(size) = get(style, "fontSizePt").as_f64().filter(|value| value.is_finite() && *value > 0.0) {
        let size = (size * 2.0).round().clamp(2.0, 3276.0) as u32;
        run.push_str(&format!("<w:sz w:val=\"{size}\"/><w:szCs w:val=\"{size}\"/>"));
    }
    if get(style, "underline").as_bool() == Some(true) { run.push_str("<w:u w:val=\"single\"/>"); }
    if !run.is_empty() { properties.push_str(&format!("<w:rPr>{run}</w:rPr>")); }
    format!("<w:lvl w:ilvl=\"{index}\">{properties}</w:lvl>")
}
