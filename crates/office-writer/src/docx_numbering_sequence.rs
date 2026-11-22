use std::collections::{HashMap, HashSet};

use serde_json::{Value, json};

use super::{NumberingInstance, abstract_definition, abstract_key, array, cell_blocks, get, integer, numbering_level, text_color};
use crate::{ExportError, ExportWarning, xml};
use crate::docx::numbering_marker::PreservedMarkerPlan;

const MAPPING_NAMESPACE: &str = "urn:office-studio:numbering:1";
const BOOKMARK_PREFIX: &str = "office_list_v1_";

#[derive(Clone)]
pub(crate) struct SequenceParagraph {
    pub prefix: String,
    pub sequence: bool,
    pub native_marker: bool,
    pub left: f64,
    pub right: f64,
    pub first_line: f64,
    pub hanging: f64,
    pub tab: bool,
}

#[derive(Default)]
pub(crate) struct SequencePlan {
    pub paragraphs: HashMap<usize, SequenceParagraph>,
    pub mapping_xml: Option<String>,
    pub next_bookmark: u32,
    pub sequence_prefix: String,
    pub bookmark_names: HashSet<String>,
    pub warnings: Vec<ExportWarning>,
}

#[derive(Clone)]
enum InstructionPart { Text(String), Field(Field) }

#[derive(Clone)]
struct Field { parts: Vec<InstructionPart>, result: String }

impl Field {
    fn plain(instruction: impl Into<String>, result: impl Into<String>) -> Self {
        Self { parts: vec![InstructionPart::Text(instruction.into())], result: result.into() }
    }

    fn code(&self) -> String {
        let code = self.parts.iter().map(|part| match part {
            InstructionPart::Text(text) => text.clone(),
            InstructionPart::Field(field) => format!("{{{}}}", field.code()),
        }).collect::<String>();
        normalize_instruction(&code)
    }

    fn render(&self, properties: &str) -> String {
        let mut output = format!("<w:r>{properties}<w:fldChar w:fldCharType=\"begin\" w:dirty=\"true\"/></w:r>");
        for part in &self.parts {
            match part {
                InstructionPart::Text(text) => output.push_str(&format!("<w:r>{properties}<w:instrText xml:space=\"preserve\">{}</w:instrText></w:r>", xml::escape(text))),
                InstructionPart::Field(field) => output.push_str(&field.render(properties)),
            }
        }
        output.push_str(&format!("<w:r>{properties}<w:fldChar w:fldCharType=\"separate\"/></w:r>"));
        if !self.result.is_empty() { output.push_str(&format!("<w:r>{properties}<w:t xml:space=\"preserve\">{}</w:t></w:r>", xml::escape(&self.result))); }
        output.push_str(&format!("<w:r>{properties}<w:fldChar w:fldCharType=\"end\"/></w:r>"));
        output
    }
}

fn normalize_instruction(code: &str) -> String {
    let mut output = String::new();
    let mut quoted = false;
    let mut escaped = false;
    let mut space = false;
    for character in code.trim().chars() {
        if !quoted && character.is_whitespace() { space = true; continue; }
        if space && !output.is_empty() { output.push(' '); }
        space = false;
        output.push(character);
        if character == '"' && !escaped { quoted = !quoted; }
        escaped = character == '\\' && !escaped;
    }
    output
}

fn text_part(text: impl Into<String>) -> InstructionPart { InstructionPart::Text(text.into()) }
fn nested(field: Field) -> InstructionPart { InstructionPart::Field(field) }
fn current(name: &str, value: u64) -> Field { Field::plain(format!(" SEQ {name} \\c "), value.to_string()) }
fn expression(parts: Vec<InstructionPart>, result: i64) -> Field {
    let mut code = vec![text_part(" = ")];
    code.extend(parts);
    Field { parts: code, result: result.to_string() }
}

fn quoted(text: &str) -> String { text.replace('\\', "\\\\").replace('"', "\\\"") }

fn positive_result(name: &str, value: u64, body: Vec<InstructionPart>, result: String) -> Field {
    let mut parts = vec![text_part(" IF "), nested(current(name, value)), text_part(" > 0 \"")];
    parts.extend(body);
    parts.push(text_part("\" \"0\" "));
    Field { parts, result: if value == 0 { "0".into() } else { result } }
}

enum CustomToken { Decimal { zero: u32, width: usize }, Alphabet(Vec<char>), Roman(bool) }
struct CustomFormat { prefix: String, suffix: String, token: CustomToken }

fn custom_format(pattern: &str) -> Option<CustomFormat> {
    let characters = pattern.chars().collect::<Vec<_>>();
    let start = characters.iter().position(|character| character.is_alphanumeric())?;
    let end = characters.iter().enumerate().skip(start).find(|(_, character)| !character.is_alphanumeric()).map_or(characters.len(), |(index, _)| index);
    let token = characters[start..end].iter().collect::<String>();
    let prefix = characters[..start].iter().collect();
    let suffix_start = characters.iter().rposition(|character| character.is_alphanumeric())? + 1;
    let suffix = characters[suffix_start..].iter().collect();
    let alphabet = match token.as_str() {
        "A" => Some("ABCDEFGHIJKLMNOPQRSTUVWXYZ"), "a" => Some("abcdefghijklmnopqrstuvwxyz"),
        "Α" => Some("ΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡΣΤΥΦΧΨΩ"), "α" => Some("αβγδεζηθικλμνξοπρστυφχψω"),
        "А" => Some("АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯ"), "а" => Some("абвгдеёжзийклмнопрстуфхцчшщъыьэюя"),
        "ア" => Some("アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワヲン"),
        "あ" => Some("あいうえおかきくけこさしすせそたちつてとなにぬねのはひふへほまみむめもやゆよらりるれろわをん"), _ => None,
    };
    let token = if let Some(alphabet) = alphabet { CustomToken::Alphabet(alphabet.chars().collect()) }
    else if token == "I" || token == "i" { CustomToken::Roman(token == "I") }
    else {
        const ZEROS: &[u32] = &[
            0x30, 0x660, 0x6f0, 0x7c0, 0x966, 0x9e6, 0xa66, 0xae6, 0xb66, 0xbe6,
            0xc66, 0xce6, 0xd66, 0xde6, 0xe50, 0xed0, 0xf20, 0x1040, 0x1090, 0x17e0,
            0x1810, 0x1946, 0x19d0, 0x1a80, 0x1a90, 0x1b50, 0x1bb0, 0x1c40, 0x1c50,
            0xa620, 0xa8d0, 0xa900, 0xa9d0, 0xa9f0, 0xaa50, 0xabf0, 0xff10, 0x104a0,
            0x10d30, 0x11066, 0x110f0, 0x11136, 0x111d0, 0x112f0, 0x11450, 0x114d0,
            0x11650, 0x116c0, 0x11730, 0x118e0, 0x11950, 0x11c50, 0x11d50, 0x11da0,
            0x16a60, 0x16ac0, 0x16b50, 0x1d7ce, 0x1d7d8, 0x1d7e2, 0x1d7ec, 0x1d7f6,
            0x1e140, 0x1e2f0, 0x1e950, 0x1fbf0,
        ];
        let digits = token.chars().collect::<Vec<_>>();
        let last = *digits.last()? as u32;
        let zero = ZEROS.iter().copied().find(|zero| last == zero + 1 && digits[..digits.len() - 1].iter().all(|glyph| *glyph as u32 == *zero))?;
        if digits.len() > 255 { return None; }
        CustomToken::Decimal { zero, width: digits.len() }
    };
    Some(CustomFormat { prefix, suffix, token })
}

fn roman(mut value: u64, uppercase: bool) -> String {
    if value == 0 || value > 3_999_999 { return value.to_string(); }
    let mut result = String::new();
    for (amount, glyph) in [(1000, "m"), (900, "cm"), (500, "d"), (400, "cd"), (100, "c"), (90, "xc"), (50, "l"), (40, "xl"), (10, "x"), (9, "ix"), (5, "v"), (4, "iv"), (1, "i")] {
        while value >= amount { result.push_str(glyph); value -= amount; }
    }
    if uppercase { result.to_ascii_uppercase() } else { result }
}

fn custom_text(value: u64, format: &CustomFormat) -> String {
    if value == 0 { return "0".into(); }
    let core = match &format.token {
        CustomToken::Decimal { zero, width } => format!("{value:0width$}", width = *width).chars().map(|digit| char::from_u32(zero + digit.to_digit(10).unwrap_or(0)).unwrap_or(digit)).collect::<String>(),
        CustomToken::Roman(uppercase) => roman(value, *uppercase),
        CustomToken::Alphabet(alphabet) => {
            let mut digits = Vec::new();
            let mut remaining = value;
            while remaining > 0 { remaining -= 1; digits.push(alphabet[remaining as usize % alphabet.len()]); remaining /= alphabet.len() as u64; }
            digits.iter().rev().collect()
        }
    };
    format!("{}{core}{}", format.prefix, format.suffix)
}

fn lookup(parts: Vec<InstructionPart>, value: usize, glyphs: &[char], offset: usize) -> Field {
    let glyph = glyphs[offset].to_string();
    let mut code = vec![text_part(" IF ")];
    code.extend(parts.clone());
    code.push(text_part(format!(" = {} \"{}\" \"", offset + 1, quoted(&glyph))));
    if offset + 1 < glyphs.len() { code.push(nested(lookup(parts, value, glyphs, offset + 1))); }
    code.push(text_part("\" "));
    Field { parts: code, result: glyphs.get(value.saturating_sub(1)).map(char::to_string).unwrap_or_default() }
}

fn decimal_field(name: &str, value: u64, width: usize, zero: u32) -> Field {
    if zero == 0x30 {
        let field = Field::plain(format!(" SEQ {name} \\c \\# \"{}\" ", "0".repeat(width.max(1))), format!("{value:0width$}", width = width.max(1)));
        return positive_result(name, value, vec![nested(field.clone())], field.result);
    }
    let glyphs = (0..10).filter_map(|digit| char::from_u32(zero + digit)).collect::<Vec<_>>();
    let mut body = Vec::new();
    let maximum_digits = 16usize;
    if width > maximum_digits { body.push(text_part(glyphs[0].to_string().repeat(width - maximum_digits))); }
    for place in (0..maximum_digits).rev() {
        let power = 10u64.pow(place as u32);
        let digit = (value / power % 10) as usize;
        let operand = expression(vec![text_part("MOD(INT("), nested(current(name, value)), text_part(format!("/{power}),10)+1"))], digit as i64 + 1);
        let rendered = lookup(vec![nested(operand)], digit + 1, &glyphs, 0);
        if place < width { body.push(nested(rendered)); }
        else {
            body.push(nested(Field { parts: vec![text_part(" IF "), nested(current(name, value)), text_part(format!(" >= {power} \"")), nested(rendered), text_part("\" \"\" ")], result: if value >= power { glyphs[digit].to_string() } else { String::new() } }));
        }
    }
    let result = format!("{value:0width$}", width = width.max(1)).chars().map(|digit| glyphs[digit.to_digit(10).unwrap_or(0) as usize]).collect();
    positive_result(name, value, body, result)
}

fn alphabetic_field(name: &str, value: u64, alphabet: &[char]) -> Field {
    let base = alphabet.len() as u64;
    let mut body = Vec::new();
    let mut power = 1u64;
    let mut positions = Vec::new();
    while power <= 9_007_199_254_740_991 {
        positions.push(power);
        let Some(next) = power.checked_mul(base) else { break; };
        power = next;
    }
    for power in positions.into_iter().rev() {
        let offset = (power - 1) / (base - 1);
        let quotient = value.saturating_sub(offset) / power;
        let quotient_parts = vec![text_part("MAX(0,INT(("), nested(current(name, value)), text_part(format!("-{offset})/{power}))"))];
        let letter = quotient.saturating_sub(1) as usize % alphabet.len();
        let letter_parts = vec![text_part("MOD("), nested(expression(quotient_parts.clone(), quotient as i64)), text_part(format!("-1,{base})+1"))];
        let glyph = if alphabet.first() == Some(&'A') || alphabet.first() == Some(&'a') {
            let mut parts = vec![text_part(" = ")];
            parts.extend(letter_parts);
            parts.push(text_part(if alphabet[0] == 'A' { " \\* ALPHABETIC " } else { " \\* alphabetic " }));
            Field { parts, result: alphabet[letter].to_string() }
        } else { lookup(vec![nested(expression(letter_parts, letter as i64 + 1))], letter + 1, alphabet, 0) };
        body.push(nested(Field { parts: vec![text_part(" IF "), nested(expression(quotient_parts, quotient as i64)), text_part(" > 0 \""), nested(glyph), text_part("\" \"\" ")], result: if quotient > 0 { alphabet[letter].to_string() } else { String::new() } }));
    }
    let format = CustomFormat { prefix: String::new(), suffix: String::new(), token: CustomToken::Alphabet(alphabet.to_vec()) };
    positive_result(name, value, body, custom_text(value, &format))
}

fn ordinal_field(name: &str, value: u64, level: &Value, legal: bool, warnings: &mut HashSet<String>) -> Field {
    if !legal {
        if let Some(pattern) = get(level, "customFormat").as_str() {
            if let Some(format) = custom_format(pattern) {
                let core = match &format.token {
                    CustomToken::Decimal { zero, width } => decimal_field(name, value, *width, *zero),
                    CustomToken::Alphabet(alphabet) => alphabetic_field(name, value, alphabet),
                    CustomToken::Roman(uppercase) => Field::plain(format!(" SEQ {name} \\c \\* {} ", if *uppercase { "ROMAN" } else { "roman" }), roman(value, *uppercase)),
                };
                return positive_result(name, value, vec![text_part(quoted(&format.prefix)), nested(core), text_part(quoted(&format.suffix))], custom_text(value, &format));
            }
            warnings.insert(format!("自定义编号格式 {pattern:?} 未定义可移植的序列域表示；保留原始格式并使用其 OOXML 回退格式。"));
        }
    }
    let format = if legal { "decimal" } else { get(level, "format").as_str().unwrap_or("decimal") };
    let width = if legal { 0 } else { integer(get(level, "minimumDigits")).unwrap_or(0).clamp(0, 255) as usize };
    if matches!(format, "decimal" | "decimalZero") {
        let width = width.max(if format == "decimalZero" { 2 } else { 1 });
        return decimal_field(name, value, width, 0x30);
    }
    let (switch, result) = match format {
        "lowerRoman" => ("roman", roman(value, false)), "upperRoman" => ("ROMAN", roman(value, true)),
        "lowerLetter" | "upperLetter" if value > 0 && value.div_ceil(26) <= 4096 => {
            let letter = char::from_u32((if format == "lowerLetter" { 97 } else { 65 }) + ((value - 1) % 26) as u32).unwrap_or('A');
            (if format == "lowerLetter" { "alphabetic" } else { "ALPHABETIC" }, letter.to_string().repeat(value.div_ceil(26) as usize))
        }
        "ordinal" => ("ordinal", format!("{value}{}", if (11..=13).contains(&(value % 100)) { "th" } else { match value % 10 { 1 => "st", 2 => "nd", 3 => "rd", _ => "th" } })),
        "hex" => ("Hex", format!("{value:X}")),
        _ => {
            warnings.insert(format!("编号格式 {format:?} 的序列域使用十进制回退，原始格式仍保存在可编辑编号定义中。"));
            ("Arabic", value.to_string())
        }
    };
    let field = Field::plain(format!(" SEQ {name} \\c \\* {switch} "), result.clone());
    if value == 0 && matches!(format, "lowerLetter" | "upperLetter" | "lowerRoman" | "upperRoman") { positive_result(name, value, vec![nested(field)], "0".into()) }
    else { field }
}

#[derive(Default)]
struct CounterState { total: Option<u64>, levels: [Option<u64>; 9], initialized: HashMap<u32, HashSet<usize>> }

fn effective_level(definition: &Value, index: usize) -> Value {
    let base = abstract_definition(definition);
    let mut level = array(get(&base, "levels")).iter().find(|level| integer(get(level, "ilvl")) == Some(index as i64)).cloned().unwrap_or_else(|| json!({ "ilvl": index }));
    if let Some(overrides) = array(get(get(definition, "instance"), "levelOverrides")).iter().find(|level| integer(get(level, "ilvl")) == Some(index as i64)).and_then(Value::as_object) {
        for (key, value) in overrides { level[key] = value.clone(); }
    }
    level
}

fn start(definition: &Value, index: usize, first: bool) -> u64 {
    let explicit = first.then(|| integer(get(get(get(definition, "instance"), "levelStartOverrides"), &index.to_string()))).flatten();
    explicit.or_else(|| integer(get(&effective_level(definition, index), "start"))).filter(|value| *value >= 0).unwrap_or(1) as u64
}

fn primary(instance: &NumberingInstance, group: &Value) -> bool {
    if group.is_null() { return true; }
    if let Some(id) = integer(group) { return instance.definition.as_ref().and_then(|definition| integer(get(get(definition, "instance"), "numId"))) == Some(id) || instance.name.as_deref() == Some(format!("num-{id}").as_str()); }
    group.as_str().is_some_and(|group| instance.name.as_deref() == Some(group))
}

fn group_key(instance: &NumberingInstance) -> (String, String) {
    let definition = instance.definition.as_ref().expect("a sequence instance has a definition");
    let group = get(get(definition, "instance"), "consecutiveListId");
    let name = if let Some(id) = integer(group) { format!("source:{id}") }
        else if let Some(group) = group.as_str() { format!("name:{group}") }
        else { format!("instance:{}", instance.id) };
    (abstract_key(instance.kind, Some(definition)), name)
}

fn marker_style(block: &Value, level: &Value) -> (String, Value) {
    let run = array(get(block, "runs")).first().unwrap_or(&Value::Null);
    let style = get(level, "runStyle");
    let font = get(level, "bulletFontFamily").as_str().or_else(|| get(style, "fontFamilyAscii").as_str()).or_else(|| get(style, "resolvedFontFamilyAscii").as_str()).or_else(|| get(style, "fontFamily").as_str()).or_else(|| get(run, "fontFamily").as_str()).unwrap_or("Arial");
    let east_asia = get(style, "fontFamilyEastAsia").as_str().or_else(|| get(style, "resolvedFontFamilyEastAsia").as_str()).unwrap_or(font);
    let color = get(level, "bulletColor").as_str().or_else(|| get(style, "color").as_str()).or_else(|| get(run, "color").as_str()).and_then(|color| text_color(color.to_owned())).unwrap_or_else(|| "1F2937".into());
    let size = get(style, "fontSizePt").as_f64().or_else(|| get(run, "fontSize").as_f64()).filter(|size| size.is_finite() && *size > 0.0).unwrap_or(11.0);
    let size = (size * 2.0).round().clamp(2.0, 3276.0) as u32;
    let flag = |name: &str| get(style, name).as_bool().or_else(|| get(run, name).as_bool()).unwrap_or(false);
    let signature = json!({ "font": font, "eastAsia": east_asia, "color": color.to_ascii_uppercase(), "size": size, "bold": flag("bold"), "italic": flag("italic"), "strike": flag("strike"), "underline": flag("underline") });
    let underline = if flag("underline") { "<w:u w:val=\"single\"/>" } else { "" };
    let properties = format!("<w:rPr><w:rFonts w:ascii=\"{}\" w:hAnsi=\"{}\" w:eastAsia=\"{}\"/><w:b w:val=\"{}\"/><w:i w:val=\"{}\"/><w:strike w:val=\"{}\"/><w:color w:val=\"{}\"/><w:sz w:val=\"{size}\"/><w:szCs w:val=\"{size}\"/>{underline}</w:rPr>", xml::escape(font), xml::escape(font), xml::escape(east_asia), u8::from(flag("bold")), u8::from(flag("italic")), u8::from(flag("strike")), xml::escape(&color));
    (properties, signature)
}

fn push_literal(output: &mut String, contract: &mut Vec<Value>, text: &str, properties: &str) {
    if text.is_empty() { return; }
    output.push_str(&format!("<w:r>{properties}<w:t xml:space=\"preserve\">{}</w:t></w:r>", xml::escape(text)));
    if let Some(last) = contract.last_mut().filter(|last| last["text"].is_string()) {
        let joined = format!("{}{text}", last["text"].as_str().unwrap_or_default());
        last["text"] = json!(joined);
    } else { contract.push(json!({ "text": text })); }
}

fn native_xml_contract(numbering_xml: &str, id: u32) -> Option<String> {
    let start = numbering_xml.find(&format!("<w:num w:numId=\"{id}\""))?;
    let end = start + numbering_xml[start..].find("</w:num>")? + "</w:num>".len();
    let number = &numbering_xml[start..end];
    let key = "<w:abstractNumId w:val=\"";
    let abstract_start = number.find(key)? + key.len();
    let abstract_end = abstract_start + number[abstract_start..].find('"')?;
    let abstract_id = &number[abstract_start..abstract_end];
    let start = numbering_xml.find(&format!("<w:abstractNum w:abstractNumId=\"{abstract_id}\""))?;
    let end = start + numbering_xml[start..].find("</w:abstractNum>")? + "</w:abstractNum>".len();
    Some(format!("<nativeContract>{}{number}</nativeContract>", &numbering_xml[start..end]))
}

pub(crate) fn prepare_sequences(blocks: &[Value], lookup: &HashMap<usize, u32>, instances: &[NumberingInstance], numbering_xml: &str, preserved: &PreservedMarkerPlan) -> Result<SequencePlan, ExportError> {
    let mut plan = SequencePlan { next_bookmark: preserved.next_bookmark, sequence_prefix: preserved.sequence_prefix.clone(), bookmark_names: preserved.bookmark_names.clone(), ..Default::default() };
    let by_id = instances.iter().map(|instance| (instance.id, instance)).collect::<HashMap<_, _>>();
    let mut groups = HashMap::<(String, String), u32>::new();
    let mut next_group = instances.len() as u32 + 1;
    for instance in instances {
        let Some(definition) = &instance.definition else { continue; };
        if get(get(definition, "abstract"), "consecutiveNumbering").as_bool() != Some(true) { continue; }
        let key = group_key(instance);
        if groups.contains_key(&key) { continue; }
        let group = get(get(definition, "instance"), "consecutiveListId");
        let primary_id = if group.is_null() { Some(instance.id) } else { instances.iter().find(|candidate| candidate.definition.is_some() && abstract_key(candidate.kind, candidate.definition.as_ref()) == key.0 && primary(candidate, group)).map(|candidate| candidate.id) };
        let id = primary_id.unwrap_or_else(|| { let id = next_group; next_group += 1; id });
        groups.insert(key, id);
    }
    let mut mapped_instances = Vec::new();
    for instance in instances {
        let Some(definition) = &instance.definition else { continue; };
        if get(get(definition, "abstract"), "consecutiveNumbering").as_bool() != Some(true) && get(get(definition, "instance"), "consecutiveListId").is_null() { continue; }
        mapped_instances.push(json!({ "numId": instance.id, "counterId": groups.get(&group_key(instance)).copied().unwrap_or(instance.id), "listId": instance.name, "definition": definition, "nativeXml": native_xml_contract(numbering_xml, instance.id) }));
    }
    let mut mappings = Vec::new();
    let mut counters = HashMap::<u32, CounterState>::new();
    let mut warnings = HashSet::new();
    fn visit(blocks: &[Value], lookup: &HashMap<usize, u32>, instances: &HashMap<u32, &NumberingInstance>, groups: &HashMap<(String, String), u32>, counters: &mut HashMap<u32, CounterState>, mappings: &mut Vec<Value>, warnings: &mut HashSet<String>, plan: &mut SequencePlan, preserved: &PreservedMarkerPlan) -> Result<(), ExportError> {
        for block in blocks {
            if get(block, "type").as_str() == Some("table") {
                for row in array(get(get(get(block, "metadata"), "table"), "rows")) { for cell in array(get(row, "cells")) { visit(cell_blocks(cell), lookup, instances, groups, counters, mappings, warnings, plan, preserved)?; } }
                continue;
            }
            let pointer = block as *const Value as usize;
            if preserved.paragraphs.get(&pointer).is_some_and(|paragraph| paragraph.replays()) { continue; }
            let Some(id) = lookup.get(&pointer).copied() else { continue; };
            let Some(instance) = instances.get(&id).copied() else { continue; };
            let Some(definition) = &instance.definition else { continue; };
            let sequence = get(get(definition, "abstract"), "consecutiveNumbering").as_bool() == Some(true);
            if !sequence && get(get(definition, "instance"), "consecutiveListId").is_null() { continue; }
            let index = numbering_level(block) as usize;
            let level = effective_level(definition, index);
            let native_marker = !sequence || matches!(get(&level, "format").as_str(), Some("bullet" | "none")) || get(&level, "pictureBullet").is_object();
            let bookmark_id = plan.next_bookmark;
            plan.next_bookmark = plan.next_bookmark.checked_add(1).ok_or_else(|| ExportError::invalid("Too many numbered paragraphs."))?;
            let mut hash = 2_166_136_261u32;
            for byte in get(block, "id").as_str().unwrap_or_default().bytes().chain(bookmark_id.to_le_bytes()) { hash = (hash ^ u32::from(byte)).wrapping_mul(16_777_619); }
            let mut bookmark = format!("{BOOKMARK_PREFIX}{id:08X}{index:X}{hash:08X}");
            if !plan.bookmark_names.insert(bookmark.to_ascii_lowercase()) {
                bookmark = format!("{BOOKMARK_PREFIX}new_{id:08X}_{bookmark_id:08X}");
                if !plan.bookmark_names.insert(bookmark.to_ascii_lowercase()) { return Err(ExportError::invalid("编号域书签名称重复。")); }
            }
            let mut prefix = format!("<w:bookmarkStart w:id=\"{bookmark_id}\" w:name=\"{bookmark}\"/>");
            let mut contract = Vec::new();
            let (properties, style) = marker_style(block, &level);
            if sequence {
                let counter_id = groups[&group_key(instance)];
                let state = counters.entry(counter_id).or_default();
                let initialized = state.initialized.entry(id).or_default();
                let first = initialized.is_empty();
                let original_group = get(get(definition, "instance"), "consecutiveListId");
                let restart = (first && !primary(instance, original_group)).then(|| (0..9).find_map(|slot| integer(get(get(get(definition, "instance"), "levelStartOverrides"), &slot.to_string())).filter(|start| *start >= 0))).flatten().map(|start| start as u64);
                let initial = integer(get(get(get(definition, "instance"), "levelStartOverrides"), &index.to_string())).filter(|start| *start >= 0).map(|start| start as u64);
                let value = restart.or_else(|| state.total.map(|value| value.saturating_add(1))).unwrap_or_else(|| initial.unwrap_or_else(|| start(definition, 0, true)));
                let name = format!("{}{counter_id}", plan.sequence_prefix);
                let reset = if restart.is_some() || state.total.is_none() { format!(" \\r {value}") } else { String::new() };
                let total = Field::plain(format!(" SEQ {name}{reset} \\h "), "");
                prefix.push_str(&total.render(&properties));
                contract.push(json!({ "field": total.code() }));
                for slot in 0..index {
                    if state.levels[slot].is_none() {
                        let value = start(definition, slot, !initialized.contains(&slot));
                        state.levels[slot] = Some(value);
                        let field = Field::plain(format!(" SEQ {name}_l{slot} \\r {value} \\h "), "");
                        prefix.push_str(&field.render(&properties));
                        contract.push(json!({ "field": field.code() }));
                    }
                }
                let setter = Field { parts: vec![text_part(format!(" SEQ {name}_l{index} \\r ")), nested(current(&name, value)), text_part(" \\h ")], result: String::new() };
                prefix.push_str(&setter.render(&properties));
                contract.push(json!({ "field": setter.code() }));
                state.total = Some(value);
                state.levels[index] = Some(value);
                initialized.insert(index);
                state.levels[index + 1..].fill(None);
                if !native_marker {
                    let pattern = get(&level, "text").as_str().unwrap_or_default();
                    let mut referenced = HashSet::new();
                    let mut characters = pattern.chars().peekable();
                    while let Some(character) = characters.next() {
                        if character == '%' && characters.peek().is_some_and(|character| ('1'..='9').contains(character)) { referenced.insert(characters.next().unwrap() as usize - '1' as usize); }
                    }
                    let mut referenced = referenced.into_iter().collect::<Vec<_>>();
                    referenced.sort_unstable();
                    for slot in referenced.into_iter().filter(|slot| *slot > index) {
                        let field = Field::plain(format!(" SEQ {name}_l{slot} \\r {} \\h ", start(definition, slot, true)), "");
                        prefix.push_str(&field.render(&properties));
                        contract.push(json!({ "field": field.code() }));
                    }
                    let legal = get(&level, "isLgl").as_bool() == Some(true);
                    let mut literal = String::new();
                    let mut characters = pattern.chars().peekable();
                    while let Some(character) = characters.next() {
                        if character != '%' || characters.peek().is_none_or(|character| !('1'..='9').contains(character)) { literal.push(character); continue; }
                        push_literal(&mut prefix, &mut contract, &literal, &properties);
                        literal.clear();
                        let slot = characters.next().unwrap() as usize - '1' as usize;
                        let referenced_level = effective_level(definition, slot);
                        let ordinal = state.levels[slot].unwrap_or_else(|| start(definition, slot, true));
                        let field = ordinal_field(&format!("{name}_l{slot}"), ordinal, &referenced_level, legal, warnings);
                        prefix.push_str(&field.render(&properties));
                        contract.push(json!({ "field": field.code() }));
                    }
                    push_literal(&mut prefix, &mut contract, &literal, &properties);
                    match get(&level, "suffix").as_str().unwrap_or("tab") {
                        "nothing" => {}, "space" => push_literal(&mut prefix, &mut contract, " ", &properties),
                        _ => { prefix.push_str(&format!("<w:r>{properties}<w:tab/></w:r>")); if let Some(last) = contract.last_mut().filter(|last| last["text"].is_string()) { last["text"] = json!(format!("{}\t", last["text"].as_str().unwrap_or_default())); } else { contract.push(json!({ "text": "\t" })); } },
                    }
                }
            }
            prefix.push_str(&format!("<w:bookmarkEnd w:id=\"{bookmark_id}\"/>"));
            let mode = if !sequence { "native" } else if native_marker { "sequence-native-marker" } else { "sequence" };
            mappings.push(json!({ "bookmark": bookmark, "numId": id, "ilvl": index, "mode": mode, "contract": contract, "markerStyle": if sequence { style } else { Value::Null } }));
            let indent = get(&level, "indent");
            let dimension = |name: &str, fallback: f64| get(indent, name).as_f64().unwrap_or(fallback) / 15.0;
            plan.paragraphs.insert(pointer, SequenceParagraph { prefix, sequence, native_marker, left: dimension("leftTwips", 420.0 * (index as f64 + 1.0)), right: dimension("rightTwips", 0.0), first_line: dimension("firstLineTwips", 0.0), hanging: dimension("hangingTwips", 420.0), tab: get(&level, "suffix").as_str().unwrap_or("tab") == "tab" });
        }
        Ok(())
    }
    visit(blocks, lookup, &by_id, &groups, &mut counters, &mut mappings, &mut warnings, &mut plan, preserved)?;
    if !mappings.is_empty() {
        let mapping = json!({ "version": 1, "instances": mapped_instances, "paragraphs": mappings });
        plan.mapping_xml = Some(format!("{}<officeNumbering xmlns=\"{MAPPING_NAMESPACE}\" version=\"1\"><mapping>{}</mapping></officeNumbering>", xml::DECL, xml::escape(&mapping.to_string())));
    }
    let mut warnings = warnings.into_iter().collect::<Vec<_>>();
    warnings.sort_unstable();
    plan.warnings = warnings.into_iter().map(|warning| ExportWarning::new("DOCX_SEQUENCE_FORMAT_FALLBACK", warning)).collect();
    Ok(plan)
}
