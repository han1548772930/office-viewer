//! Native calculation views for imported formulas with retained date calendars.
//! References keep their geometry in sparse hidden worksheets. The writer is a
//! projection, not another formula evaluator; Excel computes the emitted formulas.

use super::{WorkbookInput, SheetInput, table, cell_from_entry, column_name};
use crate::{ExportError, ExportWarning};
use office_io_common::sheet_formula_calendar_source::{CalculationFingerprint, CellSource, HelperSource, SheetSource, TableSource, TableColumnSource, FormulaCalendarSource};
use serde_json::{json, Value};
use std::collections::{BTreeMap, BTreeSet};

#[path = "xlsx_formula_calendar_functions.rs"]
mod functions;
#[path = "xlsx_formula_builtin_names.rs"]
mod builtin_names;

const MAX_BYTES: usize = 64 * 1024 * 1024;
const MAX_CELLS: usize = 250_000;
const MAX_FORMULA: usize = 32_768;
type LocalBindings = BTreeMap<String, String>;

#[derive(Clone)]
struct Mirror { id: String, name: String }
#[derive(Clone)]
struct Name { name: String, formula: String, owner: usize, scope: Option<usize> }
struct Projection<'a> {
    workbook: &'a WorkbookInput,
    frames: BTreeMap<bool, Vec<Mirror>>,
    names: Vec<Name>,
    global_names: BTreeMap<(bool, String), String>,
    tables: BTreeMap<(bool, String), String>,
    alias_sheet: Mirror,
    aliases: BTreeMap<(bool, usize, Option<usize>), (usize, usize)>,
    alias_cells: serde_json::Map<String, Value>,
    variable_prefix: String,
    calls: usize,
    bytes: usize,
    cells: usize,
}
fn number(value: &Value) -> Option<f64> {
    value.as_f64().or_else(|| value.as_str().filter(|text| !text.trim().is_empty()).and_then(|text| text.parse::<f64>().ok())).filter(|value| value.is_finite())
}
fn formula_source(value: &Value) -> Option<&Value> {
    let marker = value.get("metadata")?.get("chartDependencyFormulaSource")?;
    (marker.get("version").and_then(Value::as_u64) == Some(1)
        && marker.get("formula").and_then(Value::as_str) == value.get("formula").and_then(Value::as_str)
        && marker.get("formula").and_then(Value::as_str).is_some()
        && marker.get("date1904").and_then(Value::as_bool).is_some()
        && marker.get("dateResult").and_then(Value::as_bool).is_some()
        && marker.get("aliasSheetId").is_none_or(|value| value.as_str().is_some_and(|value| !value.is_empty() && value.len() <= 512))).then_some(marker)
}
fn calendar_serial(value: f64, before: bool, after: bool) -> f64 {
    office_io_common::excel_date::convert_serial(value, before, after)
}
fn date_source(value: &Value, date1904: bool) -> Option<&Value> {
    let marker = value.get("metadata")?.get("chartDependencyDateSource")?;
    let original = number(marker.get("value")?)?;
    let source = marker.get("date1904")?.as_bool()?;
    if marker.get("targetDate1904").is_some_and(|target| target.as_bool() != Some(date1904))
        || marker.get("targetFormula").is_some_and(|formula| Some(formula) != value.get("formula")) { return None; }
    let target = marker.get("targetValue").and_then(number).unwrap_or_else(|| calendar_serial(original, source, date1904));
    (value.get("value").and_then(number) == Some(target)).then_some(marker)
}
fn uses_calendar(value: &Value, date1904: bool) -> bool {
    if formula_source(value).and_then(|marker| marker.get("dateResult")).and_then(Value::as_bool) == Some(true)
        || date_source(value, date1904).is_some() { return true; }
    let data_type = if value.get("formula").and_then(Value::as_str).is_some() { value.get("cachedDataType").or_else(|| value.get("dataType")) } else { value.get("dataType") };
    if data_type.and_then(Value::as_str).is_some_and(|value| matches!(value.to_ascii_lowercase().as_str(), "date" | "d" | "datetime")) { return true; }
    // Strip quoted literals, escapes and bracket directives before recognizing
    // calendar fields. Pure hours/minutes/seconds and elapsed formats are numbers.
    let format = value.get("style").and_then(|style| style.get("numberFormat")).and_then(Value::as_str).unwrap_or("");
    let mut quote = false; let mut bracket = false; let mut escaped = false; let mut fields = String::new();
    for character in format.chars() {
        if escaped { escaped = false; continue; }
        if character == '\\' || character == '_' || character == '*' { escaped = true; continue; }
        if character == '"' { quote = !quote; continue; }
        if quote { continue; }
        if character == '[' { bracket = true; continue; }
        if character == ']' { bracket = false; continue; }
        if !bracket && matches!(character, 'y' | 'Y' | 'd' | 'D') { return true; }
        if !bracket && matches!(character, 'm' | 'M' | 'h' | 'H' | 's' | 'S') { fields.push(character.to_ascii_lowercase()); }
    }
    fields.contains("mmm") || fields.contains('m') && !fields.contains(['h', 's'])
}
fn quote_sheet(name: &str) -> String { format!("'{}'!", name.replace('\'', "''")) }
fn literal(value: &str) -> String { format!("\"{}\"", value.replace('"', "\"\"")) }
fn identifier(character: char) -> bool { character.is_alphanumeric() || matches!(character, '_' | '.' | '\\') }
fn identifier_start(character: char) -> bool { character.is_alphabetic() || matches!(character, '_' | '\\') }
fn quoted_end(text: &str, start: usize, delimiter: u8) -> usize {
    let mut cursor = start + 1;
    while cursor < text.len() {
        if text.as_bytes()[cursor] == delimiter {
            cursor += 1;
            if text.as_bytes().get(cursor) == Some(&delimiter) { cursor += 1; } else { break; }
        } else { cursor += text[cursor..].chars().next().unwrap().len_utf8(); }
    }
    cursor
}
fn bracket_end(text: &str, start: usize) -> usize {
    let mut cursor = start + 1; let mut depth = 1;
    while cursor < text.len() && depth > 0 {
        let character = text[cursor..].chars().next().unwrap();
        if character == '\'' && text[cursor + 1..].chars().next().is_some_and(|value| matches!(value, '[' | ']' | '#' | '@' | '\'')) {
            cursor += 1; cursor += text[cursor..].chars().next().unwrap().len_utf8(); continue;
        }
        if character == '[' { depth += 1; } else if character == ']' { depth -= 1; }
        cursor += character.len_utf8();
    }
    cursor
}
struct CallableHead<'a> { end: usize, name: &'a str, qualified: bool }
fn identifier_end(text: &str, start: usize) -> Option<usize> {
    let first = text.get(start..)?.chars().next()?;
    if !identifier_start(first) { return None; }
    let mut end = start + first.len_utf8();
    while text[end..].chars().next().is_some_and(identifier) { end += text[end..].chars().next().unwrap().len_utf8(); }
    Some(end)
}
fn callable_head(text: &str, start: usize) -> Option<CallableHead<'_>> {
    let quoted = text.as_bytes().get(start) == Some(&b'\'');
    let first_end = if quoted { quoted_end(text, start, b'\'') } else { identifier_end(text, start)? };
    if text.as_bytes().get(first_end) == Some(&b'!') {
        let name_start = first_end + 1; let end = identifier_end(text, name_start)?;
        Some(CallableHead { end, name: &text[name_start..end], qualified: true })
    } else if quoted { None }
    else { Some(CallableHead { end: first_end, name: &text[start..first_end], qualified: false }) }
}
impl Projection<'_> {
    fn charge(&mut self, bytes: usize) -> Result<(), ExportError> {
        self.bytes = self.bytes.checked_add(bytes).filter(|bytes| *bytes <= MAX_BYTES)
            .ok_or_else(|| ExportError::invalid("FORMULA_CALENDAR_PROJECTION_LIMIT: native calculation views exceed 64 MiB"))?;
        Ok(())
    }
    fn mirror(&self, frame: bool, owner: usize) -> &Mirror { &self.frames[&frame][owner] }
    fn owner(&self, name: &str) -> Option<usize> { self.workbook.sheets.iter().position(|sheet| sheet.name.to_lowercase() == name.to_lowercase()) }
    fn scoped_name(&self, name: &str, owner: usize) -> Option<&Name> {
        self.names.iter().find(|item| item.scope == Some(owner) && item.name.to_lowercase() == name.to_lowercase())
            .or_else(|| self.names.iter().find(|item| item.scope.is_none() && item.name.to_lowercase() == name.to_lowercase()))
    }
    fn emitted_name(&self, name: &Name, frame: bool) -> &str {
        if name.scope.is_some() { &name.name } else { &self.global_names[&(frame, name.name.to_lowercase())] }
    }
    fn qualified_name(&self, text: &str, start: usize, owner: usize, frame: bool) -> Option<(String, usize)> {
        let first = text[start..].chars().next()?;
        if !identifier_start(first) { return None; }
        let mut end = start + first.len_utf8();
        while text[end..].chars().next().is_some_and(identifier) { end += text[end..].chars().next().unwrap().len_utf8(); }
        if text.as_bytes().get(end) == Some(&b'[') { return None; }
        let defined = self.scoped_name(&text[start..end], owner)?;
        Some((format!("{}{}", quote_sheet(&self.mirror(frame, defined.scope.unwrap_or(defined.owner)).name), self.emitted_name(defined, frame)), end))
    }
    fn leaf(&self, text: &str, owner: usize, frame: bool, locals: &LocalBindings) -> String {
        let spans = super::conditional_formula::references(text);
        let mut mapped = String::new(); let mut cursor = 0; let mut previous_owner = owner; let mut previous_end = 0;
        for span in &spans {
            mapped.push_str(&self.identifiers(&text[cursor..span.start], owner, frame, locals));
            let declared = span.prefix.strip_suffix('!').map(|value| if value.starts_with('\'') && value.ends_with('\'') {
                value[1..value.len() - 1].replace("''", "'") } else { value.to_owned() });
            let inherited = span.prefix.is_empty() && text[previous_end..span.start].trim() == ":";
            let actual = declared.as_deref().and_then(|name| self.owner(name)).or_else(|| span.prefix.is_empty().then_some(if inherited { previous_owner } else { owner }));
            let prefix = if let Some(actual) = actual { previous_owner = actual; quote_sheet(&self.mirror(frame, actual).name) }
                else if declared.as_ref().is_some_and(|name| name.contains(':') && !name.contains('[')) {
                    let parts = declared.as_ref().unwrap().split(':').map(|name| self.owner(name).map(|index| self.mirror(frame, index).name.clone())).collect::<Option<Vec<_>>>();
                    parts.map_or_else(|| span.prefix.clone(), |parts| quote_sheet(&parts.join(":")))
                } else { span.prefix.clone() };
            mapped.push_str(&prefix); mapped.push_str(&text[span.start + span.prefix.len()..span.end]);
            previous_end = span.end; cursor = span.end;
        }
        mapped.push_str(&self.identifiers(&text[cursor..], owner, frame, locals)); mapped
    }
    fn identifiers(&self, text: &str, owner: usize, frame: bool, locals: &LocalBindings) -> String {
        let mut output = String::new(); let mut cursor = 0;
        while cursor < text.len() {
            let start = cursor; let character = text[cursor..].chars().next().unwrap();
            if character == '"' { cursor = quoted_end(text, cursor, b'"'); output.push_str(&text[start..cursor]); continue; }
            if character == '[' { cursor = bracket_end(text, cursor); output.push_str(&text[start..cursor]); continue; }
            if character == '\'' {
                cursor = quoted_end(text, cursor, b'\'');
                if text.as_bytes().get(cursor) == Some(&b'!') {
                    let caption = text[start + 1..cursor - 1].replace("''", "'");
                    if let Some(index) = self.owner(&caption) {
                        if let Some((mapped, end)) = self.qualified_name(text, cursor + 1, index, frame) { output.push_str(&mapped); cursor = end; }
                        else { output.push_str(&quote_sheet(&self.mirror(frame, index).name)); cursor += 1; }
                        continue;
                    }
                }
                output.push_str(&text[start..cursor]); continue;
            }
            if !identifier_start(character) { cursor += character.len_utf8(); output.push(character); continue; }
            cursor += character.len_utf8();
            while text[cursor..].chars().next().is_some_and(identifier) { cursor += text[cursor..].chars().next().unwrap().len_utf8(); }
            let name = &text[start..cursor];
            if text.as_bytes().get(cursor) == Some(&b'!') {
                if let Some(index) = self.owner(name) {
                    if let Some((mapped, end)) = self.qualified_name(text, cursor + 1, index, frame) { output.push_str(&mapped); cursor = end; }
                    else { output.push_str(&quote_sheet(&self.mirror(frame, index).name)); cursor += 1; }
                    continue;
                }
            }
            if let Some(local) = locals.get(&local_identifier(name)) { output.push_str(local); }
            else if let Some(table) = self.tables.get(&(frame, name.to_lowercase())) { output.push_str(table); }
            else if text[..start].chars().next_back() == Some('!') { output.push_str(name); }
            else if let Some(name) = self.scoped_name(name, owner) {
                output.push_str(&quote_sheet(&self.mirror(frame, name.scope.unwrap_or(name.owner)).name)); output.push_str(self.emitted_name(name, frame));
            } else { output.push_str(name); }
        }
        output
    }
    fn expression(&mut self, text: &str, owner: usize, frame: bool, alias_owner: Option<usize>, depth: usize) -> Result<String, ExportError> {
        self.expression_with_locals(text, owner, frame, alias_owner, depth, &LocalBindings::new())
    }
    fn expression_with_locals(&mut self, text: &str, owner: usize, frame: bool, alias_owner: Option<usize>, depth: usize,
        locals: &LocalBindings) -> Result<String, ExportError> {
        if text.len() > MAX_FORMULA || depth > 256 { return Err(ExportError::invalid("FORMULA_CALENDAR_PROJECTION_LIMIT: formula is too long or deep")); }
        let mut output = String::new(); let mut cursor = 0; let mut leaf_start = 0;
        while cursor < text.len() {
            let character = text[cursor..].chars().next().unwrap();
            if character == '"' { cursor = quoted_end(text, cursor, b'"'); continue; }
            if character == '[' { cursor = bracket_end(text, cursor); continue; }
            let start = cursor;
            if character != '\'' && (!identifier_start(character) || text[..cursor].chars().next_back().is_some_and(identifier)) { cursor += character.len_utf8(); continue; }
            let Some(head) = callable_head(text, start) else {
                cursor = if character == '\'' { quoted_end(text, cursor, b'\'') } else { cursor + character.len_utf8() };
                continue;
            };
            cursor = head.end;
            let name = head.name; let mut open = cursor;
            while text.as_bytes().get(open).is_some_and(u8::is_ascii_whitespace) { open += 1; }
            if text.as_bytes().get(open) != Some(&b'(') { continue; }
            let mut nested = 1usize; let mut array_depth = 0usize; let mut position = open + 1; let mut arg_start = position; let mut arguments = Vec::new();
            while position < text.len() && nested > 0 {
                let current = text[position..].chars().next().unwrap();
                if current == '"' || current == '\'' { position = quoted_end(text, position, current as u8); continue; }
                if current == '[' { position = bracket_end(text, position); continue; }
                if current == '(' { nested += 1; }
                else if current == '{' { array_depth += 1; }
                else if current == '}' { array_depth = array_depth.checked_sub(1).ok_or_else(|| ExportError::invalid("FORMULA_CALENDAR_PROJECTION_INVALID: unmatched array delimiter"))?; }
                else if current == ')' {
                    nested -= 1;
                    if nested == 0 {
                        if position > arg_start || !arguments.is_empty() { arguments.push(&text[arg_start..position]); }
                        break;
                    }
                } else if current == ',' && nested == 1 && array_depth == 0 {
                    arguments.push(&text[arg_start..position]); arg_start = position + 1;
                }
                position += current.len_utf8();
            }
            if nested != 0 || array_depth != 0 { return Err(ExportError::invalid("FORMULA_CALENDAR_PROJECTION_INVALID: unclosed function call or array")); }
            output.push_str(&self.leaf(&text[leaf_start..start], owner, frame, locals));
            self.calls += 1;
            let prefix = format!("{}{}_", self.variable_prefix, self.calls);
            let normalized_name = name.to_ascii_uppercase();
            let unprefixed = normalized_name.strip_prefix("_XLFN.").unwrap_or(&normalized_name);
            let normalized = unprefixed.strip_prefix("_XLWS.").unwrap_or(unprefixed);
            let intrinsic = !head.qualified && !locals.contains_key(&local_identifier(name)) && builtin_names::contains(normalized);
            let mut scope = locals.clone();
            let mut mapped_arguments = Vec::with_capacity(arguments.len());
            for (index, argument) in arguments.iter().enumerate() {
                let binder = intrinsic && ((normalized == "LET" && index % 2 == 0 && index + 1 < arguments.len())
                    || (normalized == "LAMBDA" && index + 1 < arguments.len()));
                if binder {
                    let spelling = if normalized == "LAMBDA" { lambda_parameter(argument) } else { argument.trim().to_owned() };
                    mapped_arguments.push(spelling.clone());
                    // LET's value uses the previous scope; its new name becomes
                    // visible to subsequent bindings and the final expression.
                    if normalized == "LAMBDA" { scope.insert(local_identifier(argument.trim()), spelling); }
                } else {
                    mapped_arguments.push(self.expression_with_locals(argument, owner, frame, alias_owner, depth + 1, &scope)?);
                    if intrinsic && normalized == "LET" && index % 2 == 1 {
                        scope.insert(local_identifier(arguments[index - 1].trim()), mapped_arguments[index - 1].clone());
                    }
                }
            }
            // Fixed lexer/function service resolves a registered callable before
            // a same-name defined name. A lexical callback or explicit sheet
            // qualifier still owns its actual local/defined-name spelling.
            let authored_name = if intrinsic && matches!(normalized, "LET" | "LAMBDA") { format!("_xlfn.{normalized}") }
                else if intrinsic { name.to_owned() }
                else { self.identifiers(&text[start..head.end], owner, frame, locals) };
            let result = if intrinsic && normalized == "INDIRECT" { self.indirect(&mapped_arguments, owner, frame, alias_owner, &prefix)? }
                else if intrinsic { functions::project(normalized, &authored_name, &mapped_arguments, frame, self.workbook.date1904, &prefix) }
                else { format!("{authored_name}({})", mapped_arguments.join(",")) };
            output.push_str(&result); cursor = position + 1; leaf_start = cursor;
            if output.len() > MAX_FORMULA { return Err(ExportError::invalid("FORMULA_CALENDAR_PROJECTION_LIMIT: emitted formula exceeds 32768 bytes")); }
        }
        output.push_str(&self.leaf(&text[leaf_start..], owner, frame, locals));
        if output.len() > MAX_FORMULA { return Err(ExportError::invalid("FORMULA_CALENDAR_PROJECTION_LIMIT: emitted formula exceeds 32768 bytes")); }
        Ok(output)
    }
    fn push_alias(&mut self, mappings: &mut BTreeMap<String, String>, name: String, target: String) -> Result<(), ExportError> {
        self.charge(name.len() + target.len() + 192)?;
        if mappings.len().checked_add(self.cells / 2).is_none_or(|count| count > MAX_CELLS / 2) {
            return Err(ExportError::invalid("FORMULA_CALENDAR_PROJECTION_LIMIT: calculation aliases exceed the sparse-cell budget"));
        }
        mappings.insert(name, target); Ok(())
    }
    fn alias_lookup(&mut self, frame: bool, owner: usize, alias_owner: Option<usize>) -> Result<(usize, usize), ExportError> {
        if let Some(range) = self.aliases.get(&(frame, owner, alias_owner)) { return Ok(*range); }
        let mut mappings = BTreeMap::<String, String>::new();
        for index in 0..self.workbook.sheets.len() {
            let original = self.workbook.sheets[index].name.clone();
            let target = quote_sheet(&self.mirror(frame, index).name);
            self.push_alias(&mut mappings, format!("{original}!"), target.clone())?;
            self.push_alias(&mut mappings, quote_sheet(&original), target)?;
        }
        let names = self.names.clone();
        for name in &names {
            let scope = name.scope.unwrap_or(name.owner);
            let mapped = format!("{}{}", quote_sheet(&self.mirror(frame, scope).name), self.emitted_name(name, frame));
            let qualified_scopes = if name.scope.is_some() { vec![scope] } else { (0..self.workbook.sheets.len()).collect::<Vec<_>>() };
            for qualified in qualified_scopes {
                if self.scoped_name(&name.name, qualified).is_none_or(|candidate| candidate.scope != name.scope) { continue; }
                let caption = self.workbook.sheets[qualified].name.clone();
                self.push_alias(&mut mappings, format!("{caption}!{}", name.name), mapped.clone())?;
                self.push_alias(&mut mappings, format!("{}{}", quote_sheet(&caption), name.name), mapped.clone())?;
            }
            if self.scoped_name(&name.name, owner).is_some_and(|candidate| candidate.scope == name.scope) {
                self.push_alias(&mut mappings, name.name.clone(), mapped)?;
            }
        }
        let tables = self.tables.iter().filter(|((calendar, _), _)| *calendar == frame).map(|((_, name), mapped)| (name.clone(), mapped.clone())).collect::<Vec<_>>();
        for (name, mapped) in tables { self.push_alias(&mut mappings, name, mapped)?; }
        let aliases = alias_owner.and_then(|index| self.workbook.sheets[index].metadata.as_ref())
            .and_then(|metadata| metadata.get("chartDependencyImport")).and_then(|import| import.get("formulaAliases")).cloned();
        let aliases = aliases.as_ref();
        for (name, value) in aliases.and_then(|aliases| aliases.get("sheets")).and_then(Value::as_object).into_iter().flatten() {
            let target = value.as_str().and_then(|value| self.owner(value)).map_or("#REF!".to_owned(), |index| quote_sheet(&self.mirror(frame, index).name));
            self.push_alias(&mut mappings, format!("{name}!"), target.clone())?;
            self.push_alias(&mut mappings, quote_sheet(name), target)?;
        }
        for (name, value) in aliases.and_then(|aliases| aliases.get("names")).and_then(Value::as_object).into_iter().flatten() {
            let Some(value) = value.as_str() else { continue; };
            let explicit_scope = aliases.and_then(|aliases| aliases.get("nameSheets")).and_then(|values| values.get(name)).and_then(Value::as_str);
            let scope = explicit_scope.and_then(|name| self.owner(name)).or_else(|| explicit_scope.is_none().then_some(owner));
            let target = scope.and_then(|scope| self.scoped_name(value, scope)).map_or("#REF!".to_owned(), |defined| {
                format!("{}{}", quote_sheet(&self.mirror(frame, defined.scope.unwrap_or(defined.owner)).name), self.emitted_name(defined, frame))
            });
            self.push_alias(&mut mappings, name.clone(), target.clone())?;
            if let Some((sheet_name, caption)) = name.rsplit_once('!') { self.push_alias(&mut mappings, format!("{}{caption}", quote_sheet(sheet_name)), target)?; }
        }
        for (name, value) in aliases.and_then(|aliases| aliases.get("tables")).and_then(Value::as_object).into_iter().flatten() {
            let target = value.as_str().and_then(|value| self.tables.get(&(frame, value.to_lowercase()))).cloned().unwrap_or("#REF!".into());
            self.push_alias(&mut mappings, name.clone(), target)?;
        }
        let first = self.alias_cells.len() / 2 + 1;
        for (index, (name, target)) in mappings.into_iter().enumerate() {
            let row = first + index;
            self.cells += 2;
            if self.cells > MAX_CELLS || row > 1_048_576 { return Err(ExportError::invalid("FORMULA_CALENDAR_PROJECTION_LIMIT: calculation aliases exceed the sparse-cell budget")); }
            self.alias_cells.insert(format!("{}:0", row - 1), json!({ "row": row - 1, "col": 0, "value": name }));
            self.alias_cells.insert(format!("{}:1", row - 1), json!({ "row": row - 1, "col": 1, "value": target }));
        }
        let last = self.alias_cells.len() / 2;
        if last < first { return Err(ExportError::invalid("FORMULA_CALENDAR_PROJECTION_INVALID: empty alias catalog")); }
        self.aliases.insert((frame, owner, alias_owner), (first, last)); Ok((first, last))
    }
    fn indirect(&mut self, args: &[String], owner: usize, frame: bool, alias_owner: Option<usize>, prefix: &str) -> Result<String, ExportError> {
        if args.is_empty() || args.len() > 2 { return Ok("#VALUE!".into()); }
        let variable = format!("{prefix}text"); let bang = format!("{prefix}bang"); let bracket = format!("{prefix}bracket");
        let (first, last) = self.alias_lookup(frame, owner, alias_owner)?;
        let lookup = format!("{}$A${first}:$B${last}", quote_sheet(&self.alias_sheet.name));
        let exact = format!("VLOOKUP({variable},{lookup},2,FALSE)");
        let sheet = format!("IFERROR(VLOOKUP(LEFT({variable},{bang}),{lookup},2,FALSE)&MID({variable},{bang}+1,32767),{variable})");
        let table = format!("IFERROR(VLOOKUP(LEFT({variable},{bracket}-1),{lookup},2,FALSE)&MID({variable},{bracket},32767),{}&{variable})", literal(&quote_sheet(&self.mirror(frame, owner).name)));
        let qualified = format!("IFERROR({exact},IF({bang}>0,{sheet},IF({bracket}>0,{table},{}&{variable})))", literal(&quote_sheet(&self.mirror(frame, owner).name)));
        Ok(format!("INDIRECT(_xlfn.LET({variable},{},{bang},IFERROR(FIND(CHAR(1),SUBSTITUTE({variable},\"!\",CHAR(1),LEN({variable})-LEN(SUBSTITUTE({variable},\"!\",\"\")))),0),{bracket},IFERROR(FIND(\"[\",{variable}),0),{qualified}){})", args[0],
            args.get(1).map_or(String::new(), |value| format!(",{value}"))))
    }
}

fn local_identifier(value: &str) -> String {
    let lower = value.to_lowercase(); lower.strip_prefix("_xlpm.").unwrap_or(&lower).to_owned()
}
fn lambda_parameter(value: &str) -> String {
    let value = value.trim();
    let name = value.get(..6).filter(|prefix| prefix.eq_ignore_ascii_case("_xlpm.")).map_or(value, |_| &value[6..]);
    format!("_xlpm.{name}")
}

fn unique_name(base: &str, used: &mut BTreeSet<String>) -> String {
    let mut candidate = base.to_owned(); let mut suffix = 1usize;
    while used.contains(&candidate.to_lowercase()) { suffix += 1; candidate = format!("{base}_{suffix}"); }
    used.insert(candidate.to_lowercase()); candidate
}
fn names(workbook: &WorkbookInput) -> Result<Vec<Name>, ExportError> {
    let default_owner = workbook.active_sheet_id.as_ref().and_then(|id| workbook.sheets.iter().position(|sheet| &sheet.id == id)).unwrap_or(0);
    let safe_names = super::safe_sheet_names(workbook.sheets.iter().map(|sheet| sheet.name.as_str()));
    let mut result = Vec::<Name>::new();
    for (sheet_scope, metadata) in std::iter::once((None, workbook.metadata.as_ref())).chain(workbook.sheets.iter().enumerate().map(|(index, sheet)| (Some(index), sheet.metadata.as_ref()))) {
        for (key, local) in [("definedNames", false), ("localDefinedNames", true)] {
            let Some(values) = metadata.and_then(|value| value.get(key)).and_then(Value::as_array) else { continue; };
            for value in values {
                let Some(name) = value.get("name").and_then(Value::as_str) else { continue; };
                let Some(formula) = value.get("refersTo").or_else(|| value.get("refers_to")).or_else(|| value.get("formulaOrRefString")).and_then(Value::as_str) else { continue; };
                let map = value.as_object().ok_or_else(|| ExportError::invalid("FORMULA_CALENDAR_PROJECTION_INVALID: invalid defined name"))?;
                let scope = super::defined_names::read_scope(workbook, &safe_names, map, sheet_scope, local)?;
                let actual = scope.or(sheet_scope).unwrap_or(default_owner);
                if let Some(previous) = result.iter().find(|item| item.scope == scope && item.name.to_lowercase() == name.to_lowercase()) {
                    if previous.formula != formula { return Err(ExportError::invalid("FORMULA_CALENDAR_PROJECTION_INVALID: conflicting defined names")); }
                } else { result.push(Name { name: name.to_owned(), formula: formula.to_owned(), owner: actual, scope }); }
                if result.len() > 8192 { return Err(ExportError::invalid("FORMULA_CALENDAR_PROJECTION_LIMIT: more than 8192 authored defined names")); }
            }
        }
    }
    Ok(result)
}

/// Complete the candidate before replacing authored cells or extending sheets.
pub(super) fn prepare(workbook: &mut WorkbookInput, warnings: &mut Vec<ExportWarning>) -> Result<Option<FormulaCalendarSource>, ExportError> {
    build(workbook, warnings)
}

fn build(workbook: &mut WorkbookInput, warnings: &mut Vec<ExportWarning>) -> Result<Option<FormulaCalendarSource>, ExportError> {
    let mut calendars = BTreeSet::new(); let mut saved_sheets = Vec::new(); let mut author_keys = BTreeMap::new(); let mut has_source = false;
    for (owner, sheet) in workbook.sheets.iter().enumerate() {
        let mut cells = Vec::new();
        for (key, value) in &sheet.cells {
            let context = formula_source(value); let retained = date_source(value, workbook.date1904);
            if context.is_none() && retained.is_none() { continue; }
            let cell = cell_from_entry(key, value).ok_or_else(|| ExportError::invalid("FORMULA_CALENDAR_PROJECTION_INVALID: invalid cell coordinate"))?;
            let row = cell.row.ok_or_else(|| ExportError::invalid("FORMULA_CALENDAR_PROJECTION_INVALID: missing cell row"))?;
            let col = cell.col.ok_or_else(|| ExportError::invalid("FORMULA_CALENDAR_PROJECTION_INVALID: missing cell column"))?;
            if let Some(context) = context { calendars.insert(context["date1904"].as_bool().unwrap()); }
            if author_keys.insert((owner, row, col), key.clone()).is_some() { return Err(ExportError::invalid("FORMULA_CALENDAR_PROJECTION_INVALID: duplicate authored cell coordinate")); }
            cells.push(CellSource { row, col, formula: cell.formula, projected_formula: None, formula_source: context.cloned(),
                date_source: retained.cloned(), calculation_sheet: None, calculation_formula: None });
        }
        let import = sheet.metadata.as_ref().and_then(|metadata| metadata.get("chartDependencyImport")).cloned();
        has_source |= !cells.is_empty() || import.is_some();
        saved_sheets.push(SheetSource { name: sheet.name.clone(), sheet_id: sheet.id.clone(), office_sheet_id: (owner + 1) as u32, import, cells });
    }
    if !has_source && workbook.tables.is_empty() { return Ok(None); }
    // Table/column identities and exact dates need no calculation worksheets.
    // Persist them through the same admitted native provenance part.
    if calendars.is_empty() {
        let source = FormulaCalendarSource { version: 1, date1904: workbook.date1904, sheets: saved_sheets,
            helpers: Vec::new(), tables: workbook.tables.iter().map(table_identity_source).collect::<Result<Vec<_>, _>>()? };
        source.to_xml().map_err(ExportError::invalid)?;
        return Ok(Some(source));
    }
    if workbook.sheets.len() > 1024 { return Err(ExportError::invalid("FORMULA_CALENDAR_PROJECTION_LIMIT: more than 1024 source worksheets")); }
    let mut used = workbook.sheets.iter().map(|sheet| sheet.name.to_lowercase()).collect::<BTreeSet<_>>();
    let mut frames = BTreeMap::new();
    for calendar in calendars {
        let views = workbook.sheets.iter().enumerate().map(|(index, _)| {
            let name = unique_name(&format!("_ovcal{}_{}", if calendar { 1904 } else { 1900 }, index + 1), &mut used);
            Mirror { id: format!("office-calendar:{name}"), name }
        }).collect::<Vec<_>>();
        frames.insert(calendar, views);
    }
    let alias_name = unique_name("_ovcal_aliases", &mut used);
    let author_names = names(workbook)?;
    let mut prefix = "_ovfc_".to_owned();
    while workbook.sheets.iter().flat_map(|sheet| sheet.cells.values()).any(|cell| cell.get("formula").and_then(Value::as_str).is_some_and(|formula| formula.to_lowercase().contains(&prefix)))
        || author_names.iter().any(|name| name.name.to_lowercase().contains(&prefix) || name.formula.to_lowercase().contains(&prefix)) {
        prefix.insert(0, '_');
    }
    let mut projection = Projection { workbook, frames, names: author_names, global_names: BTreeMap::new(), tables: BTreeMap::new(),
        alias_sheet: Mirror { id: format!("office-calendar:{alias_name}"), name: alias_name },
        aliases: BTreeMap::new(), alias_cells: serde_json::Map::new(), variable_prefix: prefix, calls: 0, bytes: 0, cells: 0 };
    let mut imported_tables = Vec::new();
    let calendars = projection.frames.keys().copied().collect::<Vec<_>>();
    let mut table_names = workbook.tables.iter().map(|table| table.name.to_lowercase()).chain(projection.names.iter().map(|name| name.name.to_lowercase())).collect::<BTreeSet<_>>();
    for frame in &calendars {
        let global_names = projection.names.iter().filter(|name| name.scope.is_none()).map(|name| name.name.to_lowercase()).collect::<Vec<_>>();
        for (index, name) in global_names.into_iter().enumerate() {
            let mapped = unique_name(&format!("_ovcal_name_{}_{}", usize::from(*frame), index + 1), &mut table_names);
            projection.charge(name.len() + mapped.len() + 96)?;
            projection.global_names.insert((*frame, name), mapped);
        }
    }
    for frame in &calendars {
        for (index, table) in workbook.tables.iter().enumerate() {
            let owner = workbook.sheets.iter().position(|sheet| sheet.id == table.sheet_id)
                .ok_or_else(|| ExportError::invalid("FORMULA_CALENDAR_PROJECTION_INVALID: table source sheet is missing"))?;
            let name = unique_name(&format!("_ovcal_table_{}_{}", usize::from(*frame), index + 1), &mut table_names);
            projection.tables.insert((*frame, table.name.to_lowercase()), name.clone());
            let mut table = table.clone(); table.id = format!("office-calendar:{name}"); table.name = name;
            table.sheet_id = projection.mirror(*frame, owner).id.clone();
            table::clear_calculation_source(&mut table);
            projection.charge(serde_json::to_string(&json!({ "id": table.id, "name": table.name, "columns": table.columns.iter().map(|column| &column.name).collect::<Vec<_>>() })).map_err(|error| ExportError::invalid(error.to_string()))?.len() + 256)?;
            imported_tables.push(table);
        }
    }
    // INDIRECT lookup ranges are built only when an emitted expression uses them.
    let mut helpers = Vec::new(); let mut helper_sources = Vec::new(); let original_names = projection.names.clone();
    for frame in &calendars {
        for (owner, original) in workbook.sheets.iter().enumerate() {
            let mirror = projection.mirror(*frame, owner).clone();
            let mut cells = serde_json::Map::new();
            for (key, cell) in &original.cells {
                let parsed = cell_from_entry(key, cell).ok_or_else(|| ExportError::invalid("FORMULA_CALENDAR_PROJECTION_INVALID: invalid sparse cell"))?;
                let (row, col) = (parsed.row.unwrap_or(0), parsed.col.unwrap_or(0));
                if parsed.formula.is_none() && parsed.value.as_ref().is_none_or(Value::is_null) { continue; }
                projection.cells += 1;
                if projection.cells > MAX_CELLS { return Err(ExportError::invalid("FORMULA_CALENDAR_PROJECTION_LIMIT: calculation views exceed 250000 sparse cells")); }
                let address = format!("{}{}", column_name(col), row + 1);
                let author_reference = format!("{}{address}", quote_sheet(&original.name));
                let context = formula_source(cell); let retained = date_source(cell, workbook.date1904); let date = uses_calendar(cell, workbook.date1904);
                let expression = if let Some(context) = context {
                    let calendar = context["date1904"].as_bool().unwrap();
                    if calendar == *frame {
                        let alias_owner = context.get("aliasSheetId").and_then(Value::as_str).map_or(Some(owner), |id| workbook.sheets.iter().position(|sheet| sheet.id == id));
                        projection.expression(context["formula"].as_str().unwrap().trim_start_matches('='), owner, *frame, alias_owner, 0)?
                    }
                    else {
                        let reference = format!("{}{address}", quote_sheet(&projection.mirror(calendar, owner).name));
                        if date { functions::convert(&reference, calendar, *frame, &format!("{}view_", projection.variable_prefix)) } else { reference }
                    }
                } else if date {
                    let converted = functions::convert(&author_reference, workbook.date1904, *frame, &format!("{}view_", projection.variable_prefix));
                    if let Some(retained) = retained {
                        let before = retained["date1904"].as_bool().unwrap(); let original_value = number(&retained["value"]).unwrap();
                        let old_target = cell.get("value").and_then(number).unwrap();
                        format!("IF({author_reference}={old_target},{},{converted})", calendar_serial(original_value, before, *frame))
                    } else { converted }
                } else { author_reference };
                projection.charge(expression.len() + 256)?;
                let cached = parsed.value.as_ref().and_then(number).map(|value| {
                    let source = retained.and_then(|retained| number(&retained["value"]).zip(retained["date1904"].as_bool()));
                    if date { source.map_or_else(|| calendar_serial(value, workbook.date1904, *frame), |(value, before)| calendar_serial(value, before, *frame)) } else { value }
                });
                let value = cached.map_or_else(|| parsed.value.unwrap_or(Value::Null), |value| json!(value));
                cells.insert(format!("{row}:{col}"), json!({ "row": row, "col": col, "formula": format!("={expression}"), "value": value,
                    "dataType": "formula", "cachedDataType": parsed.cached_data_type.or(parsed.data_type) }));
            }
            let mut definitions = Vec::new();
            // A global definition has a private native name at its real binding
            // owner. A same-name local there cannot hijack another sheet's global.
            for name in &original_names {
                if name.scope.unwrap_or(name.owner) != owner { continue; }
                let formula = projection.expression(name.formula.trim_start_matches('='), name.owner, *frame, Some(name.owner), 0)?;
                let emitted_name = projection.emitted_name(name, *frame).to_owned();
                projection.charge(emitted_name.len() + formula.len() + 192)?;
                definitions.push(json!({ "name": emitted_name, "refersTo": format!("={formula}"), "sheetId": mirror.id, "hidden": true }));
            }
            helpers.push(SheetInput { id: mirror.id.clone(), name: mirror.name.clone(), visible: false, cells,
                row_count: original.row_count, column_count: original.column_count, merged_ranges: original.merged_ranges.clone(),
                metadata: Some(json!({ "hiddenState": "veryHidden", "formulaCalendarHelper": { "version": 1, "date1904": frame, "sourceSheetId": original.id }, "definedNames": definitions })),
                ..SheetInput::default() });
            helper_sources.push(HelperSource { name: mirror.name, sheet_id: mirror.id, office_sheet_id: (workbook.sheets.len() + helpers.len()) as u32,
                date1904: *frame, source_sheet_id: Some(original.id.clone()), fingerprint: String::new(), reclaim: false });
        }
    }
    if !projection.alias_cells.is_empty() {
        helpers.push(SheetInput { id: projection.alias_sheet.id.clone(), name: projection.alias_sheet.name.clone(), visible: false,
            cells: std::mem::take(&mut projection.alias_cells), metadata: Some(json!({ "hiddenState": "veryHidden", "formulaCalendarHelper": { "version": 1, "kind": "aliases" } })), ..SheetInput::default() });
        helper_sources.push(HelperSource { name: projection.alias_sheet.name.clone(), sheet_id: projection.alias_sheet.id.clone(), office_sheet_id: (workbook.sheets.len() + helpers.len()) as u32,
            date1904: workbook.date1904, source_sheet_id: None, fingerprint: String::new(), reclaim: false });
    }
    let helper_indices = helpers.iter().enumerate().map(|(index, helper)| (helper.id.as_str(), index)).collect::<BTreeMap<_, _>>();
    let mut replacements = Vec::new();
    for saved in &mut saved_sheets {
        let owner = workbook.sheets.iter().position(|sheet| sheet.id == saved.sheet_id).unwrap();
        for cell in &mut saved.cells {
            let Some(context) = cell.formula_source.as_ref() else { continue; };
            let calendar = context["date1904"].as_bool().unwrap(); let mirror = projection.mirror(calendar, owner);
            let reference = format!("{}{}{}", quote_sheet(&mirror.name), column_name(cell.col), cell.row + 1);
            let formula = if context["dateResult"].as_bool() == Some(true) { functions::convert(&reference, calendar, workbook.date1904, &format!("{}author_", projection.variable_prefix)) } else { reference };
            cell.projected_formula = Some(format!("={formula}")); cell.calculation_sheet = Some(mirror.name.clone());
            let calculation = helper_indices.get(mirror.id.as_str()).and_then(|index| helpers[*index].cells.get(&format!("{}:{}", cell.row, cell.col)))
                .and_then(|cell| cell.get("formula")).and_then(Value::as_str).ok_or_else(|| ExportError::invalid("FORMULA_CALENDAR_PROJECTION_INVALID: calculation formula disappeared"))?;
            cell.calculation_formula = Some(calculation.to_owned());
            let key = author_keys.get(&(owner, cell.row, cell.col)).ok_or_else(|| ExportError::invalid("FORMULA_CALENDAR_PROJECTION_INVALID: author coordinate disappeared"))?.clone();
            replacements.push((owner, key, cell.projected_formula.clone().unwrap()));
        }
    }
    let reserved_ids = workbook.tables.iter().filter_map(|table| table.office_table_id).collect::<Vec<_>>();
    table::prepare_identities(&mut imported_tables, &reserved_ids)?;
    let source_tables = workbook.tables.iter().chain(&imported_tables).map(table_identity_source).collect::<Result<Vec<_>, _>>()?;
    let sheet_identities = saved_sheets.iter().map(|sheet| (sheet.name.to_lowercase(), sheet.sheet_id.clone()))
        .chain(helper_sources.iter().map(|sheet| (sheet.name.to_lowercase(), sheet.sheet_id.clone()))).collect::<BTreeMap<_, _>>();
    let table_identities = source_tables.iter().map(|table| (table.name.to_lowercase(), table.table_id.clone())).collect::<BTreeMap<_, _>>();
    for (helper, descriptor) in helpers.iter().zip(&mut helper_sources) {
        descriptor.fingerprint = helper_fingerprint(helper, &imported_tables, &sheet_identities, &table_identities)?;
    }
    let source = FormulaCalendarSource { version: 1, date1904: workbook.date1904, sheets: saved_sheets, helpers: helper_sources, tables: source_tables };
    // Admit provenance and every helper before publishing any workbook mutation.
    let resource = source.to_xml().map_err(ExportError::invalid)?; projection.charge(resource.len())?;
    let bytes = projection.bytes; let cell_count = projection.cells; drop(projection);
    for (owner, key, formula) in replacements {
        workbook.sheets[owner].cells.get_mut(&key).and_then(Value::as_object_mut).ok_or_else(|| ExportError::invalid("FORMULA_CALENDAR_PROJECTION_INVALID: author cell is not an object"))?
            .insert("formula".into(), Value::String(formula));
    }
    workbook.sheets.extend(helpers); workbook.tables.extend(imported_tables);
    if cell_count > 0 { warnings.push(ExportWarning::new("XLSX_FORMULA_CALENDAR_VIEWS", format!("Imported formulas use {cell_count} sparse calendar-view cells ({bytes} admitted Rust projection bytes); source formulas and exact serials are preserved in a custom XML part. LET support, external Office edits and calendar edge cases require interoperability verification."))); }
    Ok(Some(source))
}

fn table_identity_source(table: &table::TableInput) -> Result<TableSource, ExportError> {
    Ok(TableSource { name: table.name.clone(), table_id: table.id.clone(), sheet_id: table.sheet_id.clone(),
        office_table_id: table.office_table_id.ok_or_else(|| ExportError::invalid("TABLE_SOURCE_INVALID: table has no prepared native identity"))?,
        columns: table.columns.iter().map(|column| Ok(TableColumnSource { column_id: column.id.clone(),
            office_column_id: column.office_column_id.ok_or_else(|| ExportError::invalid("TABLE_SOURCE_INVALID: column has no prepared native identity"))? }))
            .collect::<Result<Vec<_>, ExportError>>()? })
}

fn helper_fingerprint(sheet: &SheetInput, tables: &[table::TableInput], sheets: &BTreeMap<String, String>, table_names: &BTreeMap<String, String>) -> Result<String, ExportError> {
    let mut fingerprint = CalculationFingerprint::default();
    for (key, value) in &sheet.cells {
        let cell = cell_from_entry(key, value).ok_or_else(|| ExportError::invalid("FORMULA_CALENDAR_PROJECTION_INVALID: invalid helper coordinate"))?;
        fingerprint.cell(cell.row.unwrap_or(0), cell.col.unwrap_or(0), cell.formula.as_deref(), cell.value.as_ref().unwrap_or(&Value::Null), sheets, table_names).map_err(ExportError::invalid)?;
    }
    for range in &sheet.merged_ranges { fingerprint.merged_range(super::range_bounds(range).ok_or_else(|| ExportError::invalid("FORMULA_CALENDAR_PROJECTION_INVALID: invalid helper merge"))?).map_err(ExportError::invalid)?; }
    for name in sheet.metadata.as_ref().and_then(|metadata| metadata.get("definedNames")).and_then(Value::as_array).into_iter().flatten() {
        fingerprint.defined_name(name["name"].as_str().unwrap(), name["refersTo"].as_str().unwrap(), sheets, table_names).map_err(ExportError::invalid)?;
    }
    for table in tables.iter().filter(|table| table.sheet_id == sheet.id) {
        let columns = table.columns.iter().map(|column| column.name.clone()).collect::<Vec<_>>();
        fingerprint.table(&table.id, super::range_bounds(&table.range).ok_or_else(|| ExportError::invalid("FORMULA_CALENDAR_PROJECTION_INVALID: invalid helper table"))?, &columns, table.show_header_row, table.show_totals_row).map_err(ExportError::invalid)?;
    }
    Ok(fingerprint.finish())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn projected(formula: &str, metadata: Value) -> String {
        let mut workbook: WorkbookInput = serde_json::from_value(json!({
            "kind": "workbook", "date1904": true, "activeSheetId": "data", "metadata": metadata,
            "sheets": [
                { "id": "data", "name": "Data", "cells": {
                    "0:0": { "row": 0, "col": 0, "formula": formula, "value": 1,
                        "metadata": { "chartDependencyFormulaSource": {
                            "version": 1, "formula": formula, "date1904": false, "dateResult": false
                        } } }
                } },
                { "id": "other", "name": "Other's data", "cells": {} }
            ]
        })).unwrap();
        let source = prepare(&mut workbook, &mut Vec::new()).unwrap().unwrap();
        source.sheets[0].cells[0].calculation_formula.clone().unwrap()
    }

    #[test]
    fn registered_calls_are_not_hijacked_by_equal_defined_names() {
        let formula = projected("=SUM(1)+DAY(60)+SUM", json!({ "definedNames": [
            { "name": "SUM", "refersTo": "=7" }, { "name": "DAY", "refersTo": "=99" }
        ] }));
        assert!(formula.starts_with("=SUM(1)+_xlfn.LET("));
        assert!(formula.contains("=60,29,"));
        // The bare name still binds to its isolated global declaration.
        assert!(formula.ends_with("!_ovcal_name_0_1"), "{formula}");
    }

    #[test]
    fn lambda_parameters_have_native_prefixes_and_case_insensitive_body_bindings() {
        let formula = projected("=LAMBDA(Temp,TEMP+1)(2)", json!({ "definedNames": [
            { "name": "Temp", "refersTo": "=100" }
        ] }));
        assert_eq!(formula, "=_xlfn.LAMBDA(_xlpm.Temp,_xlpm.Temp+1)(2)");
        let prefixed = projected("=_xlfn.LAMBDA(_XLPM.Temp,_xlpm.temp+1)(2)", json!({}));
        assert_eq!(prefixed, formula);
    }

    #[test]
    fn let_value_reads_the_outer_lambda_before_the_new_local_shadows_it() {
        let formula = projected("=LAMBDA(Temp,LET(Temp,TEMP+1,Temp))(2)", json!({}));
        assert_eq!(formula, "=_xlfn.LAMBDA(_xlpm.Temp,_xlfn.LET(Temp,_xlpm.Temp+1,Temp))(2)");
    }

    #[test]
    fn lexical_callbacks_keep_their_names_without_builtin_date_projection() {
        let formula = projected("=LET(DAY,LAMBDA(value,value+1),DAY(60))", json!({}));
        assert_eq!(formula, "=_xlfn.LET(DAY,_xlfn.LAMBDA(_xlpm.value,_xlpm.value+1),DAY(60))");
    }

    #[test]
    fn qualified_named_lambdas_keep_one_prefix_and_the_declared_scope() {
        let formula = projected("='Other''s data'!Tax(2)", json!({ "localDefinedNames": [
            { "name": "Tax", "refersTo": "=LAMBDA(amount,amount*2)", "sheetId": "other" }
        ] }));
        assert_eq!(formula, "='_ovcal1900_2'!Tax(2)");
    }

    #[test]
    fn future_builtin_prefixes_still_outrank_defined_names() {
        let formula = projected("=_xlfn._xlws.FILTER({1;2},{TRUE;FALSE})", json!({ "definedNames": [
            { "name": "FILTER", "refersTo": "=LAMBDA(value,value)" }
        ] }));
        assert_eq!(formula, "=_xlfn._xlws.FILTER({1;2},{TRUE;FALSE})");
    }
}
