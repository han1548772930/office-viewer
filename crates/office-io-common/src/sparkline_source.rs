//! Retained broken and named/foreign source syntax. This module does not evaluate formulas.
use serde::{Deserialize, Serialize};
use super::parse_source_reference;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SparklineSourceKind { A1, Named, External, ThreeDimensional, Broken }

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct SparklineSourceReference {
    pub version: u8,
    pub invalid_range: SparklineBrokenSourceRange,
}
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct SparklineBrokenSourceRange {
    pub start_row: i32,
    pub end_row: i32,
    pub start_column: i32,
    pub end_column: i32,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sheet_id: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub unit_id: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub range_type: Option<u8>,
}
impl SparklineSourceReference {
    pub fn validate(&self) -> Result<(), String> {
        let range = &self.invalid_range;
        if self.version != 1 || [range.start_row, range.end_row, range.start_column, range.end_column].iter().any(|value| *value != -9999)
            || range.range_type.is_some_and(|value| value > 3) {
            return Err("SPARKLINE_REFERENCE_INVALID: invalid version or -9999 retained source sentinel".into());
        }
        for id in [&range.sheet_id, &range.unit_id].into_iter().flatten() {
            if id.trim().is_empty() || id.len() > 4096 || id.chars().any(char::is_control) {
                return Err("SPARKLINE_REFERENCE_INVALID: invalid retained source identity".into());
            }
        }
        Ok(())
    }
}

fn named(value: &str) -> bool {
    let mut chars = value.chars();
    let Some(first) = chars.next() else { return false; };
    if value.chars().count() > 255 || !(first.is_alphabetic() || matches!(first, '_' | '\\'))
        || !chars.all(|character| character.is_alphanumeric() || matches!(character, '_' | '.' | '\\'))
        || value.eq_ignore_ascii_case("r") || value.eq_ignore_ascii_case("c") || parse_source_reference(value).is_ok() { return false; }
    let lower = value.to_ascii_lowercase();
    if let Some((row, col)) = lower.strip_prefix('r').and_then(|value| value.split_once('c')) {
        if !row.is_empty() && !col.is_empty() && row.bytes().all(|value| value.is_ascii_digit()) && col.bytes().all(|value| value.is_ascii_digit()) { return false; }
    }
    true
}
fn unquote(value: &str) -> Option<String> {
    if let Some(inner) = value.strip_prefix('\'').and_then(|value| value.strip_suffix('\'')) {
        (!inner.replace("''", "").contains('\'')).then(|| inner.replace("''", "'"))
    } else { (!value.is_empty() && !value.contains('\'') && value.trim() == value).then(|| value.to_owned()) }
}
fn parts(text: &str) -> Result<(Option<&str>, &str, bool), String> {
    let mut quoted = false; let mut brackets = 0usize; let mut delimiter = None; let mut broken = false;
    let mut chars = text.char_indices().peekable();
    while let Some((index, character)) = chars.next() {
        if character == '\'' {
            if quoted && chars.peek().is_some_and(|(_, next)| *next == '\'') { chars.next(); } else { quoted = !quoted; }
        } else if !quoted {
            match character {
                '[' => brackets += 1,
                ']' => { brackets = brackets.checked_sub(1).ok_or("SPARKLINE_REFERENCE_INVALID: unmatched external workbook bracket")?; },
                '#' if brackets == 0 && text[index..].get(..5).is_some_and(|value| value.eq_ignore_ascii_case("#REF!")) => {
                    broken = true; for _ in 0..4 { chars.next(); }
                },
                '!' if brackets == 0 => {
                    if delimiter.replace(index).is_some() { return Err("SPARKLINE_REFERENCE_INVALID: multiple source qualifiers".into()); }
                },
                _ => {},
            }
        }
    }
    if quoted || brackets != 0 { return Err("SPARKLINE_REFERENCE_INVALID: unclosed source quote/bracket".into()); }
    Ok(match delimiter { Some(index) => (Some(&text[..index]), &text[index + 1..], broken), None => (None, text, broken) })
}

pub fn source_kind(input: &str) -> Result<SparklineSourceKind, String> {
    if input.trim().is_empty() || input.len() > 4096 || input.chars().any(char::is_control) {
        return Err("SPARKLINE_REFERENCE_INVALID: empty/control/over-budget source text".into());
    }
    let text = input.trim().strip_prefix('=').unwrap_or(input.trim()).trim();
    if parse_source_reference(text).is_ok() { return Ok(SparklineSourceKind::A1); }
    let (qualifier, body, broken) = parts(text)?;
    if broken { return Ok(SparklineSourceKind::Broken); }
    if let Some(raw) = qualifier {
        let qualifier = unquote(raw).ok_or("SPARKLINE_REFERENCE_INVALID: invalid source qualifier")?;
        if qualifier.is_empty() || !(parse_source_reference(body).is_ok() || named(body)) { return Err("SPARKLINE_REFERENCE_INVALID: source must be a range or name".into()); }
        if let Some(start) = qualifier.find('[') {
            let end = qualifier[start + 1..].find(']').map(|index| start + index + 1).ok_or("SPARKLINE_REFERENCE_INVALID: unclosed external workbook")?;
            if end == start + 1 || qualifier[start + 1..end].contains('[') || qualifier[..start].contains(']')
                || qualifier[end + 1..].contains(['[', ']']) { return Err("SPARKLINE_REFERENCE_INVALID: invalid external workbook qualifier".into()); }
            return Ok(SparklineSourceKind::External);
        }
        if qualifier.contains(':') {
            let sheets = qualifier.split(':').collect::<Vec<_>>();
            if sheets.len() != 2 || sheets.iter().any(|sheet| sheet.is_empty() || sheet.contains(['[', ']', '\\', '/', '?', '*'])) {
                return Err("SPARKLINE_REFERENCE_INVALID: invalid three-dimensional boundaries".into());
            }
            return Ok(SparklineSourceKind::ThreeDimensional);
        }
        if qualifier.contains(['[', ']', ':', '\\', '/', '?', '*']) || !named(body) { return Err("SPARKLINE_REFERENCE_INVALID: invalid named source qualifier".into()); }
        return Ok(SparklineSourceKind::Named);
    }
    if let Some(rest) = text.strip_prefix('[') {
        if let Some((book, name)) = rest.split_once(']') {
            if !book.is_empty() && !book.contains(['[', ']']) && named(name) { return Ok(SparklineSourceKind::External); }
        }
    }
    if named(text) { return Ok(SparklineSourceKind::Named); }
    Err("SPARKLINE_REFERENCE_INVALID: source is not an admitted range/name/broken reference".into())
}
