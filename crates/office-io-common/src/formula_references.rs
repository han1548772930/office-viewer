//! Shared A1 reference spans for lossless Office projections and provenance.
//! This scanner preserves string literals and structured-table captions.

const MAX_ROW: u32 = 1_048_575;
const MAX_COLUMN: u32 = 16_383;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ReferenceKind { Cell, WholeRow, WholeColumn }

#[derive(Debug, Clone)]
pub struct Reference {
    pub start: usize, pub end: usize, pub prefix: String, pub kind: ReferenceKind,
    pub row: u32, pub col: u32, pub absolute_row: bool, pub absolute_col: bool,
}

fn identifier(character: char) -> bool { character.is_alphanumeric() || matches!(character, '_' | '.' | '\\') }
fn qualifier_end(source: &str, start: usize) -> Option<usize> {
    let first = source[start..].chars().next()?;
    if !first.is_alphabetic() && !matches!(first, '_' | '\\') { return None }
    let mut cursor = start;
    while let Some(character) = source[cursor..].chars().next() {
        if !identifier(character) { break }
        cursor += character.len_utf8();
    }
    if source.as_bytes().get(cursor) == Some(&b':') {
        let second = cursor + 1;
        let first = source[second..].chars().next()?;
        if !first.is_alphabetic() && !matches!(first, '_' | '\\') { return None; }
        cursor = second;
        while let Some(character) = source[cursor..].chars().next() {
            if !identifier(character) { break }
            cursor += character.len_utf8();
        }
    }
    (source.as_bytes().get(cursor) == Some(&b'!')).then_some(cursor + 1)
}

fn read_address(source: &str, start: usize) -> Option<Reference> {
    let mut cursor = start;
    let bytes = source.as_bytes();
    let absolute_first = bytes.get(cursor) == Some(&b'$');
    if absolute_first { cursor += 1; }
    let letters = cursor;
    while bytes.get(cursor).is_some_and(u8::is_ascii_alphabetic) { cursor += 1; }
    if cursor - letters > 3 { return None }
    let has_column = cursor > letters;
    let col = if has_column {
        source[letters..cursor].bytes().fold(0u32, |value, letter| value * 26 + u32::from(letter.to_ascii_uppercase() - b'A') + 1).checked_sub(1)?
    } else { 0 };
    let absolute_row = if has_column { bytes.get(cursor) == Some(&b'$') } else { absolute_first };
    if has_column && absolute_row { cursor += 1; }
    let digits = cursor;
    while bytes.get(cursor).is_some_and(u8::is_ascii_digit) { cursor += 1; }
    let has_row = cursor > digits;
    if !has_row && (!has_column || absolute_row) { return None }
    if source[cursor..].chars().next().is_some_and(identifier) || source[cursor..].trim_start().starts_with('(') { return None }
    let row = if has_row { source[digits..cursor].parse::<u32>().ok()?.checked_sub(1)? } else { 0 };
    if row > MAX_ROW || col > MAX_COLUMN { return None }
    Some(Reference { start, end: cursor, prefix: String::new(),
        kind: if !has_column { ReferenceKind::WholeRow } else if !has_row { ReferenceKind::WholeColumn } else { ReferenceKind::Cell },
        row, col, absolute_row: !has_row || absolute_row, absolute_col: !has_column || absolute_first })
}

pub fn references(source: &str) -> Vec<Reference> {
    let mut result = Vec::new();
    let mut cursor = 0;
    while cursor < source.len() {
        let character = source[cursor..].chars().next().unwrap();
        let start = cursor;
        if character == '"' {
            cursor += 1;
            while cursor < source.len() {
                if source.as_bytes()[cursor] == b'"' {
                    cursor += 1;
                    if source.as_bytes().get(cursor) == Some(&b'"') { cursor += 1; } else { break }
                } else { cursor += source[cursor..].chars().next().unwrap().len_utf8(); }
            }
            continue;
        }
        if character == '\'' {
            cursor += 1;
            while cursor < source.len() {
                if source.as_bytes()[cursor] == b'\'' {
                    cursor += 1;
                    if source.as_bytes().get(cursor) == Some(&b'\'') { cursor += 1; continue }
                    if source.as_bytes().get(cursor) == Some(&b'!') {
                        if let Some(mut reference) = read_address(source, cursor + 1) {
                            reference.start = start; reference.prefix = source[start..cursor + 1].to_owned(); cursor = reference.end; result.push(reference);
                        } else { cursor += 1; }
                    }
                    break;
                }
                cursor += source[cursor..].chars().next().unwrap().len_utf8();
            }
            continue;
        }
        if character == '[' {
            let mut nesting = 1;
            cursor += 1;
            while cursor < source.len() && nesting > 0 {
                let next = source[cursor..].chars().next().unwrap();
                // A structured caption uses an apostrophe to escape [, ], #,
                // @ or apostrophe. Its escaped closing bracket is not a scope end.
                if next == '\'' && source[cursor + 1..].chars().next().is_some_and(|value| matches!(value, '[' | ']' | '#' | '@' | '\'')) {
                    cursor += 1;
                    cursor += source[cursor..].chars().next().unwrap().len_utf8();
                    continue;
                }
                if next == '[' { nesting += 1; } else if next == ']' { nesting -= 1; }
                cursor += next.len_utf8();
            }
            if !source[..start].chars().next_back().is_some_and(identifier) {
                if let Some(address_start) = qualifier_end(source, cursor) {
                    if let Some(mut reference) = read_address(source, address_start) {
                        reference.start = start; reference.prefix = source[start..address_start].to_owned(); cursor = reference.end; result.push(reference);
                    }
                }
            }
            continue;
        }
        if source[..cursor].chars().next_back().is_some_and(identifier) { cursor += character.len_utf8(); continue }
        let address_start = qualifier_end(source, cursor).unwrap_or(cursor);
        if let Some(mut reference) = read_address(source, address_start) {
            reference.start = cursor; reference.prefix = source[cursor..address_start].to_owned(); cursor = reference.end; result.push(reference);
        } else { cursor += character.len_utf8(); }
    }
    // A bare name or numeric literal is not an axis reference. Both endpoints
    // must describe the same whole axis and be joined by the range operator.
    let axis_ranges = result.iter().enumerate().map(|(index, reference)| {
        reference.kind == ReferenceKind::Cell
            || index > 0 && result[index - 1].kind == reference.kind && source[result[index - 1].end..reference.start].trim() == ":"
            || index + 1 < result.len() && result[index + 1].kind == reference.kind && source[reference.end..result[index + 1].start].trim() == ":"
    }).collect::<Vec<_>>();
    result.into_iter().zip(axis_ranges).filter_map(|(reference, keep)| keep.then_some(reference)).collect()
}

/// Rewrite worksheet qualifiers, including qualified names and 3D ranges.
/// Values inside double-quoted INDIRECT strings remain authored aliases.
/// External workbook qualifiers are preserved as external references.
pub fn rewrite_sheet_qualifiers(source: &str, names: &std::collections::BTreeMap<String, String>) -> String {
    walk_sheet_qualifiers(source, |caption, original| mapped_qualifier(caption, original, names))
}

/// Shared qualifier visitor also sees worksheet-qualified defined names and
/// 3D endpoints. Quoted formula strings and external books stay separate.
pub fn sheet_qualifiers(source: &str) -> Vec<String> {
    let mut result = Vec::new();
    walk_sheet_qualifiers(source, |caption, original| {
        if !caption.contains(['[', ']']) { result.extend(caption.split(':').map(str::to_owned)); }
        original.to_owned()
    });
    result
}

fn walk_sheet_qualifiers(source: &str, mut map: impl FnMut(&str, &str) -> String) -> String {
    let mut output = String::with_capacity(source.len());
    let mut cursor = 0;
    while cursor < source.len() {
        let start = cursor;
        let character = source[cursor..].chars().next().unwrap();
        if character == '"' || character == '\'' {
            let delimiter = character as u8;
            cursor += 1;
            while cursor < source.len() {
                if source.as_bytes()[cursor] == delimiter {
                    cursor += 1;
                    if source.as_bytes().get(cursor) == Some(&delimiter) { cursor += 1; }
                    else { break; }
                } else { cursor += source[cursor..].chars().next().unwrap().len_utf8(); }
            }
            if character == '\'' && source.as_bytes().get(cursor) == Some(&b'!') {
                let caption = source[start + 1..cursor - 1].replace("''", "'");
                cursor += 1;
                output.push_str(&map(&caption, &source[start..cursor]));
            } else { output.push_str(&source[start..cursor]); }
            continue;
        }
        if character == '[' {
            let mut depth = 1;
            cursor += 1;
            while cursor < source.len() && depth > 0 {
                let next = source[cursor..].chars().next().unwrap();
                if next == '\'' && source[cursor + 1..].chars().next().is_some_and(|value| matches!(value, '[' | ']' | '#' | '@' | '\'')) {
                    cursor += 1;
                    cursor += source[cursor..].chars().next().unwrap().len_utf8();
                    continue;
                }
                if next == '[' { depth += 1; } else if next == ']' { depth -= 1; }
                cursor += next.len_utf8();
            }
            // [Book.xlsx]Sheet!A1 must not be rebound to a local Sheet.
            if let Some(end) = qualifier_end(source, cursor) { cursor = end; }
            output.push_str(&source[start..cursor]);
            continue;
        }
        if !source[..cursor].chars().next_back().is_some_and(identifier)
            && let Some(end) = qualifier_end(source, cursor) {
            let caption = &source[cursor..end - 1];
            output.push_str(&map(caption, &source[cursor..end]));
            cursor = end;
        } else { cursor += character.len_utf8(); output.push(character); }
    }
    output
}

fn mapped_qualifier(caption: &str, original: &str, names: &std::collections::BTreeMap<String, String>) -> String {
    if caption.contains(['[', ']']) { return original.to_owned(); }
    let parts = caption.split(':').map(|name| names.get(&name.to_lowercase())).collect::<Option<Vec<_>>>();
    if let Some(parts) = parts {
        let name = parts.into_iter().map(String::as_str).collect::<Vec<_>>().join(":");
        format!("'{}'!", name.replace('\'', "''"))
    } else { original.to_owned() }
}

pub fn rewrite_table_identifiers(source: &str, names: &std::collections::BTreeMap<String, String>) -> String {
    walk_identifiers(source, |name, callable| {
        if !callable && let Some(mapped) = names.get(&name.to_lowercase()) { mapped.clone() } else { name.to_owned() }
    })
}

pub fn identifiers(source: &str) -> Vec<String> {
    let mut result = Vec::new();
    walk_identifiers(source, |name, callable| { if !callable { result.push(name.to_owned()); } name.to_owned() });
    result
}

fn walk_identifiers(source: &str, mut map: impl FnMut(&str, bool) -> String) -> String {
    let mut output = String::with_capacity(source.len()); let mut cursor = 0;
    while cursor < source.len() {
        let start = cursor; let character = source[cursor..].chars().next().unwrap();
        if character == '"' || character == '\'' || character == '[' {
            if character == '[' {
                let mut depth = 1; cursor += 1;
                while cursor < source.len() && depth > 0 {
                    let next = source[cursor..].chars().next().unwrap();
                    if next == '\'' && source[cursor + 1..].chars().next().is_some_and(|value| matches!(value, '[' | ']' | '#' | '@' | '\'')) {
                        cursor += 1; cursor += source[cursor..].chars().next().unwrap().len_utf8(); continue;
                    }
                    if next == '[' { depth += 1; } else if next == ']' { depth -= 1; }
                    cursor += next.len_utf8();
                }
            } else {
                cursor += 1;
                while cursor < source.len() {
                    if source.as_bytes()[cursor] == character as u8 {
                        cursor += 1;
                        if source.as_bytes().get(cursor) == Some(&(character as u8)) { cursor += 1; } else { break; }
                    } else { cursor += source[cursor..].chars().next().unwrap().len_utf8(); }
                }
            }
            output.push_str(&source[start..cursor]); continue;
        }
        if character.is_alphabetic() || matches!(character, '_' | '\\') {
            cursor += character.len_utf8();
            while source[cursor..].chars().next().is_some_and(identifier) { cursor += source[cursor..].chars().next().unwrap().len_utf8(); }
            let name = &source[start..cursor];
            output.push_str(&map(name, source[cursor..].trim_start().starts_with('(')));
        } else { cursor += character.len_utf8(); output.push(character); }
    }
    output
}

/// Case and standard future-function prefixes can change during an Office
/// save. Literal string contents, escaped captions and intersection spaces are
/// retained. This is a freshness comparison, not formula evaluation.
pub fn canonical_formula(source: &str, sheets: &std::collections::BTreeMap<String, String>, tables: &std::collections::BTreeMap<String, String>) -> String {
    let mapped = rewrite_table_identifiers(&rewrite_sheet_qualifiers(source.trim().trim_start_matches('='), sheets), tables);
    let mut output = String::with_capacity(mapped.len()); let mut cursor = 0;
    while cursor < mapped.len() {
        let start = cursor; let character = mapped[cursor..].chars().next().unwrap();
        if character == '"' || character == '\'' {
            cursor += 1;
            while cursor < mapped.len() {
                if mapped.as_bytes()[cursor] == character as u8 {
                    cursor += 1;
                    if mapped.as_bytes().get(cursor) == Some(&(character as u8)) { cursor += 1; } else { break; }
                } else { cursor += mapped[cursor..].chars().next().unwrap().len_utf8(); }
            }
            output.push_str(&mapped[start..cursor]); continue;
        }
        if character.is_alphabetic() || matches!(character, '_' | '\\') {
            cursor += character.len_utf8();
            while mapped[cursor..].chars().next().is_some_and(identifier) { cursor += mapped[cursor..].chars().next().unwrap().len_utf8(); }
            let name = mapped[start..cursor].to_lowercase();
            let normalized = name.strip_prefix("_xlfn.").unwrap_or(&name);
            output.push_str(normalized.strip_prefix("_xlpm.").unwrap_or(normalized));
        } else { cursor += character.len_utf8(); output.push(character); }
    }
    output
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::collections::BTreeMap;

    #[test]
    fn scanner_preserves_escaped_structured_captions_and_literal_addresses() {
        let formula = r#"SUM(T[[#Data],[A']B]],'A''s'!$C$4,D9,"A1","T[A']B]",First:Last!A:A)"#;
        let found = references(formula);
        assert_eq!(found.len(), 4);
        assert_eq!(found[0].prefix, "'A''s'!");
        assert_eq!((found[0].row, found[0].col), (3, 2));
        assert_eq!(found[2].prefix, "First:Last!");
        assert_eq!(found[2].kind, ReferenceKind::WholeColumn);
        assert_eq!(found[3].kind, ReferenceKind::WholeColumn);
    }
    #[test]
    fn qualifier_rewrites_cover_names_and_unicode_without_rewriting_indirect_text() {
        let names = BTreeMap::from([("ä".into(), "Renamed".into()), ("a's".into(), "O'Brien".into()), ("last".into(), "End".into())]);
        assert_eq!(rewrite_sheet_qualifiers(r#"Ä!Tax+'A''s'!A1+Ä:Last!B:B+INDIRECT("Ä!A1")+[Book.xlsx]Ä!A1+T[A']B]"#, &names),
            r#"'Renamed'!Tax+'O''Brien'!A1+'Renamed:End'!B:B+INDIRECT("Ä!A1")+[Book.xlsx]Ä!A1+T[A']B]"#);
    }
}

