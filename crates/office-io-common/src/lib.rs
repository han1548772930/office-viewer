//! Shared, pure Office geometry and worksheet-dimension contracts.

use quick_xml::{Reader, XmlVersion};
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};

mod filter_projection;
pub use filter_projection::{OfficeFilterNode, WorksheetFilterProjection, worksheet_filter_fingerprint};

pub mod conditional_format_xml;
pub mod conditional_format_source;
pub mod conditional_format_color;
pub mod sheet_validation_source;
pub mod document_legacy_source;
pub mod chart_kind;
pub mod chart_options;
pub mod chart_color;
pub mod chart_xml;
pub mod excel_date;
pub mod document_table_source;
pub mod document_numbering_source;

pub const MAXIMUM_DIGIT_WIDTH_PIXELS: f64 = 7.0;
pub const CSS_PIXELS_PER_POINT: f64 = 4.0 / 3.0;
pub const SHEET_AXIS_NAMESPACE: &str = "urn:office-studio:sheet-axis:1";
pub const SHEET_AXIS_EXTENSION_URI: &str = "{F41EED10-CA42-42C0-A66C-9152ABCB9AE1}";

pub fn points_to_pixels(points: f64) -> f64 { points * CSS_PIXELS_PER_POINT }
pub fn pixels_to_points(pixels: f64) -> f64 { pixels / CSS_PIXELS_PER_POINT }
pub fn stored_character_width_to_pixels(width: f64) -> f64 {
    ((width * 256.0 + (128.0 / MAXIMUM_DIGIT_WIDTH_PIXELS).floor()) / 256.0 * MAXIMUM_DIGIT_WIDTH_PIXELS).floor()
}
pub fn pixels_to_stored_character_width(pixels: f64) -> f64 {
    (pixels / MAXIMUM_DIGIT_WIDTH_PIXELS * 256.0).round() / 256.0
}

/// A content-independent fingerprint of the authored Office axis projection.
/// Prefixes, attribute order and redundant false flags do not affect it. Row
/// contents/styles are excluded, so external edits invalidate only stale axes.
/// This is an accidental-staleness guard, not a security signature.
pub fn worksheet_axis_fingerprint(xml: &str) -> Result<String, String> {
    let mut reader = Reader::from_str(xml);
    let mut fingerprint = WorksheetAxisFingerprint::default();
    loop {
        match reader.read_event() {
            Ok(Event::Start(element)) => fingerprint.start(&element, reader.decoder(), false)?,
            Ok(Event::Empty(element)) => fingerprint.start(&element, reader.decoder(), true)?,
            Ok(Event::End(_)) => fingerprint.end()?,
            Ok(Event::DocType(_)) => return Err("Office worksheet DTDs are unsupported".into()),
            Ok(Event::Eof) => break,
            Err(error) => return Err(error.to_string()),
            _ => {}
        }
    }
    // The writer calls this on a well-formed prefix before adding extLst and
    // </worksheet>; streaming readers separately require a closed XML document.
    Ok(fingerprint.finish())
}

/// Streaming counterpart used by the parser without retaining the cell XML.
pub struct WorksheetAxisFingerprint { hash: u64, path: Vec<Vec<u8>>, zero_height: bool }

impl Default for WorksheetAxisFingerprint {
    fn default() -> Self { Self { hash: 0xcbf29ce484222325, path: Vec::new(), zero_height: false } }
}

impl WorksheetAxisFingerprint {
    pub fn depth(&self) -> usize { self.path.len() }
    pub fn path_is(&self, path: &[&[u8]]) -> bool {
        self.path.len() == path.len() && self.path.iter().zip(path).all(|(actual, expected)| actual.as_slice() == *expected)
    }
    fn hash_text(&mut self, text: &str) {
        for byte in text.bytes() { self.hash ^= u64::from(byte); self.hash = self.hash.wrapping_mul(0x100000001b3); }
    }
    pub fn start(&mut self, element: &BytesStart<'_>, decoder: Decoder, empty: bool) -> Result<(), String> {
        if self.path.len() >= 256 { return Err("Office worksheet nesting exceeds the parser limit".into()); }
        let name = element.local_name();
        let allowed: &[&str] = match name.as_ref() {
            b"sheetFormatPr" if self.path_is(&[b"worksheet"]) => &["defaultColWidth", "defaultRowHeight", "customHeight", "zeroHeight"],
            b"col" if self.path_is(&[b"worksheet", b"cols"]) => &["min", "max", "width", "customWidth", "hidden"],
            b"row" if self.path_is(&[b"worksheet", b"sheetData"]) => &["r", "ht", "customHeight", "hidden"],
            _ => &[],
        };
        if !allowed.is_empty() {
            let mut attributes = Vec::new();
            for attribute in element.attributes() {
                let attribute = attribute.map_err(|error| error.to_string())?;
                let key = std::str::from_utf8(attribute.key.as_ref()).map_err(|error| error.to_string())?;
                if !allowed.contains(&key) { continue; }
                let value = attribute.decode_and_unescape_value(decoder, XmlVersion::Implicit1_0).map_err(|error| error.to_string())?;
                let flag = matches!(key, "customHeight" | "customWidth" | "zeroHeight" | "hidden");
                let normalized = if flag {
                    let enabled = match value.as_ref() { "1" | "true" => true, "0" | "false" => false, _ => return Err("Invalid Office axis flag".into()) };
                    if key == "zeroHeight" { self.zero_height = enabled; }
                    // Under zeroHeight=1 an explicit visible row overrides the
                    // default; its false hidden flag is meaningful.
                    if !enabled && !(name.as_ref() == b"row" && key == "hidden" && self.zero_height) { continue; }
                    if enabled { "1".to_owned() } else { "0".to_owned() }
                } else {
                    let number: f64 = value.parse().map_err(|_| "Invalid Office axis number".to_owned())?;
                    if !number.is_finite() { return Err("Non-finite Office axis number".into()); }
                    if number == 0.0 { "0".to_owned() } else { number.to_string() }
                };
                attributes.push((key.to_owned(), normalized));
            }
            if name.as_ref() != b"row" || attributes.iter().any(|(key, _)| key != "r") {
                attributes.sort();
                self.hash_text(std::str::from_utf8(name.as_ref()).map_err(|error| error.to_string())?);
                for (key, value) in attributes { self.hash_text("|"); self.hash_text(&key); self.hash_text("="); self.hash_text(&value); }
                self.hash_text(";");
            }
        }
        if !empty { self.path.push(name.as_ref().to_vec()); }
        Ok(())
    }
    pub fn end(&mut self) -> Result<(), String> {
        self.path.pop().ok_or_else(|| "Unmatched Office worksheet closing tag".to_owned())?;
        Ok(())
    }
    pub fn finish(&self) -> String { format!("{:016x}", self.hash) }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn physical_units_and_stored_character_widths_use_one_contract_for_small_and_large_values() {
        assert_eq!(points_to_pixels(15.0), 20.0);
        assert_eq!(pixels_to_points(32.0), 24.0);
        for pixels in [0.0, 7.0, 21.0, 112.0, 1000.0] {
            assert_eq!(stored_character_width_to_pixels(pixels_to_stored_character_width(pixels)), pixels);
        }
    }

    #[test]
    fn fingerprint_ignores_cell_content_attribute_order_and_false_flags_but_tracks_dimensions() {
        let first = r#"<worksheet><sheetFormatPr defaultRowHeight="15" defaultColWidth="16"/><cols><col min="1" max="2" width="20" customWidth="1"/></cols><sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="2" ht="24" customHeight="1"/></sheetData></worksheet>"#;
        let rewritten = r#"<x:worksheet><x:sheetFormatPr defaultColWidth="16.0" customHeight="false" defaultRowHeight="15.0"/><x:cols><x:col customWidth="true" width="20.0" max="2" min="1" hidden="0"/></x:cols><x:sheetData><x:row r="1"><x:c r="A1"><x:v>900</x:v></x:c></x:row><x:row customHeight="true" ht="24.0" r="2"/></x:sheetData></x:worksheet>"#;
        assert_eq!(worksheet_axis_fingerprint(first).unwrap(), worksheet_axis_fingerprint(rewritten).unwrap());
        assert_ne!(worksheet_axis_fingerprint(first).unwrap(), worksheet_axis_fingerprint(&first.replace("ht=\"24\"", "ht=\"30\"")).unwrap());
    }

    #[test]
    fn fingerprint_only_reads_worksheet_axis_paths_and_accepts_the_writer_prefix() {
        let base = r#"<worksheet><sheetFormatPr defaultRowHeight="15"/><sheetData><row r="1" ht="24"/></sheetData>"#;
        let completed = format!(r#"{base}<extLst><ext><sheetFormatPr defaultRowHeight="99"/><cols><col min="1" max="9" width="90"/></cols><row r="7" ht="90"/></ext></extLst></worksheet>"#);
        assert_eq!(worksheet_axis_fingerprint(base).unwrap(), worksheet_axis_fingerprint(&completed).unwrap());
        let invisible = r#"<worksheet><sheetFormatPr defaultRowHeight="0" zeroHeight="1"/><sheetData><row r="2" ht="20"/></sheetData></worksheet>"#;
        assert_ne!(worksheet_axis_fingerprint(invisible).unwrap(), worksheet_axis_fingerprint(&invisible.replace("ht=\"20\"", "ht=\"20\" hidden=\"0\"")).unwrap());
    }
}
