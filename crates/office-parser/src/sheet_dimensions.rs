//! Office source units are projected into CSS pixels at the parser boundary.
//! Exact source measures remain in the DTO so font-dependent character widths
//! are never mistaken for verified physical geometry.

use crate::model::{AxisPayload, AxisSpanPayload};
use rxls::{ImportedAxisMeasure, Sheet};
use serde_json::{Value, json};

pub use office_io_common::MAXIMUM_DIGIT_WIDTH_PIXELS;

pub struct SheetDimensionsPayload {
    pub columns: Vec<AxisPayload>,
    pub rows: Vec<AxisPayload>,
    pub column_spans: Vec<AxisSpanPayload<f64>>,
    pub row_spans: Vec<AxisSpanPayload<f64>>,
    pub row_modes: Vec<AxisSpanPayload<bool>>,
    pub automatic_rows: Vec<AxisPayload>,
    pub default_column: Option<f64>,
    pub default_row: Option<f64>,
    pub default_row_automatic: bool,
    pub source: Value,
}

pub fn points_to_pixels(points: f32) -> f64 { office_io_common::points_to_pixels(f64::from(points)) }

/// ECMA-376 stored character width already includes its padding allowance.
pub fn stored_character_width_to_pixels(width: f64) -> f64 {
    office_io_common::stored_character_width_to_pixels(width)
}

pub fn imported_axis_pixels(measure: ImportedAxisMeasure) -> f64 {
    match measure {
        ImportedAxisMeasure::Twips(value) => f64::from(value) / 15.0,
        ImportedAxisMeasure::MillimeterHundredths(value) => f64::from(value) * 96.0 / 2540.0,
        ImportedAxisMeasure::PointRatio(numerator, denominator) => numerator as f64 / denominator as f64 * 4.0 / 3.0,
        ImportedAxisMeasure::CharacterWidth256(value) | ImportedAxisMeasure::DigitWidth256(value) => stored_character_width_to_pixels(f64::from(value) / 256.0),
        ImportedAxisMeasure::CharacterWidthRatio(numerator, denominator) => stored_character_width_to_pixels(numerator as f64 / denominator as f64),
        ImportedAxisMeasure::CharacterBaseWidth256(value) | ImportedAxisMeasure::DigitBaseWidth256(value) =>
            (f64::from(value) / 256.0 * MAXIMUM_DIGIT_WIDTH_PIXELS + 5.0).floor(),
    }
}

fn source_measure(measure: ImportedAxisMeasure) -> Value {
    match measure {
        ImportedAxisMeasure::Twips(value) => json!({"unit": "twip", "value": value}),
        ImportedAxisMeasure::MillimeterHundredths(value) => json!({"unit": "millimeterHundredth", "value": value}),
        ImportedAxisMeasure::PointRatio(numerator, denominator) => json!({"unit": "point", "numerator": numerator, "denominator": denominator}),
        ImportedAxisMeasure::CharacterWidth256(value) => json!({"unit": "storedCharacter256", "value": value}),
        ImportedAxisMeasure::CharacterWidthRatio(numerator, denominator) => json!({"unit": "storedCharacter", "numerator": numerator, "denominator": denominator}),
        ImportedAxisMeasure::CharacterBaseWidth256(value) => json!({"unit": "baseCharacter256", "value": value}),
        ImportedAxisMeasure::DigitWidth256(value) => json!({"unit": "storedDigit256", "value": value}),
        ImportedAxisMeasure::DigitBaseWidth256(value) => json!({"unit": "baseDigit256", "value": value}),
    }
}

fn compact<T: PartialEq>(entries: impl IntoIterator<Item = (u32, T)>) -> Vec<AxisSpanPayload<T>> {
    let mut result: Vec<AxisSpanPayload<T>> = Vec::new();
    for (index, value) in entries {
        if let Some(previous) = result.last_mut() {
            if previous.end + 1 == index && previous.value == value { previous.end = index; continue; }
        }
        result.push(AxisSpanPayload { start: index, end: index, value });
    }
    result
}

fn numeric_axes(entries: impl IntoIterator<Item = (u32, f64)>) -> (Vec<AxisPayload>, Vec<AxisSpanPayload<f64>>) {
    let mut points = Vec::new();
    let mut spans = Vec::new();
    for span in compact(entries) {
        if span.start == span.end { points.push(AxisPayload { index: span.start, size: span.value }); }
        else { spans.push(span); }
    }
    (points, spans)
}

pub fn sheet_dimensions(sheet: &Sheet) -> SheetDimensionsPayload {
    let column_entries = sheet.column_widths().iter().map(|(&column, &width)| {
        let pixels = sheet.imported_column_axis_measures().get(&column).copied().map(imported_axis_pixels)
            .or_else(|| sheet.physical_column_widths().get(&column).copied().map(points_to_pixels))
            .unwrap_or_else(|| stored_character_width_to_pixels(f64::from(width)));
        (u32::from(column), pixels)
    });
    let (columns, column_spans) = numeric_axes(column_entries);
    let row_entries = sheet.row_heights().iter().filter(|(row, _)| sheet.row_height_is_manual(**row)).map(|(&row, &height)|
        (row, sheet.imported_row_axis_measures().get(&row).copied().map(imported_axis_pixels).unwrap_or_else(|| points_to_pixels(height))));
    let (rows, row_spans) = numeric_axes(row_entries);
    let default_row = sheet.imported_default_row_axis_measure().map(imported_axis_pixels).or_else(|| sheet.default_row_height().map(points_to_pixels));
    let default_column = sheet.imported_default_column_axis_measure().map(imported_axis_pixels)
        .or_else(|| sheet.default_column_width().map(|width| stored_character_width_to_pixels(f64::from(width))));
    let automatic_rows = sheet.row_heights().iter().filter(|(row, _)| !sheet.row_height_is_manual(**row)).filter_map(|(&row, &height)| {
        let size = sheet.imported_row_axis_measures().get(&row).copied().map(imported_axis_pixels).unwrap_or_else(|| points_to_pixels(height));
        (Some(size) != default_row).then_some(AxisPayload { index: row, size })
    }).collect();
    let row_modes = compact(sheet.row_heights().keys().map(|&row| (row, !sheet.row_height_is_manual(row))));
    let source = json!({
        "rowUnit": "point", "columnUnit": "storedCharacter", "physicalColumnUnit": "point",
        "maximumDigitWidthPixels": MAXIMUM_DIGIT_WIDTH_PIXELS, "fontMetricVerified": false,
        "defaultRowHeight": sheet.default_row_height(), "defaultColumnWidth": sheet.default_column_width(),
        "defaultRowAxisMeasure": sheet.imported_default_row_axis_measure().map(source_measure),
        "defaultColumnAxisMeasure": sheet.imported_default_column_axis_measure().map(source_measure),
        "rowHeights": compact(sheet.row_heights().iter().map(|(&row, &size)| (row, size))),
        "columnWidths": compact(sheet.column_widths().iter().map(|(&column, &size)| (u32::from(column), size))),
        "physicalColumnWidths": compact(sheet.physical_column_widths().iter().map(|(&column, &size)| (u32::from(column), size))),
        "rowAxisMeasures": compact(sheet.imported_row_axis_measures().iter().map(|(&row, &measure)| (row, source_measure(measure)))),
        "columnAxisMeasures": compact(sheet.imported_column_axis_measures().iter().map(|(&column, &measure)| (u32::from(column), source_measure(measure)))),
    });
    SheetDimensionsPayload { columns, rows, column_spans, row_spans, row_modes, automatic_rows,
        default_column, default_row, default_row_automatic: !sheet.default_row_height_is_manual(), source }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn converts_physical_source_measures_and_stored_widths_without_a_small_width_heuristic() {
        assert_eq!(points_to_pixels(15.0), 20.0);
        assert_eq!(imported_axis_pixels(ImportedAxisMeasure::Twips(300)), 20.0);
        assert_eq!(imported_axis_pixels(ImportedAxisMeasure::MillimeterHundredths(2540)), 96.0);
        assert_eq!(imported_axis_pixels(ImportedAxisMeasure::CharacterWidth256(16 * 256)), 112.0);
        assert_eq!(stored_character_width_to_pixels(1.0), 7.0);
        assert_eq!(stored_character_width_to_pixels(0.0), 0.0);
        assert_eq!(imported_axis_pixels(ImportedAxisMeasure::CharacterBaseWidth256(8 * 256)), 61.0);
    }

    #[test]
    fn compacts_equal_adjacent_sizes_without_filling_gaps_or_losing_zero() {
        let (points, spans) = numeric_axes([(0, 0.0), (1, 0.0), (3, 20.0), (4, 30.0)]);
        assert_eq!(spans.len(), 1);
        assert_eq!((spans[0].start, spans[0].end, spans[0].value), (0, 1, 0.0));
        assert_eq!(points.iter().map(|point| (point.index, point.size)).collect::<Vec<_>>(), [(3, 20.0), (4, 30.0)]);
    }
}
