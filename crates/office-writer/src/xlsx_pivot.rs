use serde::Deserialize;
use serde_json::{Map, Value};

use crate::ExportError;

#[derive(Debug, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct PivotTableInput {
    pub(crate) id: String,
    pub(crate) source: PivotSourceInput,
    pub(crate) anchor: PivotAnchorInput,
    pub(crate) fields: Vec<PivotFieldInput>,
    #[serde(default)]
    pub(crate) options: PivotOptionsInput,
    #[serde(default)]
    pub(crate) output_range: Option<RangeInput>,
}

#[derive(Debug, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct PivotSourceInput {
    #[serde(default)]
    pub(crate) sub_unit_id: String,
    #[serde(default)]
    pub(crate) sheet_name: String,
    #[serde(default)]
    pub(crate) range: RangeInput,
}

#[derive(Debug, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct PivotAnchorInput {
    #[serde(default)]
    pub(crate) sub_unit_id: String,
    #[serde(default)]
    pub(crate) row: u32,
    #[serde(default)]
    pub(crate) col: u32,
}

#[derive(Debug, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct PivotFieldInput {
    #[serde(default)]
    pub(crate) source_column: u32,
    #[serde(default)]
    pub(crate) name: String,
    #[serde(default)]
    pub(crate) display_name: Option<String>,
    #[serde(default)]
    pub(crate) area: String,
    #[serde(default)]
    pub(crate) position: usize,
    #[serde(default)]
    pub(crate) aggregation: Option<String>,
}

#[derive(Debug, Deserialize, Clone, Default)]
#[serde(rename_all = "camelCase")]
pub(crate) struct PivotOptionsInput {
    #[serde(default = "default_true")]
    pub(crate) show_row_grand_total: bool,
    #[serde(default = "default_true")]
    pub(crate) show_col_grand_total: bool,
    #[serde(default = "default_true")]
    pub(crate) show_row_subtotals: bool,
    #[serde(default = "default_true")]
    pub(crate) show_col_subtotals: bool,
    #[serde(default)]
    pub(crate) repeat_row_labels: bool,
}

#[derive(Debug, Deserialize, Clone)]
#[serde(rename_all = "camelCase")]
pub(crate) struct RangeInput {
    #[serde(default)]
    pub(crate) start_row: Option<u32>,
    #[serde(default)]
    pub(crate) start_col: Option<u32>,
    #[serde(default)]
    pub(crate) end_row: Option<u32>,
    #[serde(default)]
    pub(crate) end_col: Option<u32>,
    #[serde(default)]
    pub(crate) first_row: Option<u32>,
    #[serde(default)]
    pub(crate) first_col: Option<u32>,
    #[serde(default)]
    pub(crate) last_row: Option<u32>,
    #[serde(default)]
    pub(crate) last_col: Option<u32>,
}

impl Default for RangeInput {
    fn default() -> Self {
        Self {
            start_row: None,
            start_col: None,
            end_row: None,
            end_col: None,
            first_row: None,
            first_col: None,
            last_row: None,
            last_col: None,
        }
    }
}

fn default_true() -> bool {
    true
}

impl RangeInput {
    pub(crate) fn bounds(&self) -> Option<(u32, u32, u32, u32)> {
        let row = self.start_row.or(self.first_row)?;
        let col = self.start_col.or(self.first_col)?;
        Some((
            row,
            col,
            self.end_row.or(self.last_row).unwrap_or(row),
            self.end_col.or(self.last_col).unwrap_or(col),
        ))
    }
}

#[derive(Debug, Clone)]
pub(crate) struct PivotParts {
    pub(crate) cache_definition: String,
    pub(crate) cache_records: String,
    pub(crate) table_definition: String,
    #[cfg(test)]
    pub(crate) output_reference: String,
    #[cfg(test)]
    pub(crate) source_reference: String,
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
enum CacheValue {
    Missing,
    String(String),
    Number(String),
    Boolean(bool),
    Error(String),
}

fn cache_value(value: Option<Value>) -> CacheValue {
    match value {
        None | Some(Value::Null) => CacheValue::Missing,
        Some(Value::String(value)) => CacheValue::String(value),
        Some(Value::Number(value)) => CacheValue::Number(value.to_string()),
        Some(Value::Bool(value)) => CacheValue::Boolean(value),
        Some(Value::Object(map)) if map.contains_key("error") => CacheValue::Error(
            map.get("error")
                .and_then(Value::as_str)
                .unwrap_or("#VALUE!")
                .to_owned(),
        ),
        Some(value) => CacheValue::String(value.to_string()),
    }
}

fn cache_item(value: &CacheValue) -> String {
    match value {
        CacheValue::Missing => "<m/>".to_owned(),
        CacheValue::String(value) => format!("<s v=\"{}\"/>", crate::xml::escape(value)),
        CacheValue::Number(value) => format!("<n v=\"{}\"/>", crate::xml::escape(value)),
        CacheValue::Boolean(value) => format!("<b v=\"{}\"/>", if *value { 1 } else { 0 }),
        CacheValue::Error(value) => format!("<e v=\"{}\"/>", crate::xml::escape(value)),
    }
}

fn cell_value(cells: &Map<String, Value>, row: u32, col: u32) -> Option<Value> {
    let key = format!("{row}:{col}");
    let entry = cells
        .get(&key)
        .or_else(|| cells.get(&format!("{}{}", column_name(col), row + 1)))?;
    if let Some(map) = entry.as_object() {
        return map
            .get("value")
            .cloned()
            .or_else(|| map.get("rawValue").cloned());
    }
    Some(entry.clone())
}

fn column_name(mut col: u32) -> String {
    let mut out = String::new();
    loop {
        let digit = (col % 26) as u8;
        out.insert(0, (b'A' + digit) as char);
        if col < 26 {
            break;
        }
        col = col / 26 - 1;
    }
    out
}

fn cell_ref(row: u32, col: u32) -> String {
    format!("{}{}", column_name(col), row + 1)
}

fn range_ref(bounds: (u32, u32, u32, u32)) -> String {
    let (r0, c0, r1, c1) = bounds;
    format!("{}:{}", cell_ref(r0, c0), cell_ref(r1, c1))
}

fn field_area_tag(area: &str) -> Option<&'static str> {
    match area {
        "row" => Some("axisRow"),
        "column" => Some("axisCol"),
        "filter" => Some("axisPage"),
        _ => None,
    }
}

fn subtotal(aggregation: Option<&str>) -> &'static str {
    match aggregation.unwrap_or("sum") {
        "count" => "count",
        "average" => "average",
        "min" => "min",
        "max" => "max",
        _ => "sum",
    }
}

pub(crate) fn build_parts(
    pivot: &PivotTableInput,
    source_name: &str,
    output_bounds: (u32, u32, u32, u32),
    source_cells: &Map<String, Value>,
    source_bounds: (u32, u32, u32, u32),
    table_name: &str,
    cache_id: usize,
) -> Result<PivotParts, ExportError> {
    let (start_row, start_col, end_row, end_col) = source_bounds;
    if end_row <= start_row || end_col < start_col {
        return Err(ExportError::invalid(format!(
            "pivot table {} source must include a header and data rows",
            pivot.id
        )));
    }
    let width = (end_col - start_col + 1) as usize;
    if pivot.fields.len() != width {
        return Err(ExportError::invalid(format!(
            "pivot table {} field count does not match its source width",
            pivot.id
        )));
    }
    let mut cache = Vec::<Vec<CacheValue>>::new();
    let mut unique = vec![Vec::<CacheValue>::new(); width];
    let mut indexes = vec![std::collections::HashMap::<CacheValue, usize>::new(); width];
    for row in start_row + 1..=end_row {
        let mut record = Vec::with_capacity(width);
        for offset in 0..width {
            let value = cache_value(cell_value(source_cells, row, start_col + offset as u32));
            if !indexes[offset].contains_key(&value) {
                let index = unique[offset].len();
                indexes[offset].insert(value.clone(), index);
                unique[offset].push(value.clone());
            }
            record.push(value);
        }
        cache.push(record);
    }

    let mut definition = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><pivotCacheDefinition xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" recordCount=\"{}\" refreshOnLoad=\"1\" saveData=\"1\" r:id=\"rIdRecords\"><cacheSource type=\"worksheet\"><worksheetSource ref=\"{}\" sheet=\"{}\"/></cacheSource><cacheFields count=\"{}\">",
        cache.len(),
        range_ref(source_bounds),
        crate::xml::escape(source_name),
        width
    );
    for (offset, field) in pivot.fields.iter().enumerate() {
        let name = cell_value(source_cells, start_row, start_col + offset as u32)
            .and_then(|value| value.as_str().map(str::to_owned))
            .unwrap_or_else(|| field.name.clone());
        definition.push_str(&format!(
            "<cacheField name=\"{}\"><sharedItems count=\"{}\">",
            crate::xml::escape(&name),
            unique[offset].len()
        ));
        for value in &unique[offset] {
            definition.push_str(&cache_item(value));
        }
        definition.push_str("</sharedItems></cacheField>");
    }
    definition.push_str("</cacheFields></pivotCacheDefinition>");

    let mut records = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><pivotCacheRecords xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"{}\">",
        cache.len()
    );
    for row in &cache {
        records.push_str("<r>");
        for (field, value) in row.iter().enumerate() {
            if let Some(index) = indexes[field].get(value) {
                records.push_str(&format!("<x v=\"{}\"/>", index));
            } else {
                return Err(ExportError::invalid(
                    "pivot cache item index was not generated",
                ));
            }
        }
        records.push_str("</r>");
    }
    records.push_str("</pivotCacheRecords>");

    let mut fields = vec![None::<&PivotFieldInput>; width];
    for field in &pivot.fields {
        let offset = field
            .source_column
            .checked_sub(start_col)
            .filter(|value| *value < width as u32)
            .ok_or_else(|| {
                ExportError::invalid(format!(
                    "pivot table {} field source is outside its range",
                    pivot.id
                ))
            })? as usize;
        fields[offset] = Some(field);
    }
    let mut table = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><pivotTableDefinition xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" name=\"{}\" cacheId=\"{}\" dataCaption=\"Values\" grandTotalCaption=\"Grand total\" rowGrandTotals=\"{}\" colGrandTotals=\"{}\" compact=\"{}\" multipleFieldFilters=\"0\" showHeaders=\"1\" useAutoFormatting=\"1\" applyNumberFormats=\"1\" applyBorderFormats=\"1\" applyFontFormats=\"1\" applyPatternFormats=\"1\" applyAlignmentFormats=\"1\" applyWidthHeightFormats=\"1\"><location ref=\"{}\" firstHeaderRow=\"1\" firstDataRow=\"2\" firstDataCol=\"1\"/><pivotFields count=\"{}\">",
        crate::xml::escape(table_name),
        cache_id,
        u8::from(pivot.options.show_row_grand_total),
        u8::from(pivot.options.show_col_grand_total),
        u8::from(!pivot.options.repeat_row_labels),
        range_ref(output_bounds),
        width
    );
    for (offset, field) in fields.iter().enumerate() {
        let field = field.ok_or_else(|| {
            ExportError::invalid(format!(
                "pivot table {} has a missing source field",
                pivot.id
            ))
        })?;
        let axis = field_area_tag(&field.area)
            .map(|axis| format!(" axis=\"{axis}\""))
            .unwrap_or_default();
        let default_subtotal = match field.area.as_str() {
            "row" if !pivot.options.show_row_subtotals => " defaultSubtotal=\"0\"",
            "column" if !pivot.options.show_col_subtotals => " defaultSubtotal=\"0\"",
            _ => "",
        };
        let data = if field.area == "value" {
            " dataField=\"1\""
        } else {
            ""
        };
        table.push_str(&format!(
            "<pivotField{}{}{} showAll=\"0\"><items count=\"{}\">",
            axis,
            default_subtotal,
            data,
            unique[offset].len()
        ));
        for index in 0..unique[offset].len() {
            table.push_str(&format!("<item x=\"{index}\"/>"));
        }
        table.push_str("</items></pivotField>");
    }
    table.push_str("</pivotFields>");
    for (area, node) in [("row", "rowFields"), ("column", "colFields")] {
        let mut area_fields = pivot
            .fields
            .iter()
            .filter(|field| field.area == area)
            .collect::<Vec<_>>();
        area_fields.sort_by_key(|field| field.position);
        table.push_str(&format!("<{node} count=\"{}\">", area_fields.len()));
        for field in area_fields {
            table.push_str(&format!(
                "<field x=\"{}\"/>",
                field.source_column - start_col
            ));
        }
        table.push_str(&format!("</{node}>"));
        let item_node = if area == "row" {
            "rowItems"
        } else {
            "colItems"
        };
        table.push_str(&format!("<{item_node} count=\"0\"/>"));
    }
    let mut page_fields = pivot
        .fields
        .iter()
        .filter(|field| field.area == "filter")
        .collect::<Vec<_>>();
    page_fields.sort_by_key(|field| field.position);
    table.push_str(&format!("<pageFields count=\"{}\">", page_fields.len()));
    for field in page_fields {
        table.push_str(&format!(
            "<pageField fld=\"{}\"/>",
            field.source_column - start_col
        ));
    }
    table.push_str("</pageFields>");
    let mut data_fields = pivot
        .fields
        .iter()
        .filter(|field| field.area == "value")
        .collect::<Vec<_>>();
    data_fields.sort_by_key(|field| field.position);
    table.push_str(&format!("<dataFields count=\"{}\">", data_fields.len()));
    for field in data_fields {
        let caption = field.display_name.as_deref().unwrap_or(&field.name);
        table.push_str(&format!(
            "<dataField name=\"{}\" fld=\"{}\" subtotal=\"{}\"/>",
            crate::xml::escape(caption),
            field.source_column - start_col,
            subtotal(field.aggregation.as_deref())
        ));
    }
    table.push_str("</dataFields><pivotTableStyleInfo name=\"PivotStyleMedium9\" showRowHeaders=\"1\" showColHeaders=\"1\" showRowStripes=\"0\" showColStripes=\"0\"/></pivotTableDefinition>");

    Ok(PivotParts {
        cache_definition: definition,
        cache_records: records,
        table_definition: table,
        #[cfg(test)]
        output_reference: range_ref(output_bounds),
        #[cfg(test)]
        source_reference: range_ref(source_bounds),
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn pivot_parts_share_source_fields_and_record_indexes() {
        let pivot: PivotTableInput = serde_json::from_str(r#"{"id":"pivot-1","source":{"subUnitId":"data","sheetName":"Data","range":{"startRow":0,"startCol":0,"endRow":2,"endCol":1}},"anchor":{"subUnitId":"pivot","row":0,"col":0},"fields":[{"sourceColumn":0,"name":"Region","area":"row","position":0},{"sourceColumn":1,"name":"Sales","area":"value","position":0}],"options":{},"outputRange":{"startRow":0,"startCol":0,"endRow":2,"endCol":1}}"#).unwrap();
        let cells = serde_json::from_value(serde_json::json!({
            "A1":{"value":"Region"},"B1":{"value":"Sales"},
            "A2":{"value":"East"},"B2":{"value":10},
            "A3":{"value":"West"},"B3":{"value":20}
        }))
        .unwrap();
        let parts = build_parts(
            &pivot,
            "Data",
            (0, 0, 2, 1),
            &cells,
            (0, 0, 2, 1),
            "PivotTable1",
            1,
        )
        .unwrap();
        assert!(
            parts
                .cache_definition
                .contains("<cacheField name=\"Region\"")
        );
        assert!(
            parts
                .cache_records
                .contains("<r><x v=\"0\"/><x v=\"0\"/></r>")
        );
        assert!(
            parts
                .cache_records
                .contains("<r><x v=\"1\"/><x v=\"1\"/></r>")
        );
        assert!(
            parts
                .table_definition
                .contains("<rowFields count=\"1\"><field x=\"0\"/>")
        );
        assert!(
            parts
                .table_definition
                .contains("<dataField name=\"Sales\" fld=\"1\" subtotal=\"sum\"/>")
        );
        assert_eq!(parts.source_reference, "A1:B3");
        assert_eq!(parts.output_reference, "A1:B3");
    }
}
