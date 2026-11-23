//! Formula-backed native helpers for authored sources that do not map to one
//! traditional chart vector. Caches never substitute for an Office binding.

use std::collections::{BTreeMap, BTreeSet};
use serde_json::{Map, Value, json};
use crate::{ExportError, ExportWarning};
use super::{ChartInput, Context, NativeBindings, NativeCategory, NativeGroup, NativeHelper, NativeVector, SheetInput, WorkbookInput,
    context, native_chart_kind, numeric, scalar, text, chart_range_formula, cell_ref, MAX_SERIES, source_view};
use super::series_plan::planning_specifications;
use super::dates;

const MAX_HELPER_CELLS: usize = 1_000_000;
const MAX_HELPER_TEXT: usize = 64 * 1024 * 1024;
const MARKER: &str = "office-studio:chart-helper:v1:";

pub(super) fn prepare(workbook: &mut WorkbookInput, diagnostics: &mut Vec<ExportWarning>) -> Result<(), ExportError> {
    // Only small authored controls are cloned. The source worksheet/cell maps
    // remain borrowed, and a completed helper takes ownership of its buffers.
    // A regenerated helper must read the current source, including when a
    // previously empty chart gains visible data again or no longer needs one.
    // Clear the actual targets so the direct-binding path cannot reuse a
    // helper from an earlier preparation of this workbook.
    if let Some(charts) = &mut workbook.charts {
        for chart in charts { chart.native_bindings = None; }
    }
    let charts = workbook.charts.clone().unwrap_or_default();
    let mut cells_used = 0usize;
    let mut text_used = 0usize;
    for (index, chart) in charts.iter().enumerate() {
        super::warnings::append(chart, diagnostics)?;
        let Some((replace, sheet, bindings)) = plan(workbook, chart, index, &mut cells_used, &mut text_used, diagnostics)? else { continue; };
        if let Some(replace) = replace { workbook.sheets[replace] = sheet; }
        else { workbook.sheets.push(sheet); }
        if let Some(target) = workbook.charts.as_mut().and_then(|charts| charts.get_mut(index)) { target.native_bindings = Some(bindings); }
    }
    Ok(())
}

fn plan(workbook: &WorkbookInput, chart: &ChartInput, index: usize, cells_used: &mut usize, text_used: &mut usize,
    diagnostics: &mut Vec<ExportWarning>) -> Result<Option<(Option<usize>, SheetInput, NativeBindings)>, ExportError> {
    let native = native_chart_kind(if chart.chart_type == "combination" { "column" } else { &chart.chart_type });
    let Some(native) = native else { return Ok(None); };
    let date_mode = dates::enabled(chart, native);
    if date_mode { dates::validate_axis(chart, workbook.date1904)?; }
    let source = context(workbook, chart)?;
    let view = source_view::read(workbook, chart, &source)?;
    let specs = planning_specifications(&source, chart)?;
    let selected_category = chart.options.mapping.as_ref().and_then(|mapping| mapping.category_indexes.as_deref()).unwrap_or(&[]);
    let category_fields = if selected_category.is_empty() || native.plot == "stockChart" { vec![chart.category_field] } else { selected_category.to_vec() };
    let category_plot = !matches!(native.plot, "scatterChart" | "bubbleChart");
    let multi_category = category_plot && category_fields.as_slice() != [chart.category_field];
    let bubble = chart.options.mapping.as_ref().and_then(|mapping| mapping.bubble.as_ref());
    let grouped = chart.chart_type == "bubble" && bubble.is_some_and(|mapping| mapping.series_index.is_some());
    let zero = chart.options.invalid_value_strategy.as_deref() == Some("zero");
    let default_size = specs.iter().any(|spec| spec.kind.plot == "bubbleChart" && spec.size_field.is_none());
    let hidden_fields = view.fields.as_ref().is_some_and(|fields| fields.len() < source.field_count);
    let hidden_points = view.points.len() != source.point_count;
    if !date_mode && !multi_category && !grouped && !zero && !default_size && !hidden_fields && !hidden_points { return Ok(None); }
    let visible = |field: usize| view.fields.as_ref().is_none_or(|fields| fields.contains(&field));
    let mut primary = if category_plot { category_fields.clone() } else { Vec::new() };
    if native.plot == "stockChart" { primary.extend(specs.iter().map(|spec| spec.field)); }
    if native.plot == "bubbleChart" {
        for spec in &specs { primary.extend([Some(spec.field), spec.x_field, spec.size_field].into_iter().flatten()); }
        if let Some(mapping) = bubble { primary.extend([mapping.label_index, mapping.series_index].into_iter().flatten()); }
    }
    for field in &primary { source.coordinates(chart, *field, None)?; }
    let empty = view.points.is_empty() || primary.iter().any(|field| !visible(*field))
        || specs.iter().all(|spec| !visible(spec.field) || spec.x_field.is_some_and(|field| !visible(field)));
    if date_mode && category_plot && category_fields.len() != 1 && !empty {
        return Err(ExportError::invalid("CHART_DATE_AXIS_CATEGORY: an elapsed-date category requires one unambiguous source date field"));
    }
    let mut numeric_fields = BTreeSet::<(usize, bool)>::new();
    let mut date_fields = BTreeSet::<usize>::new();
    for spec in &specs {
        numeric_fields.insert((spec.field, zero));
        if let Some(field) = spec.x_field {
            if date_mode { date_fields.insert(field); } else { numeric_fields.insert((field, false)); }
        }
        if let Some(field) = spec.size_field { numeric_fields.insert((field, false)); }
        for error in chart.options.error_bars.iter().filter(|error| error.series_index == spec.field && error.value_type == "cust") {
            for field in [if error.kind == "minus" { None } else { error.plus_index }, if error.kind == "plus" { None } else { error.minus_index }].into_iter().flatten() {
                source.coordinates(chart, field, None)?; numeric_fields.insert((field, false));
            }
        }
    }
    if category_plot {
        let mut seen = BTreeSet::new();
        for field in &category_fields {
            source.coordinates(chart, *field, None)?;
            if !seen.insert(*field) { return Err(ExportError::invalid("chart category fields must be distinct")); }
        }
    }
    let mut groups = Vec::<String>::new();
    let mut group_at_point = Vec::<usize>::new();
    if grouped && !empty {
        let mapping = bubble.ok_or_else(|| ExportError::invalid("grouped bubble mapping is missing"))?;
        let series_field = mapping.series_index.ok_or_else(|| ExportError::invalid("bubble group field is missing"))?;
        let mut group_ids = BTreeMap::<String, usize>::new();
        group_at_point.try_reserve(view.points.len()).map_err(|_| budget_error())?;
        for point in &view.points {
            let (row, column) = source.coordinates(chart, series_field, Some(*point))?;
            let name = text(scalar(source.sheet, row, column));
            let group_index = if let Some(index) = group_ids.get(&name) { *index } else {
                if groups.len() >= MAX_SERIES { return Err(ExportError::invalid("grouped bubble series exceed the native chart budget")); }
                let index = groups.len(); group_ids.insert(name.clone(), index); groups.push(name); index
            };
            group_at_point.push(group_index);
        }
    }
    let vector_count = if empty { 0 } else {
        (if grouped { groups.len() } else { 1 }).checked_mul(numeric_fields.len() + date_fields.len() + usize::from(default_size))
            .and_then(|count| count.checked_add(usize::from(category_plot))).ok_or_else(budget_error)?
    };
    let columns = vector_count.checked_add(1).filter(|columns| *columns <= 16_384).ok_or_else(budget_error)?;
    let cell_count = (view.points.len() + 1).checked_mul(vector_count).and_then(|count| count.checked_add(1)).ok_or_else(budget_error)?;
    *cells_used = cells_used.checked_add(cell_count).filter(|count| *count <= MAX_HELPER_CELLS).ok_or_else(budget_error)?;
    let (replace, id, name) = helper_identity(workbook, chart, index);
    let mut cells = Map::new();
    let marker = format!("{MARKER}{}", chart.id);
    add_value_text_budget(text_used, &marker)?;
    cells.insert("0:0".into(), json!({"value":marker,"dataType":"string"}));
    let mut bindings = NativeBindings { helper: Some(NativeHelper { sheet_name: name.clone(), key: chart.id.clone() }),
        visible_fields: view.fields.clone(), point_count: view.points.len(), empty, ..Default::default() };
    let mut column = 1u32;
    if category_plot && !empty {
        if date_mode {
            let field = category_fields[0];
            let vector = formula_vector(&view.points, &name, column, "Date", &mut cells, text_used, |_output_index, point| {
                let (row, source_column) = source.coordinates(chart, field, Some(point))?;
                let parsed = dates::read(scalar(source.sheet, row, source_column), workbook.date1904)?;
                Ok((dates::expression(&source_ref(&source, row, source_column), parsed, workbook.date1904), parsed.map(|date| date.serial)))
            })?;
            let mut values = Vec::new(); values.try_reserve(view.points.len()).map_err(|_| budget_error())?;
            for point in &view.points {
                let (row, source_column) = source.coordinates(chart, field, Some(*point))?;
                let value = text(scalar(source.sheet, row, source_column));
                add_value_text_budget(text_used, &value)?; values.push(value);
            }
            bindings.category = Some(NativeCategory { formula: vector.formula, values, dates: Some(vector.values) });
        } else {
        let mut values = Vec::new(); values.try_reserve(view.points.len()).map_err(|_| budget_error())?;
        cells.insert(format!("0:{column}"), json!({"value":"Category","dataType":"string"}));
        let mut numeric_categories = false;
        for (output_index, point) in view.points.iter().enumerate() {
            let mut parts = Vec::new();
            let mut labels = Vec::new();
            for field in &category_fields {
                let (row, source_column) = source.coordinates(chart, *field, Some(*point))?;
                let value = scalar(source.sheet, row, source_column);
                numeric_categories |= value.is_some_and(Value::is_number);
                labels.push(text(value)); parts.push(string_expression(&source_ref(&source, row, source_column)));
            }
            let value = labels.join(" / ");
            let formula = parts.join("&\" / \"&");
            add_text_budget(text_used, formula.len())?;
            add_value_text_budget(text_used, &value)?;
            values.push(value.clone());
            cells.insert(format!("{}:{column}", output_index + 1), json!({"formula":formula,"value":value,"dataType":"string","cachedDataType":"string"}));
        }
        bindings.category = Some(NativeCategory { formula: chart_range_formula(&name, 1, column, view.points.len() as u32, column), values, dates: None });
        if numeric_categories { diagnostics.push(ExportWarning::new("CHART_CATEGORY_TEXT_NATIVE_FORMAT", format!("Chart {} joins numeric category labels with Office's current text conversion; the exact authored field order is retained", chart.id))); }
        }
        column += 1;
    }
    if grouped && !empty {
        let mapping = bubble.ok_or_else(|| ExportError::invalid("bubble mapping is missing"))?;
        let series_field = mapping.series_index.ok_or_else(|| ExportError::invalid("bubble group field is missing"))?;
        for (group_index, group_name) in groups.into_iter().enumerate() {
            let key_cell = format!("{}{}$1", "$", cell_ref(0, column).trim_end_matches('1'));
            let mut vectors = BTreeMap::new();
            for (field, as_zero) in &numeric_fields {
                let vector = formula_vector(&view.points, &name, column, &group_name, &mut cells, text_used, |output_index, point| {
                    let (row, source_column) = source.coordinates(chart, *field, Some(point))?;
                    let (group_row, group_column) = source.coordinates(chart, series_field, Some(point))?;
                    let group_reference = source_ref(&source, group_row, group_column);
                    let active = group_at_point[output_index] == group_index && visible(*field);
                    let value = if active { numeric(scalar(source.sheet, row, source_column)).or(if *as_zero { Some(0.0) } else { None }) } else { None };
                    let matching = group_expression(&group_reference, &key_cell, &group_name);
                    let numeric = numeric_expression(&source_ref(&source, row, source_column), *as_zero);
                    Ok((if visible(*field) { format!("IF({matching},{numeric},NA())") } else { "NA()".into() }, value))
                })?;
                vectors.insert((*field, *as_zero), vector);
                column += 1;
            }
            let mut date_vectors = BTreeMap::new();
            for field in &date_fields {
                let vector = formula_vector(&view.points, &name, column, &group_name, &mut cells, text_used, |output_index, point| {
                    let (row, source_column) = source.coordinates(chart, *field, Some(point))?;
                    let (group_row, group_column) = source.coordinates(chart, series_field, Some(point))?;
                    let parsed = dates::read(scalar(source.sheet, row, source_column), workbook.date1904)?;
                    let matching = group_expression(&source_ref(&source, group_row, group_column), &key_cell, &group_name);
                    let expression = dates::expression(&source_ref(&source, row, source_column), parsed, workbook.date1904);
                    Ok((format!("IF({matching},{expression},NA())"), if group_at_point[output_index] == group_index { parsed.map(|date| date.serial) } else { None }))
                })?;
                date_vectors.insert(*field, vector); column += 1;
            }
            let size = if default_size {
                let vector = formula_vector(&view.points, &name, column, &group_name, &mut cells, text_used, |output_index, point| {
                    let (row, source_column) = source.coordinates(chart, series_field, Some(point))?;
                    Ok((format!("IF({},1,NA())", group_expression(&source_ref(&source, row, source_column), &key_cell, &group_name)),
                        (group_at_point[output_index] == group_index).then_some(1.0)))
                })?;
                column += 1; Some(vector)
            } else { None };
            bindings.groups.push(NativeGroup { name: group_name, vectors, date_vectors, default_size: size });
        }
        diagnostics.push(ExportWarning::new("CHART_BUBBLE_GROUP_BINDINGS", format!("Chart {} uses formula-backed native series for its current groups; a new group requires refreshing the chart in the workbench", chart.id)));
    } else if !empty {
        for (field, as_zero) in &numeric_fields {
            let (_, header) = source.header(chart, *field)?;
            let vector = formula_vector(&view.points, &name, column, &header, &mut cells, text_used, |_output_index, point| {
                let (row, source_column) = source.coordinates(chart, *field, Some(point))?;
                let value = if visible(*field) { numeric(scalar(source.sheet, row, source_column)).or(if *as_zero { Some(0.0) } else { None }) } else { None };
                Ok((if visible(*field) { numeric_expression(&source_ref(&source, row, source_column), *as_zero) } else { "NA()".into() }, value))
            })?;
            bindings.vectors.insert((*field, *as_zero), vector); column += 1;
        }
        for field in &date_fields {
            let (_, header) = source.header(chart, *field)?;
            let vector = formula_vector(&view.points, &name, column, &header, &mut cells, text_used, |_output_index, point| {
                let (row, source_column) = source.coordinates(chart, *field, Some(point))?;
                let parsed = dates::read(scalar(source.sheet, row, source_column), workbook.date1904)?;
                Ok((dates::expression(&source_ref(&source, row, source_column), parsed, workbook.date1904), parsed.map(|date| date.serial)))
            })?;
            bindings.date_vectors.insert(*field, vector); column += 1;
        }
        if default_size { bindings.default_size = Some(formula_vector(&view.points, &name, column, "Size", &mut cells, text_used, |_output_index, _point| Ok(("1".into(), Some(1.0))))?); }
    }
    if date_mode && !empty {
        let valid_date = bindings.category.as_ref().and_then(|category| category.dates.as_ref()).is_some_and(|values| values.iter().any(Option::is_some))
            || bindings.date_vectors.values().any(|vector| vector.values.iter().any(Option::is_some))
            || bindings.groups.iter().any(|group| group.date_vectors.values().any(|vector| vector.values.iter().any(Option::is_some)));
        // Only a series' Y vector makes a date axis necessary. A size or
        // custom-error column can contain numbers while all plotted Y values
        // are missing, and must not turn that empty plot into a date error.
        let has_values = specs.iter().any(|spec| {
            bindings.vectors.get(&(spec.field, zero)).is_some_and(|vector| vector.values.iter().any(Option::is_some))
                || bindings.groups.iter().any(|group| group.vectors.get(&(spec.field, zero))
                    .is_some_and(|vector| vector.values.iter().any(Option::is_some)))
        });
        if !valid_date && has_values { return Err(ExportError::invalid("CHART_DATE_AXIS_EMPTY: plotted values require at least one valid date in their source")); }
        diagnostics.push(ExportWarning::new("CHART_NATIVE_DATE_BINDINGS", format!("Chart {} keeps formula-backed date serials in the workbook's date system; date text syntax changes require refreshing the helper in the workbench, and numeric text follows Office's locale conversion", chart.id)));
    }
    if chart.options.plot_visible_only.unwrap_or(true) { diagnostics.push(ExportWarning::new("CHART_HELPER_SOURCE_VISIBILITY", format!("Chart {} excludes the source's current hidden data; subsequent source visibility changes must be refreshed in the workbench", chart.id))); }
    let sheet = SheetInput { id, name, visible: false, cells, axis_unit: Some("cssPixel".into()), show_gridlines: false,
        row_count: Some(view.points.len() as u32 + 1), column_count: Some(columns as u32),
        metadata: Some(json!({"officeChartHelper":{"version":1,"chartId":chart.id,"key":chart.id}})), ..Default::default() };
    Ok(Some((replace, sheet, bindings)))
}

fn formula_vector(points: &[usize], sheet_name: &str, column: u32, header: &str, cells: &mut Map<String, Value>, text_used: &mut usize,
    mut calculate: impl FnMut(usize, usize) -> Result<(String, Option<f64>), ExportError>) -> Result<NativeVector, ExportError> {
    add_value_text_budget(text_used, header)?;
    cells.insert(format!("0:{column}"), json!({"value":header,"dataType":"string"}));
    let mut values = Vec::new(); values.try_reserve(points.len()).map_err(|_| budget_error())?;
    for (output_index, point) in points.iter().enumerate() {
        let (formula, value) = calculate(output_index, *point)?;
        add_text_budget(text_used, formula.len())?;
        values.push(value);
        let (cached, kind) = value.map(|value| (json!(value), "number")).unwrap_or((json!("#N/A"), "error"));
        cells.insert(format!("{}:{column}", output_index + 1), json!({"formula":formula,"value":cached,"dataType":kind,"cachedDataType":kind}));
    }
    Ok(NativeVector { formula: chart_range_formula(sheet_name, 1, column, points.len() as u32, column), values })
}

fn helper_identity(workbook: &WorkbookInput, chart: &ChartInput, index: usize) -> (Option<usize>, String, String) {
    let source_id = chart.source.sheet_id.as_deref().unwrap_or(&chart.sheet_id);
    if let Some((index, sheet)) = workbook.sheets.iter().enumerate().find(|(_, sheet)| !sheet.visible && sheet.id != chart.sheet_id && sheet.id != source_id
        && sheet.name.starts_with("__ov_chart_") && sheet.metadata.as_ref().and_then(|metadata| metadata.get("officeChartHelper"))
            .is_some_and(|marker| marker.get("version").and_then(Value::as_u64) == Some(1)
                && marker.get("chartId").and_then(Value::as_str) == Some(chart.id.as_str())
                && marker.get("key").and_then(Value::as_str).is_some_and(|key|
                    sheet.cells.get("0:0").or_else(|| sheet.cells.get("A1")).and_then(|cell| cell.get("value")).and_then(Value::as_str) == Some(format!("{MARKER}{key}").as_str())))) {
        return (Some(index), sheet.id.clone(), sheet.name.clone());
    }
    let names = super::safe_sheet_names(workbook.sheets.iter().map(|sheet| sheet.name.as_str())).into_iter().map(|name| name.to_lowercase()).collect::<BTreeSet<_>>();
    let ids = workbook.sheets.iter().map(|sheet| sheet.id.as_str()).collect::<BTreeSet<_>>();
    let mut suffix = index + 1;
    loop {
        let name = format!("__ov_chart_{suffix}");
        let id = format!("office-chart-helper-{suffix}");
        if !names.contains(&name.to_lowercase()) && !ids.contains(id.as_str()) { return (None, id, name); }
        suffix += 1;
    }
}

fn source_ref(source: &Context<'_>, row: u32, column: u32) -> String { format!("'{}'!${}${}", source.sheet_name.replace('\'', "''"), cell_ref(0, column).trim_end_matches('1'), row + 1) }

fn string_expression(reference: &str) -> String {
    format!("IFERROR(IF(ISBLANK({reference}),\"\",IF(ISLOGICAL({reference}),IF({reference},\"true\",\"false\"),{reference}&\"\")),\"\")")
}

fn group_expression(reference: &str, key_cell: &str, name: &str) -> String {
    let exact = format!("EXACT({},{})", string_expression(reference), key_cell);
    // A numeric cell and its identical JavaScript string key form one group.
    // Numeric comparison also avoids Office's locale/exponent text formatting.
    if let Some(value) = name.parse::<f64>().ok().filter(|value| value.is_finite() && ryu_js::Buffer::new().format_finite(*value) == name) {
        format!("IFERROR(OR(AND(ISNUMBER({reference}),{reference}={value}),{exact}),FALSE)")
    } else { format!("IFERROR({exact},FALSE)") }
}

fn numeric_expression(reference: &str, as_zero: bool) -> String {
    let missing = if as_zero { "0" } else { "NA()" };
    format!("IF(OR(ISBLANK({reference}),{reference}=\"\",ISLOGICAL({reference})),{missing},IFERROR(IF(ISNUMBER({reference}),{reference},VALUE({reference})),{missing}))")
}

fn add_text_budget(used: &mut usize, bytes: usize) -> Result<(), ExportError> {
    if bytes > 8192 { return Err(ExportError::invalid("native chart helper formula exceeds Excel's expression limit")); }
    *used = used.checked_add(bytes).filter(|used| *used <= MAX_HELPER_TEXT).ok_or_else(budget_error)?;
    Ok(())
}

fn budget_error() -> ExportError { ExportError::invalid("CHART_HELPER_RESOURCE_LIMIT: native chart helpers exceed their cell/text budget") }

fn add_value_text_budget(used: &mut usize, value: &str) -> Result<(), ExportError> {
    if value.encode_utf16().count() > 32_767 { return Err(ExportError::invalid("native chart helper text exceeds Excel's cell text limit")); }
    *used = used.checked_add(value.len()).filter(|used| *used <= MAX_HELPER_TEXT).ok_or_else(budget_error)?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn workbook(chart_type: &str, options: Value, rows: &[Vec<Value>]) -> WorkbookInput {
        let mut cells = Map::new();
        for (row, values) in rows.iter().enumerate() {
            for (column, value) in values.iter().enumerate() { cells.insert(format!("{row}:{column}"), json!({"value":value})); }
        }
        serde_json::from_value(json!({
            "kind":"workbook", "sheets":[{"id":"data","name":"Data","visible":true,"rowCount":rows.len(),"columnCount":rows[0].len(),"cells":cells}],
            "charts":[{"id":"chart","sheetId":"data","type":chart_type,"categoryField":0,"valueFields":[1],
                "source":{"orientation":"columns","range":{"startRow":0,"startCol":0,"endRow":rows.len()-1,"endCol":rows[0].len()-1}},
                "position":{"x":10,"y":20},"size":{"width":480,"height":300},"options":options}]
        })).unwrap()
    }

    #[test]
    fn grouped_zero_values_preserve_raw_errors_and_dynamic_default_size() {
        let mut book = workbook("bubble", json!({"invalidValueStrategy":"zero",
            "mapping":{"bubble":{"xIndex":0,"yIndex":1,"seriesIndex":2}},
            "errorBars":[{"seriesIndex":1,"direction":"y","type":"plus","valueType":"cust","plusIndex":1}]}),
            &[vec![json!("X"),json!("Y"),json!("Group")], vec![json!(1),Value::Null,json!("A")],
                vec![json!(2),json!(10),json!("B")], vec![json!(3),json!(20),json!("A")]]);
        prepare(&mut book, &mut Vec::new()).unwrap();
        let binding = book.charts.as_ref().unwrap()[0].native_bindings.as_ref().unwrap();
        assert_eq!(binding.groups.iter().map(|group| group.name.as_str()).collect::<Vec<_>>(), ["A","B"]);
        let first = &binding.groups[0];
        assert_eq!(first.vectors[&(1,true)].values, [Some(0.0),None,Some(20.0)]);
        assert_eq!(first.vectors[&(1,false)].values, [None,None,Some(20.0)]);
        assert_eq!(first.vectors[&(0,false)].values, [Some(1.0),None,Some(3.0)]);
        assert_eq!(first.default_size.as_ref().unwrap().values, [Some(1.0),None,Some(1.0)]);
        let native = super::super::chart_xml(&book, &book.charts.as_ref().unwrap()[0], 0, None).unwrap();
        let root = office_io_common::chart_xml::parse(&native).unwrap();
        let series = root.descendants("ser");
        assert_eq!(series.len(), 2);
        assert!(series[0].child("bubbleSize").unwrap().child("numRef").is_some());
        assert_eq!(root.child("chart").unwrap().value("plotVisOnly"), Some("0"));
        assert!(book.sheets[1].cells.values().filter_map(|cell| cell.get("formula").and_then(Value::as_str)).any(|formula| formula.contains("EXACT(")));
    }

    #[test]
    fn helpers_apply_all_current_row_visibility_reasons_before_joining_categories() {
        let rows = (0..6).map(|row| vec![json!(format!("R{row}")), json!("Region"), json!(row * 10)]).collect::<Vec<_>>();
        let mut book = workbook("line", json!({"mapping":{"categoryIndexes":[0,1]}}), &rows);
        book.charts.as_mut().unwrap()[0].value_fields = vec![2];
        book.sheets[0].hidden_rows = vec![1];
        book.sheets[0].row_heights.insert("4".into(), 0.0);
        book.sheets[0].auto_filter = Some(serde_json::from_value(json!({"range":{"startRow":0,"startCol":0,"endRow":5,"endCol":2},"filteredRows":[2]})).unwrap());
        book.sheets[0].outline = Some(serde_json::from_value(json!({"groups":[{"id":"collapsed","axis":"row","start":3,"end":3,"collapsed":true}]})).unwrap());
        prepare(&mut book, &mut Vec::new()).unwrap();
        let binding = book.charts.as_ref().unwrap()[0].native_bindings.as_ref().unwrap();
        assert_eq!(binding.point_count, 1);
        assert_eq!(binding.category.as_ref().unwrap().values, ["R5 / Region"]);
        assert_eq!(binding.vectors[&(2,false)].values, [Some(50.0)]);
        assert!(book.sheets[1].cells["1:1"]["formula"].as_str().unwrap().contains("'Data'!$A$6"));
    }

    #[test]
    fn a_hidden_primary_bubble_role_writes_an_empty_native_plot_and_can_be_refreshed() {
        let mut book = workbook("bubble", json!({"mapping":{"bubble":{"xIndex":0,"yIndex":1,"labelIndex":2}}}),
            &[vec![json!("X"),json!("Y"),json!("Label")],vec![json!(1),json!(2),json!("A")]]);
        book.sheets[0].hidden_columns = vec![2];
        prepare(&mut book, &mut Vec::new()).unwrap();
        let native = super::super::chart_xml(&book, &book.charts.as_ref().unwrap()[0], 0, None).unwrap();
        let root = office_io_common::chart_xml::parse(&native).unwrap();
        assert!(root.descendants("ser").is_empty());
        assert!(root.child("chart").unwrap().child("legend").is_none());
        book.sheets[0].hidden_columns.clear();
        prepare(&mut book, &mut Vec::new()).unwrap();
        assert_eq!(book.sheets.len(), 2);
        assert!(!book.charts.as_ref().unwrap()[0].native_bindings.as_ref().unwrap().empty);
    }

    #[test]
    fn returning_to_direct_binding_discards_the_previous_helper() {
        let mut book = workbook("line", json!({"invalidValueStrategy":"zero"}),
            &[vec![json!("Category"),json!("Value")],vec![json!("A"),Value::Null],vec![json!("B"),json!(2)]]);
        prepare(&mut book, &mut Vec::new()).unwrap();
        assert!(book.charts.as_ref().unwrap()[0].native_bindings.is_some());
        book.charts.as_mut().unwrap()[0].options = serde_json::from_value(json!({})).unwrap();
        book.sheets[0].cells.get_mut("1:1").unwrap()["value"] = json!(3);
        prepare(&mut book, &mut Vec::new()).unwrap();
        assert!(book.charts.as_ref().unwrap()[0].native_bindings.is_none());
        let native = super::super::chart_xml(&book, &book.charts.as_ref().unwrap()[0], 0, None).unwrap();
        assert!(native.contains("$A$2:$A$3"));
        assert!(!native.contains("<c:f>'__ov_chart_"));
    }

    #[test]
    fn imported_helper_keys_are_reused_only_with_matching_cell_and_owner_proof() {
        let mut book = workbook("line", json!({"invalidValueStrategy":"zero"}),
            &[vec![json!("Category"),json!("Value")],vec![json!("A"),json!(2)]]);
        prepare(&mut book, &mut Vec::new()).unwrap();
        book.sheets[1].cells.get_mut("0:0").unwrap()["value"] = json!(format!("{MARKER}original-id"));
        book.sheets[1].metadata = Some(json!({"officeChartHelper":{"version":1,"chartId":"chart","key":"original-id"}}));
        prepare(&mut book, &mut Vec::new()).unwrap();
        assert_eq!(book.sheets.len(), 2);
        book.sheets[1].cells.get_mut("0:0").unwrap()["value"] = json!("user data");
        prepare(&mut book, &mut Vec::new()).unwrap();
        assert_eq!(book.sheets.len(), 3);
        assert_eq!(book.sheets[1].cells["0:0"]["value"], "user data");
    }

    #[test]
    fn sparse_sources_are_rejected_before_expanding_helpers_beyond_the_cell_budget() {
        let mut book = workbook("line", json!({"invalidValueStrategy":"zero"}),
            &[vec![json!("Category"),json!("Value")],vec![json!("A"),json!(2)]]);
        book.sheets[0].row_count = Some(1_000_000);
        book.charts.as_mut().unwrap()[0].source.range.end_row = Some(999_999);
        assert!(prepare(&mut book, &mut Vec::new()).unwrap_err().to_string().contains("CHART_HELPER_RESOURCE_LIMIT"));
        assert_eq!(book.sheets.len(), 1);
    }

    #[test]
    fn text_keys_match_javascript_scalars_and_group_formulas_use_case_sensitive_headers() {
        assert_eq!(text(Some(&json!(1.0))), "1");
        assert_eq!(text(Some(&json!(1e21))), "1e+21");
        assert_eq!(text(Some(&json!(false))), "false");
        let expression = group_expression("'Data'!$A$2", "$B$1", "North");
        assert!(expression.contains("EXACT("));
        assert!(expression.contains("$B$1"));
        assert!(!expression.contains("\"North\""));
        assert!(group_expression("'Data'!$A$2", "$B$1", "1").contains("ISNUMBER('Data'!$A$2)"));
    }

    #[test]
    fn native_date_categories_use_numeric_formula_caches_and_keep_missing_points() {
        let mut book = workbook("line", json!({"mapping":{"useDateAxis":true}}),
            &[vec![json!("Date"),json!("Value")],vec![json!("1900-02-28"),json!(2)],
                vec![json!("1900-02-29"),json!(4)],vec![Value::Null,json!(8)],vec![json!(61),json!(16)]]);
        book.charts.as_mut().unwrap()[0].x_axis = Some(serde_json::from_value(json!({"scaleType":"linear","title":"Date",
            "titleStyle":{"family":"Arial","fontSize":16,"color":"#ff0000"},"titleAlignment":"end"})).unwrap());
        prepare(&mut book, &mut Vec::new()).unwrap();
        let chart = &book.charts.as_ref().unwrap()[0];
        let category = chart.native_bindings.as_ref().unwrap().category.as_ref().unwrap();
        assert_eq!(category.dates.as_ref().unwrap(), &[Some(59.0),Some(60.0),None,Some(61.0)]);
        let native = office_io_common::chart_xml::parse(&super::super::chart_xml(&book, chart, 0, None).unwrap()).unwrap();
        let axis = native.descendants("dateAx")[0];
        assert_eq!(axis.value("baseTimeUnit"), Some("days"));
        assert_eq!(axis.child("numFmt").unwrap().attrs["formatCode"], "yyyy-mm-dd");
        let title = axis.child("title").unwrap();
        assert_eq!(title.descendants("rPr")[0].attrs["sz"], "1200");
        assert_eq!(title.descendants("latin")[0].attrs["typeface"], "Arial");
        let categories = native.descendants("ser")[0].child("cat").unwrap();
        assert!(categories.child("strRef").is_none());
        let reference = categories.child("numRef").unwrap();
        assert!(reference.child("f").unwrap().text.contains("__ov_chart_"));
        let cache = reference.child("numCache").unwrap();
        assert_eq!(cache.value("ptCount"), Some("4"));
        assert_eq!(cache.children_named("pt").map(|point| point.attrs["idx"].as_str()).collect::<Vec<_>>(), ["0","1","3"]);
        let source: Value = serde_json::from_str(&native.descendants("source")[0].text).unwrap();
        assert_eq!(source["authored"]["xAxis"]["titleAlignment"], "end");
        assert_eq!(source["authored"]["options"]["mapping"]["useDateAxis"], true);
    }

    #[test]
    fn grouped_bubble_date_x_has_its_own_epoch_and_mask_separate_from_zero_y() {
        let mut book = workbook("bubble", json!({"invalidValueStrategy":"zero","mapping":{"useDateAxis":true,
            "bubble":{"xIndex":0,"yIndex":1,"seriesIndex":2}}}),
            &[vec![json!("Date"),json!("Value"),json!("Group")],vec![json!("1904-01-01"),Value::Null,json!("A")],
                vec![json!("1904-01-02T08:00+08:00"),json!(10),json!("B")],vec![Value::Null,json!(20),json!("A")]]);
        book.date1904 = true;
        prepare(&mut book, &mut Vec::new()).unwrap();
        let binding = book.charts.as_ref().unwrap()[0].native_bindings.as_ref().unwrap();
        assert_eq!(binding.groups[0].vectors[&(1,true)].values, [Some(0.0),None,Some(20.0)]);
        assert_eq!(binding.groups[0].date_vectors[&0].values, [Some(0.0),None,None]);
        assert_eq!(binding.groups[1].date_vectors[&0].values, [None,Some(1.0),None]);
        let native = office_io_common::chart_xml::parse(&super::super::chart_xml(&book, &book.charts.as_ref().unwrap()[0], 0, None).unwrap()).unwrap();
        assert!(native.descendants("dateAx").is_empty());
        assert_eq!(native.value("date1904"), Some("1"));
        assert!(native.descendants("valAx").iter().any(|axis| axis.child("numFmt").is_some_and(|format| format.attrs["formatCode"] == "yyyy-mm-dd")));
        assert!(native.descendants("ser")[0].child("xVal").unwrap().child("numRef").is_some());
    }

    #[test]
    fn dates_reject_invalid_calendar_values_bounds_and_missing_dates_with_plotted_y() {
        for value in [json!("2026-02-30"),json!(true),Value::Null] {
            let mut book = workbook("line", json!({"mapping":{"useDateAxis":true}}),
                &[vec![json!("Date"),json!("Value")],vec![value,json!(2)]]);
            assert!(prepare(&mut book, &mut Vec::new()).unwrap_err().to_string().contains("CHART_DATE_AXIS_"));
            assert_eq!(book.sheets.len(), 1);
        }
        let mut book = workbook("line", json!({"mapping":{"useDateAxis":true}}),
            &[vec![json!("Date"),json!("Value")],vec![json!("1904-01-01"),json!(2)]]);
        book.date1904 = true;
        book.charts.as_mut().unwrap()[0].x_axis = Some(serde_json::from_value(json!({"scaleType":"linear","min":-1})).unwrap());
        assert!(prepare(&mut book, &mut Vec::new()).unwrap_err().to_string().contains("CHART_DATE_AXIS_LIMIT"));
    }

    #[test]
    fn missing_dates_without_y_are_not_made_plottable_by_size_or_error_columns() {
        for grouped in [false, true] {
            let mapping = if grouped {
                json!({"useDateAxis":true,"bubble":{"xIndex":0,"yIndex":1,"sizeIndex":2,"seriesIndex":3}})
            } else {
                json!({"useDateAxis":true,"bubble":{"xIndex":0,"yIndex":1,"sizeIndex":2}})
            };
            let mut book = workbook("bubble", json!({"mapping":mapping,
                "errorBars":[{"seriesIndex":1,"direction":"y","type":"plus","valueType":"cust","plusIndex":2}]}),
                &[vec![json!("Date"),json!("Y"),json!("Size"),json!("Group")],
                    vec![Value::Null,Value::Null,json!(10),json!("A")]]);
            prepare(&mut book, &mut Vec::new()).unwrap();
            let bindings = book.charts.as_ref().unwrap()[0].native_bindings.as_ref().unwrap();
            if grouped {
                assert_eq!(bindings.groups[0].vectors[&(1,false)].values, [None]);
                assert_eq!(bindings.groups[0].vectors[&(2,false)].values, [Some(10.0)]);
            } else {
                assert_eq!(bindings.vectors[&(1,false)].values, [None]);
                assert_eq!(bindings.vectors[&(2,false)].values, [Some(10.0)]);
            }
            // The zero strategy really creates a Y value, so this requires a
            // source date even though the unconverted Y cell is empty.
            book.charts.as_mut().unwrap()[0].options.invalid_value_strategy = Some("zero".into());
            assert!(prepare(&mut book, &mut Vec::new()).unwrap_err().to_string().contains("CHART_DATE_AXIS_EMPTY"));
        }
    }
}
