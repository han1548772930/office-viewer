//! Source identities and authored overrides are resolved before native plot and
//! axis validation. A display alias never participates in selector matching.

use std::collections::{BTreeMap, BTreeSet};
use office_io_common::chart_options::{BubbleMapping, CandlestickMapping, Series};
use super::{ChartInput, Context, ExportError, MAX_SERIES, NativeChartKind, SeriesSpec, native_chart_kind, style};

fn fields(context: &Context<'_>, chart: &ChartInput) -> Result<Vec<usize>, ExportError> {
    let mut out = chart.value_fields.clone();
    if out.is_empty() {
        for field in 0..context.field_count {
            if field == chart.category_field || chart.chart_type == "scatter" && chart.series_to_category.values().any(|xs| xs.contains(&field)) { continue; }
            let (_, values, _) = context.vector(chart, field)?;
            if values.iter().any(Option::is_some) { out.push(field); }
        }
    }
    if out.len() > MAX_SERIES { return Err(ExportError::invalid("chart series count exceeds the native export budget")); }
    let mut seen = BTreeSet::new();
    for field in &out {
        if *field >= context.field_count || !seen.insert(*field) { return Err(ExportError::invalid(format!("chart {} has a duplicate or out-of-range value field", chart.id))); }
    }
    if out.is_empty() { return Err(ExportError::invalid(format!("chart {} has no value series", chart.id))); }
    Ok(out)
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::{Map, Value, json};
    use super::super::{WorkbookInput, chart_xml, context, prepare_native_bindings};
    use office_io_common::chart_xml as native_xml;

    fn workbook(kind: &str, options: Value, rows: &[Vec<Value>], values: &[usize]) -> WorkbookInput {
        let mut cells = Map::new();
        for (row, entries) in rows.iter().enumerate() {
            for (column, value) in entries.iter().enumerate() { cells.insert(format!("{row}:{column}"), json!({"value":value})); }
        }
        serde_json::from_value(json!({"kind":"workbook", "sheets":[{"id":"source","name":"Source","visible":true,
            "rowCount":rows.len(),"columnCount":rows[0].len(),"cells":cells}],
            "charts":[{"id":"chart","sheetId":"source","type":kind,"categoryField":0,"valueFields":values,
                "source":{"orientation":"columns","range":{"startRow":0,"startCol":0,"endRow":rows.len()-1,"endCol":rows[0].len()-1}},
                "position":{"x":0,"y":0},"size":{"width":480,"height":300},"options":options}]})).unwrap()
    }

    fn regular(options: Value) -> WorkbookInput {
        workbook("combination", options, &[vec![json!("Category"),json!("Revenue"),json!("Cost")],
            vec![json!("A"),json!(10),json!(3)],vec![json!("B"),json!(20),json!(7)]], &[1,2])
    }

    #[test]
    fn original_names_overlay_fields_before_native_combo_axis_and_plot_selection() {
        let book = regular(json!({
            "allSeriesStyle":{"name":"Legend alias","color":"#111111","border":{"width":2},"label":{"visible":true,"fontSize":16}},
            "series":{"2":{"color":"#222222","border":{"dashType":"dotted"},"label":{"color":"#334455"}}},
            "namedSeries":[{"name":"Cost","style":{"name":"Expense","type":"line","rightAxis":true,"color":"#ff0000"}},
                {"name":"Legend alias","style":{"color":"#00ff00"}}],
            "rightYAxis":{"scaleType":"linear","title":"Expenses"}
        }));
        let chart = &book.charts.as_ref().unwrap()[0];
        let specs = specifications(&context(&book, chart).unwrap(), chart).unwrap();
        assert_eq!(specs[1].source_name.as_deref(), Some("Cost"));
        assert_eq!(specs[1].kind.plot, "lineChart");
        assert!(specs[1].secondary);
        assert_eq!(specs[0].style.color.as_deref(), Some("#111111"));
        assert_eq!(specs[1].style.color.as_deref(), Some("#ff0000"));
        assert_eq!(specs[1].style.border.as_ref().unwrap().width, Some(2.0));
        assert_eq!(specs[1].style.border.as_ref().unwrap().dash_type.as_deref(), Some("dotted"));
        assert_eq!(specs[1].style.label.as_ref().unwrap().font.font_size, Some(16.0));
        let root = native_xml::parse(&chart_xml(&book, chart, 0, None).unwrap()).unwrap();
        let line = root.descendants("lineChart")[0];
        assert_eq!(line.children_named("ser").next().unwrap().child("tx").unwrap().child("v").unwrap().text, "Expense");
        assert_eq!(line.children_named("axId").nth(1).unwrap().attrs["val"], "1000005");
        assert_eq!(root.descendants("barChart")[0].children_named("ser").count(), 1);
    }

    #[test]
    fn named_stacked_override_is_validated_against_its_effective_axis() {
        let mut book = regular(json!({"namedSeries":[{"name":"Cost","style":{"type":"columnStacked","rightAxis":true}}],
            "rightYAxis":{"scaleType":"log"}}));
        let chart = &book.charts.as_ref().unwrap()[0];
        assert!(specifications(&context(&book, chart).unwrap(), chart).unwrap_err().to_string().contains("linear value axis"));
        book.charts.as_mut().unwrap()[0].options = serde_json::from_value(json!({"namedSeries":[{"name":"Cost","style":{"type":"line","rightAxis":true}}],
            "rightYAxis":{"scaleType":"log"}})).unwrap();
        let chart = &book.charts.as_ref().unwrap()[0];
        assert_eq!(specifications(&context(&book, chart).unwrap(), chart).unwrap()[1].kind.plot, "lineChart");
    }

    #[test]
    fn grouped_bubbles_match_case_space_and_empty_names_without_using_a_header_template() {
        let mut book = workbook("bubble", json!({
            "mapping":{"bubble":{"xIndex":0,"yIndex":1,"seriesIndex":2}},
            "allSeriesStyle":{"border":{"width":2}},"series":{"1":{"color":"#22aa22","fillOpacity":0.4}},
            "namedSeries":[{"name":"North","style":{"color":"#ff0000","rightAxis":true}},
                {"name":" north ","style":{"color":"#abcdef"}}, {"name":"","style":{"color":"#112233"}},
                {"name":"Y","style":{"color":"#ffff00"}}, {"name":"Inactive","style":{"name":"Future"}}]
        }), &[vec![json!("X"),json!("Y"),json!("Group")],vec![json!(1),json!(2),json!("North")],
            vec![json!(2),json!(3),json!("north")],vec![json!(3),json!(4),json!(" north ")],vec![json!(4),json!(5),Value::Null]], &[1]);
        prepare_native_bindings(&mut book, &mut Vec::new()).unwrap();
        let chart = &book.charts.as_ref().unwrap()[0];
        let specs = specifications(&context(&book, chart).unwrap(), chart).unwrap();
        assert_eq!(specs.iter().map(|spec| spec.source_name.as_deref().unwrap()).collect::<Vec<_>>(), ["North","north"," north ",""]);
        assert_eq!(specs.iter().map(|spec| spec.style.color.as_deref().unwrap()).collect::<Vec<_>>(), ["#ff0000","#22aa22","#abcdef","#112233"]);
        assert_eq!(specs.iter().map(|spec| spec.secondary).collect::<Vec<_>>(), [true,false,false,false]);
        assert!(specs.iter().all(|spec| spec.style.fill_opacity == Some(0.4) && spec.style.border.as_ref().unwrap().width == Some(2.0)));
        let root = native_xml::parse(&chart_xml(&book, chart, 0, None).unwrap()).unwrap();
        assert_eq!(root.descendants("bubbleChart").len(), 2);
        assert_eq!(root.descendants("ser").len(), 4);
        let source = &root.descendants("source")[0].text;
        let manifest: Value = serde_json::from_str(source).unwrap();
        assert_eq!(manifest["authored"]["options"]["namedSeries"].as_array().unwrap().len(), 5);
    }

    #[test]
    fn selectors_reject_duplicate_configuration_and_ambiguous_active_names_only() {
        let duplicate = regular(json!({"namedSeries":[{"name":"Future","style":{}},{"name":"Future","style":{}}]}));
        let chart = &duplicate.charts.as_ref().unwrap()[0];
        assert!(planning_specifications(&context(&duplicate, chart).unwrap(), chart).unwrap_err().to_string().contains("CHART_NAMED_SERIES_DUPLICATE"));
        let mut ambiguous = regular(json!({"namedSeries":[{"name":"Revenue","style":{"color":"#ff0000"}}]}));
        ambiguous.sheets[0].cells.get_mut("0:2").unwrap()["value"] = json!("Revenue");
        let chart = &ambiguous.charts.as_ref().unwrap()[0];
        assert!(specifications(&context(&ambiguous, chart).unwrap(), chart).unwrap_err().to_string().contains("CHART_NAMED_SERIES_AMBIGUOUS"));
        ambiguous.sheets[0].hidden_columns = vec![2];
        prepare_native_bindings(&mut ambiguous, &mut Vec::new()).unwrap();
        let chart = &ambiguous.charts.as_ref().unwrap()[0];
        assert_eq!(specifications(&context(&ambiguous, chart).unwrap(), chart).unwrap().len(), 1);
        ambiguous.sheets[0].hidden_columns.clear();
        prepare_native_bindings(&mut ambiguous, &mut Vec::new()).unwrap();
        let chart = &ambiguous.charts.as_ref().unwrap()[0];
        assert!(chart_xml(&ambiguous, chart, 0, None).unwrap_err().to_string().contains("CHART_NAMED_SERIES_AMBIGUOUS"));
        assert!(serde_json::from_value::<office_io_common::chart_options::Options>(json!({"namedSeries":[{"name":"Missing style"}]})).is_err());
    }

    #[test]
    fn blank_bubble_headers_keep_the_raw_empty_selector_while_regular_headers_use_source_fallback() {
        let book = workbook("bubble", json!({"mapping":{"bubble":{"xIndex":0,"yIndex":1}},
            "namedSeries":[{"name":"","style":{"color":"#123456"}},{"name":"Series 2","style":{"color":"#ffffff"}}]}),
            &[vec![json!("X"),Value::Null],vec![json!(1),json!(2)]], &[1]);
        let chart = &book.charts.as_ref().unwrap()[0];
        assert_eq!(specifications(&context(&book, chart).unwrap(), chart).unwrap()[0].style.color.as_deref(), Some("#123456"));
        let mut book = regular(json!({"namedSeries":[{"name":"Series 2","style":{"color":"#123456"}}]}));
        book.sheets[0].cells.get_mut("0:1").unwrap()["value"] = Value::Null;
        let chart = &book.charts.as_ref().unwrap()[0];
        assert_eq!(specifications(&context(&book, chart).unwrap(), chart).unwrap()[0].style.color.as_deref(), Some("#123456"));
    }

    #[test]
    fn hidden_effective_plot_settings_are_validated_when_the_source_becomes_active() {
        let mut book = regular(json!({"namedSeries":[{"name":"Cost","style":{"type":"columnStacked","rightAxis":true}}],
            "rightYAxis":{"scaleType":"log"}}));
        book.sheets[0].hidden_columns = vec![2];
        prepare_native_bindings(&mut book, &mut Vec::new()).unwrap();
        let chart = &book.charts.as_ref().unwrap()[0];
        assert_eq!(specifications(&context(&book, chart).unwrap(), chart).unwrap().len(), 1);
        assert!(chart_xml(&book, chart, 0, None).is_ok());
        book.sheets[0].hidden_columns.clear();
        prepare_native_bindings(&mut book, &mut Vec::new()).unwrap();
        let chart = &book.charts.as_ref().unwrap()[0];
        assert!(chart_xml(&book, chart, 0, None).unwrap_err().to_string().contains("linear value axis"));
    }

    #[test]
    fn stock_native_vectors_share_the_close_field_logical_series_style() {
        let book = workbook("candlestick", json!({"series":{"1":{"color":"#ff0000"},"3":{"color":"#223344","border":{"width":2}}},
            "namedSeries":[{"name":"最高 / 最低 / 收盘","style":{"fillOpacity":0.5}}]}),
            &[vec![json!("Date"),json!("High"),json!("Low"),json!("Close")],vec![json!(1),json!(10),json!(4),json!(7)]], &[1,2,3]);
        let chart = &book.charts.as_ref().unwrap()[0];
        let specs = specifications(&context(&book, chart).unwrap(), chart).unwrap();
        assert_eq!(specs.iter().map(|spec| spec.field).collect::<Vec<_>>(), [1,2,3]);
        assert!(specs.iter().all(|spec| spec.style.color.as_deref() == Some("#223344") && spec.style.fill_opacity == Some(0.5)
            && spec.style.border.as_ref().unwrap().width == Some(2.0)));
    }
}

fn named_styles(chart: &ChartInput) -> Result<BTreeMap<&str, &Series>, ExportError> {
    let mut names = BTreeMap::new();
    for entry in &chart.options.named_series {
        if names.insert(entry.name.as_str(), &entry.style).is_some() {
            return Err(ExportError::invalid(format!("CHART_NAMED_SERIES_DUPLICATE: chart {} repeats a source name selector", chart.id)));
        }
    }
    Ok(names)
}

fn specification(context: &Context<'_>, chart: &ChartInput, field: usize, style_field: usize, native: NativeChartKind,
    bubble: Option<&BubbleMapping>, group_index: Option<usize>, source_name: Option<String>, named: &BTreeMap<&str, &Series>, validate: bool) -> Result<SeriesSpec, ExportError> {
    let inherited: Series = style::overlay(chart.options.all_series_style.as_ref(), chart.options.series.get(&style_field))?;
    let specific = source_name.as_deref().and_then(|name| named.get(name).copied());
    let style: Series = style::overlay(Some(&inherited), specific)?;
    let kind = if chart.chart_type == "combination" {
        match native_chart_kind(style.kind.as_deref().unwrap_or("column"))
            .filter(|kind| matches!(kind.plot, "barChart" | "lineChart" | "areaChart")) {
            Some(kind) => kind,
            None if !validate => native,
            None => return Err(ExportError::invalid("combination series requires column, line or area type")),
        }
    } else { native };
    let secondary = style.right_axis.unwrap_or(false);
    if validate && secondary && (kind.pie || kind.three_d || kind.plot == "radarChart" || kind.plot == "stockChart") {
        return Err(ExportError::invalid("this chart type does not support a secondary value axis"));
    }
    if validate && (matches!(kind.grouping, Some("stacked" | "percentStacked")) || kind.plot == "radarChart")
        && (if secondary { chart.options.right_y_axis.as_ref() } else { chart.y_axis.as_ref() }).is_some_and(|axis| axis.scale_type != "linear") {
        return Err(ExportError::invalid(if kind.plot == "radarChart" { "radar charts require a linear value axis" } else { "stacked charts require a linear value axis" }));
    }
    let x_field = if let Some(bubble) = bubble { Some(bubble.x_index) } else if kind.plot == "scatterChart" {
        let fields = chart.series_to_category.get(&field);
        if fields.is_some_and(|fields| fields.len() != 1) { return Err(ExportError::invalid("scatter series must have one X field")); }
        Some(fields.and_then(|fields| fields.first().copied()).unwrap_or(chart.category_field))
    } else { None };
    if let Some(field) = x_field { context.coordinates(chart, field, None)?; }
    Ok(SeriesSpec { field, kind, secondary, x_field, size_field: bubble.and_then(|mapping| mapping.size_index), style, group_index, source_name })
}

pub(super) fn specifications(context: &Context<'_>, chart: &ChartInput) -> Result<Vec<SeriesSpec>, ExportError> {
    resolve_specifications(context, chart, true)
}

/// Helpers still need all numeric roles before they construct visible groups.
/// Ambiguous active names are checked against the final, prepared source.
pub(super) fn planning_specifications(context: &Context<'_>, chart: &ChartInput) -> Result<Vec<SeriesSpec>, ExportError> {
    resolve_specifications(context, chart, false)
}

fn resolve_specifications(context: &Context<'_>, chart: &ChartInput, validate_active: bool) -> Result<Vec<SeriesSpec>, ExportError> {
    let named = named_styles(chart)?;
    let selected = fields(context, chart)?;
    let native = native_chart_kind(if chart.chart_type == "combination" { "column" } else { &chart.chart_type })
        .ok_or_else(|| ExportError::unsupported(format!("XLSX chart type {} requires its native chart adapter", chart.chart_type)))?;
    let mapping = chart.options.mapping.as_ref();
    let bubble = (chart.chart_type == "bubble").then(|| mapping.and_then(|mapping| mapping.bubble.clone()).unwrap_or(BubbleMapping {
        x_index: chart.category_field, y_index: selected[0], size_index: selected.get(1).copied(), label_index: None, series_index: None,
    }));
    let stock = (chart.chart_type == "candlestick").then(|| mapping.and_then(|mapping| mapping.candlestick.clone()).unwrap_or_else(|| {
        if selected.len() >= 4 { CandlestickMapping { open_index: Some(selected[0]), high_index: selected[1], low_index: selected[2], close_index: selected[3] } }
        else { CandlestickMapping { open_index: None, high_index: selected[0], low_index: selected.get(1).copied().unwrap_or(usize::MAX), close_index: selected.get(2).copied().unwrap_or(usize::MAX) } }
    }));
    let selected = if let Some(bubble) = &bubble { vec![bubble.y_index] }
        else if let Some(stock) = &stock { stock.open_index.into_iter().chain([stock.high_index, stock.low_index, stock.close_index]).collect() }
        else if native.pie && chart.value_fields.is_empty() { selected.into_iter().take(1).collect() } else { selected };
    if native.pie && selected.len() != 1 { return Err(ExportError::invalid("pie charts require exactly one value series")); }
    let grouped = bubble.as_ref().is_some_and(|mapping| mapping.series_index.is_some());
    if let Some(bubble) = &bubble {
        for field in [Some(bubble.x_index), Some(bubble.y_index), bubble.size_index, bubble.label_index, bubble.series_index].into_iter().flatten() { context.coordinates(chart, field, None)?; }
        if grouped && chart.native_bindings.as_ref().is_some_and(|bindings| !bindings.empty && bindings.groups.is_empty()) { return Err(ExportError::invalid("grouped bubble chart formula bindings were not prepared")); }
    }
    let mut out = Vec::new();
    let mut selected_fields = BTreeSet::new();
    for field in selected {
        context.coordinates(chart, field, None)?;
        if !selected_fields.insert(field) { return Err(ExportError::invalid("chart source roles must use distinct value fields")); }
        let visible = |field: usize| chart.native_bindings.as_ref().and_then(|bindings| bindings.visible_fields.as_ref())
            .is_none_or(|visible| visible.contains(&field));
        let x_field = (chart.chart_type == "scatter").then(|| chart.series_to_category.get(&field)
            .and_then(|fields| fields.first()).copied().unwrap_or(chart.category_field));
        if validate_active && (chart.native_bindings.as_ref().is_some_and(|bindings| bindings.empty)
            || !visible(field) || x_field.is_some_and(|field| !visible(field))) { continue; }
        // HLC and OHLC are a single logical series keyed by closeIndex in the
        // workbench. Native vectors keep their own formula field identities.
        let style_field = stock.as_ref().map_or(field, |stock| stock.close_index);
        if grouped {
            if let Some(bindings) = &chart.native_bindings {
                for (group_index, group) in bindings.groups.iter().enumerate() {
                    out.push(specification(context, chart, field, style_field, native, bubble.as_ref(), Some(group_index), Some(group.name.clone()), &named, validate_active)?);
                }
            } else {
                // Helpers first plan the shared numeric roles. Names bind only
                // to actual groups after current visibility has been applied.
                out.push(specification(context, chart, field, style_field, native, bubble.as_ref(), None, None, &named, validate_active)?);
            }
        } else {
            let (_, header) = context.header(chart, field)?;
            let name = if stock.is_some() { "最高 / 最低 / 收盘".to_owned() }
                else if chart.chart_type != "bubble" && header.is_empty() { format!("Series {}", field + 1) } else { header };
            out.push(specification(context, chart, field, style_field, native, bubble.as_ref(), None, Some(name), &named, validate_active)?);
        }
    }
    for field in chart.series_to_category.keys() {
        if chart.chart_type != "scatter" || !selected_fields.contains(field) { return Err(ExportError::invalid("scatter X mapping references an unselected series")); }
    }
    if let Some(bindings) = &chart.native_bindings {
        if bindings.empty { return Ok(Vec::new()); }
        if let Some(visible) = &bindings.visible_fields { out.retain(|spec| visible.contains(&spec.field) && spec.x_field.is_none_or(|field| visible.contains(&field))); }
    }
    // Three or four stock vectors form one logical candlestick series. Other
    // active source series must be uniquely identifiable by a named selector.
    if validate_active && stock.is_none() {
        let mut counts = BTreeMap::<&str, usize>::new();
        for spec in &out {
            if let Some(name) = spec.source_name.as_deref().filter(|name| named.contains_key(name)) { *counts.entry(name).or_default() += 1; }
        }
        if counts.values().any(|count| *count > 1) { return Err(ExportError::invalid("CHART_NAMED_SERIES_AMBIGUOUS: a source name selector matches multiple active chart series")); }
    }
    Ok(out)
}
