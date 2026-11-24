//! Draft source ownership, cache fidelity and byte-budget contracts.

use super::*;
use serde_json::json;

fn workbook() -> WorkbookInput {
    serde_json::from_value(json!({"kind":"workbook","sheets":[{"id":"data","name":"Data","rowCount":3,"columnCount":2,
        "cells":{"0:0":{"value":"Category"},"0:1":{"value":"Value"},
            "1:0":{"value":"A"},"1:1":{"value":10},"2:0":{"value":"B"},"2:1":{"value":null}}}],
        "charts":[{"id":"chart","sheetId":"data","type":"bar","categoryField":0,"valueFields":[1],
            "source":{"orientation":"columns","range":{"startRow":0,"startCol":0,"endRow":2,"endCol":1}},
            "position":{"x":0,"y":0},"size":{"width":480,"height":300}}]})).unwrap()
}

fn spec() -> SeriesSpec {
    SeriesSpec { field:1, kind:native_chart_kind("bar").unwrap(), secondary:false, x_field:None, size_field:None,
        style:Series::default(), group_index:None, source_name:None }
}

#[test]
fn prepared_numeric_and_date_vectors_are_borrowed_from_the_same_source_owner() {
    let mut book = workbook();
    book.charts[0].native_bindings = Some(NativeBindings {
        vectors: BTreeMap::from([((1,false),NativeVector { formula:"'Data'!$B$2:$B$3".into(),values:vec![Some(10.0),None] })]),
        date_vectors: BTreeMap::from([(1,NativeVector { formula:"'Dates'!$B$2:$B$3".into(),values:vec![Some(1.0),Some(2.0)] })]),
        ..Default::default()
    });
    let chart = &book.charts[0];
    let context = context(&book,chart).unwrap();
    let (formula, values) = bound_vector(&context,chart,&spec(),1,false).unwrap();
    let owner = &chart.native_bindings.as_ref().unwrap().vectors[&(1,false)];
    assert!(matches!(&formula,Cow::Borrowed(_)));
    assert!(matches!(&values,Cow::Borrowed(_)));
    assert_eq!(formula.as_ptr(),owner.formula.as_ptr());
    assert_eq!(values.as_ptr(),owner.values.as_ptr());
    let (_, dates) = bound_date_vector(chart,&spec(),1).unwrap();
    assert!(matches!(&dates,Cow::Borrowed(_)));
    assert_eq!(dates.as_ptr(),chart.native_bindings.as_ref().unwrap().date_vectors[&1].values.as_ptr());
}

#[test]
fn direct_numeric_vectors_own_only_numeric_data_and_zero_policy_keeps_source_cells_unchanged() {
    let book = workbook();
    let chart = &book.charts[0];
    let context = context(&book,chart).unwrap();
    let (_, raw) = bound_vector(&context,chart,&spec(),1,false).unwrap();
    let (_, zero) = bound_vector(&context,chart,&spec(),1,true).unwrap();
    assert!(matches!(&raw,Cow::Owned(_)));
    assert_eq!(raw.as_ref(),[Some(10.0),None]);
    assert_eq!(zero.as_ref(),[Some(10.0),Some(0.0)]);
    assert!(book.sheets[0].cells["2:1"]["value"].is_null());
    assert_eq!(context.category_vector(chart,0).unwrap().1,["A","B"]);
}

#[test]
fn directly_appended_caches_preserve_escaped_text_formulas_and_missing_numeric_positions() {
    let mut out = format!("<c:ser xmlns:c=\"{}\">",chart_xml::CHART_NS);
    append_number_reference(&mut out,"val","'A & B'!$B$2:$B$4",&[Some(2.0),None,Some(-3.0)],"0.00 < 1").unwrap();
    append_string_cache(&mut out,&["A<&\"'>中文".into()]).unwrap();
    append_chart_xml(&mut out,"</c:ser>").unwrap();
    let root = chart_xml::parse(&out).unwrap();
    let reference = root.child("val").unwrap().child("numRef").unwrap();
    assert_eq!(reference.child("f").unwrap().text,"'A & B'!$B$2:$B$4");
    let cache = reference.child("numCache").unwrap();
    assert_eq!(cache.child("formatCode").unwrap().text,"0.00 < 1");
    assert_eq!(cache.value("ptCount"),Some("3"));
    let points = cache.children_named("pt").collect::<Vec<_>>();
    assert_eq!(points.len(),2);
    assert_eq!(points[0].attrs["idx"],"0");
    assert_eq!(points[1].attrs["idx"],"2");
    assert_eq!(root.child("strCache").unwrap().children_named("pt").next().unwrap().child("v").unwrap().text,"A<&\"'>中文");
}

#[test]
fn byte_budget_rejects_growth_before_append_and_counts_utf8_and_xml_entities() {
    let mut out = "x".repeat(chart_xml::MAX_XML_BYTES-2);
    append_chart_xml(&mut out,"é").unwrap();
    assert_eq!(out.len(),chart_xml::MAX_XML_BYTES);
    assert!(append_chart_xml(&mut out,"x").unwrap_err().message.contains("CHART_XML_RESOURCE_LIMIT"));
    assert_eq!(out.len(),chart_xml::MAX_XML_BYTES);
    let mut escaped = "x".repeat(chart_xml::MAX_XML_BYTES-4);
    assert!(append_chart_escaped(&mut escaped,"&").unwrap_err().message.contains("CHART_XML_RESOURCE_LIMIT"));
    assert_eq!(escaped.len(),chart_xml::MAX_XML_BYTES-4);
}

#[test]
fn numeric_caches_refuse_nonfinite_values_instead_of_writing_invalid_native_numbers() {
    for value in [f64::NAN,f64::INFINITY,f64::NEG_INFINITY] {
        let mut out = String::new();
        assert!(append_number_reference(&mut out,"val","'Data'!$B$2",&[Some(value)],"General").is_err());
    }
}

#[test]
fn series_append_respects_the_shared_part_budget_instead_of_building_a_second_full_series() {
    let book = workbook();
    let chart = &book.charts[0];
    let context = context(&book, chart).unwrap();
    let mut out = "x".repeat(chart_xml::MAX_XML_BYTES - 120);
    let error = append_series(&mut out, &context, chart, &spec(), 0).unwrap_err();
    assert!(error.message.contains("CHART_XML_RESOURCE_LIMIT"));
    assert!(out.len() <= chart_xml::MAX_XML_BYTES);
}
