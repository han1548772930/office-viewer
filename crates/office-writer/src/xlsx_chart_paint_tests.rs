//! Draft actual DrawingML contracts. No native visual equivalence is inferred.

use super::*;
use office_io_common::chart_xml::{self as native_xml, Node};
use serde_json::{Value, json};

fn gradient(start: f64, end: f64, stops: Value) -> ChartColor {
    serde_json::from_value(json!({"type":"linear-gradient","start":{"x":start,"y":0},"end":{"x":end,"y":0},"stops":stops})).unwrap()
}

fn shape(source: &str) -> Node {
    native_xml::parse(&format!("<c:spPr xmlns:c=\"{}\" xmlns:a=\"{}\">{source}</c:spPr>",native_xml::CHART_NS,native_xml::DRAWING_NS)).unwrap()
}

fn workbook(kind: &str, options: Value) -> super::super::WorkbookInput {
    serde_json::from_value(json!({"kind":"workbook","name":"Paints","sheets":[{"id":"data","name":"Data","rowCount":3,"columnCount":3,
        "cells":{"0:0":{"value":"Category"},"0:1":{"value":"A"},"0:2":{"value":"B"},
            "1:0":{"value":"One"},"1:1":{"value":-5},"1:2":{"value":3},"2:0":{"value":"Two"},"2:1":{"value":2},"2:2":{"value":4}}}],
        "charts":[{"id":"paint","sheetId":"data","type":kind,"categoryField":0,"valueFields":[1],
            "source":{"orientation":"columns","range":{"startRow":0,"startCol":0,"endRow":2,"endCol":2}},
            "position":{"x":0,"y":0},"size":{"width":480,"height":300},"options":options}]})).unwrap()
}

fn chart(book: &super::super::WorkbookInput) -> Node {
    native_xml::parse(&super::super::chart_xml(book,&book.charts[0],0,None).unwrap()).unwrap()
}

#[test]
fn native_stops_sort_stably_and_compose_css_stop_and_fill_alpha() {
    let color = gradient(0.0,1.0,json!([{"offset":0.8,"color":"blue"},{"offset":0.2,"color":"rgba(255,0,0,0.5)","opacity":0.4},{"offset":0.2,"color":"green"}]));
    let before = color.clone();
    let native = shape(&fill(&color,Some(0.5),None).unwrap());
    let stops = native.descendants("gs");
    assert_eq!(stops.iter().map(|stop| stop.attrs["pos"].as_str()).collect::<Vec<_>>(),["20000","20000","80000"]);
    assert_eq!(stops[0].child("srgbClr").unwrap().attrs["val"],"FF0000");
    assert_eq!(stops[0].child("srgbClr").unwrap().value("alpha"),Some("10000"));
    assert_eq!(stops[1].child("srgbClr").unwrap().attrs["val"],"008000");
    assert_eq!(color,before);
}

#[test]
fn out_of_mark_stops_interpolate_edges_instead_of_clamping_source_colors() {
    let color = gradient(-1.0,2.0,json!([{"offset":0,"color":"red"},{"offset":1,"color":"blue"}]));
    let native = shape(&fill(&color,None,None).unwrap());
    let stops = native.descendants("gs");
    assert_eq!(stops.len(),2);
    assert_eq!(stops[0].attrs["pos"],"0");
    assert_eq!(stops[0].child("srgbClr").unwrap().attrs["val"],"AA0055");
    assert_eq!(stops[1].attrs["pos"],"100000");
    assert_eq!(stops[1].child("srgbClr").unwrap().attrs["val"],"5500AA");
}

#[test]
fn coincident_endpoints_use_no_fill_and_invalid_native_geometry_is_rejected() {
    let stops = json!([{"offset":0,"color":"red"},{"offset":1,"color":"blue"}]);
    assert_eq!(fill(&gradient(0.5,0.5,stops.clone()),None,None).unwrap(),"<a:noFill/>");
    assert!(fill(&gradient(-f64::MAX,f64::MAX,stops),None,None).is_err());
}

#[test]
fn native_pattern_overrides_the_gradient_and_keeps_both_color_alphas() {
    let color = gradient(0.0,1.0,json!([{"offset":0,"color":"red"},{"offset":1,"color":"blue"}]));
    let pattern = FillPattern { preset:"diagCross".into(),foreground_color:"rgba(255,0,0,0.5)".into(),background_color:"rgba(0,0,255,0.25)".into() };
    let native = shape(&fill(&color,Some(0.4),Some(&pattern)).unwrap());
    assert!(native.descendants("gradFill").is_empty());
    let pattern = native.child("pattFill").unwrap();
    assert_eq!(pattern.attrs["prst"],"diagCross");
    assert_eq!(pattern.child("fgClr").unwrap().child("srgbClr").unwrap().value("alpha"),Some("20000"));
    assert_eq!(pattern.child("bgClr").unwrap().child("srgbClr").unwrap().value("alpha"),Some("10000"));
}

#[test]
fn automatic_defaults_and_negative_points_change_native_paint_without_persisting_derived_stops() {
    let bar = workbook("bar",json!({}));
    assert!(chart(&bar).descendants("gradFill").is_empty());
    let area = workbook("area",json!({}));
    assert!(!chart(&area).descendants("gradFill").is_empty());
    assert!(area.charts[0].options.auto_gradient_fill.is_none());
    let enabled = workbook("bar",json!({"autoGradientFill":true}));
    let native = chart(&enabled);
    let series = native.descendants("ser")[0];
    assert_eq!(series.child("spPr").unwrap().descendants("lin")[0].attrs["ang"],"16200000");
    let point = series.children_named("dPt").next().unwrap();
    assert_eq!(point.value("idx"),Some("0"));
    assert_eq!(point.descendants("lin")[0].attrs["ang"],"5400000");
    assert!(enabled.charts[0].options.series.is_empty());
}

#[test]
fn explicit_gradients_and_patterns_do_not_add_unneeded_negative_point_overrides() {
    for series in [json!({"color":{"type":"linear-gradient","start":{"x":0,"y":0},"end":{"x":1,"y":0},
        "stops":[{"offset":0,"color":"red"},{"offset":1,"color":"blue"}]}}),
        json!({"fillPattern":{"preset":"cross","foregroundColor":"red","backgroundColor":"white"}})] {
        let book = workbook("bar",json!({"autoGradientFill":true,"series":{"1":series}}));
        assert!(chart(&book).descendants("dPt").is_empty());
    }
}

#[test]
fn secondary_axis_reverse_changes_only_its_series_automatic_gradient() {
    let mut book = workbook("bar",json!({"autoGradientFill":true,"series":{"2":{"rightAxis":true}},"rightYAxis":{"scaleType":"linear","reverse":true}}));
    book.charts[0].value_fields = vec![1,2];
    let native = chart(&book);
    let series = native.descendants("ser");
    assert_eq!(series[0].child("spPr").unwrap().descendants("lin")[0].attrs["ang"],"16200000");
    assert_eq!(series[1].child("spPr").unwrap().descendants("lin")[0].attrs["ang"],"5400000");
}

#[test]
fn manual_plot_layout_is_inner_edge_origin_and_factor_sized() {
    let book = workbook("bar",json!({"plotArea":{"left":0.1,"top":0.2,"width":0.7,"height":0.6}}));
    let native = chart(&book);
    let layout = native.descendants("manualLayout")[0];
    assert_eq!(layout.value("layoutTarget"),Some("inner"));
    assert_eq!(layout.value("xMode"),Some("edge"));
    assert_eq!(layout.value("yMode"),Some("edge"));
    assert_eq!(layout.value("wMode"),Some("factor"));
    assert_eq!(layout.value("x"),Some("0.1"));
}

#[test]
fn explicit_null_point_pattern_removes_only_the_inherited_texture() {
    let base: Series = serde_json::from_value(json!({"fillPattern":{"preset":"cross","foregroundColor":"red","backgroundColor":"white"},"dataPoints":{"0":{"fillPattern":{"preset":"vert","foregroundColor":"blue","backgroundColor":"white"}}}})).unwrap();
    let patch: Series = serde_json::from_value(json!({"dataPoints":{"0":{"fillPattern":null}}})).unwrap();
    let merged: Series = style::overlay(Some(&base),Some(&patch)).unwrap();
    assert!(merged.fill_pattern.as_ref().unwrap().is_some());
    assert_eq!(merged.data_points[&0].fill_pattern,Some(None));
}
