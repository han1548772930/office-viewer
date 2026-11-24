//! Draft native writer -> independent parser contracts for unified validation.

use serde_json::{Value, json};
use crate::parser::parse_office_bytes;

fn reopen(model: Value) -> Value {
    let (bytes, _) = office_writer::export_xlsx_bytes(&model.to_string()).unwrap();
    serde_json::to_value(parse_office_bytes(&bytes, "native-chart.xlsx").unwrap()).unwrap()["payload"]["data"].clone()
}

fn workbook(chart_type: &str, options: Value, values: &[usize]) -> Value {
    json!({"kind":"workbook","name":"Native charts","sheets":[{"id":"data","name":"Data","rowCount":3,"columnCount":4,
        "cells":{"0:0":{"value":"Category"},"0:1":{"value":"Revenue"},"0:2":{"value":"Cost"},"0:3":{"value":"Error"},
            "1:0":{"value":"A"},"1:1":{"value":10},"1:2":{"value":6},"1:3":{"value":1},
            "2:0":{"value":"B"},"2:1":{"value":20},"2:2":{"value":8},"2:3":{"value":2}}}],
        "charts":[{"id":"original-chart","sheetId":"data","type":chart_type,"categoryField":0,"valueFields":values,
            "source":{"orientation":"columns","range":{"startRow":0,"startCol":0,"endRow":2,"endCol":3}},
            "position":{"x":37,"y":51},"size":{"width":480,"height":300},"title":"Native title","subtitle":"Native subtitle",
            "legend":{"visible":true,"position":"right","font":{"family":"Arial","fontSize":14}},
            "xAxis":{"scaleType":"linear","title":"Category"},"yAxis":{"scaleType":"linear","title":"Revenue"},
            "options":options}]})
}

#[test]
fn native_date_axes_and_title_fonts_reopen_with_the_matching_authored_pointer_settings() {
    let mut model = workbook("line", json!({"mapping":{"useDateAxis":true},"axisPointer":{
        "indicatorLabelColor":"#000000","indicatorLineType":"dotted","indicatorLineColor":"#ff0000","indicatorLabelTextColor":"#ffffff"}}), &[1]);
    model["date1904"] = json!(true);
    model["sheets"][0]["cells"]["1:0"]["value"] = json!("1904-01-01");
    model["sheets"][0]["cells"]["2:0"]["value"] = json!(2);
    model["charts"][0]["xAxis"]["titleStyle"] = json!({"fontSize":16,"family":"Arial","color":"#ff0000"});
    model["charts"][0]["xAxis"]["titleAlignment"] = json!("end");
    let data = reopen(model);
    assert_eq!(data["date1904"], true);
    let chart = &data["sheets"][0]["charts"][0];
    assert_eq!(chart["options"]["mapping"]["useDateAxis"], true);
    assert_eq!(chart["xAxis"]["titleStyle"]["family"], "Arial");
    assert_eq!(chart["xAxis"]["titleStyle"]["fontSize"], 16.0);
    assert_eq!(chart["authored"]["xAxis"]["titleAlignment"], "end");
    assert_eq!(chart["authored"]["options"]["axisPointer"]["indicatorLineType"], "dotted");
    assert!(chart["series"][0]["categories"].as_str().unwrap().contains("__ov_chart_"));
}

#[test]
fn mixed_native_plots_keep_series_formulas_axes_statistics_and_authored_controls() {
    let data = reopen(workbook("combination", json!({
        "series":{"1":{"type":"column","color":"#ff0000"},"2":{"type":"line","rightAxis":true,"line":{"width":3}}},
        "rightYAxis":{"scaleType":"linear","title":"Cost","min":0},
        "trendlines":[{"seriesIndex":2,"type":"linear","showFormula":true}],
        "errorBars":[{"seriesIndex":1,"direction":"y","type":"both","valueType":"cust","plusIndex":3,"minusIndex":3}],
        "customSettings":{"retained":true}
    }), &[1,2]));
    let chart = &data["sheets"][0]["charts"][0];
    assert_eq!(chart["kind"], "combination");
    assert_eq!(chart["series"][0]["values"], "'Data'!$B$2:$B$3");
    assert_eq!(chart["series"][1]["values"], "'Data'!$C$2:$C$3");
    assert_eq!(chart["series"][1]["options"]["rightAxis"], true);
    assert_eq!(chart["series"][1]["trendlines"][0]["seriesIndex"], 1);
    assert_eq!(chart["series"][0]["errorBars"][0]["plusFormula"], "'Data'!$D$2:$D$3");
    assert_eq!(chart["authored"]["options"]["customSettings"]["retained"], true);
    assert_eq!(chart["authored"]["valueFields"], json!([1,2]));
    assert_eq!(chart["drawing"]["fromOffsetEmu"], json!([37*9525,51*9525]));
    assert_eq!(chart["drawing"]["absoluteSizeEmu"], json!([480*9525,300*9525]));
}

#[test]
fn grouped_bubble_helpers_are_real_formula_sources_and_reopen_to_the_authored_range() {
    let mut model = workbook("bubble", json!({"invalidValueStrategy":"zero",
        "mapping":{"bubble":{"xIndex":1,"yIndex":2,"seriesIndex":0}}}), &[2]);
    model["sheets"][0]["cells"]["1:2"]["value"] = Value::Null;
    let data = reopen(model);
    let chart = &data["sheets"][0]["charts"][0];
    assert_eq!(chart["series"].as_array().unwrap().len(), 2);
    assert!(chart["series"][0]["values"].as_str().unwrap().contains("__ov_chart_"));
    assert!(chart["series"][0]["bubbleSizes"].as_str().unwrap().contains("__ov_chart_"));
    assert_eq!(chart["authored"]["source"]["sheetName"], "Data");
    assert_eq!(chart["authored"]["options"]["invalidValueStrategy"], "zero");
    assert_eq!(chart["authored"]["helperSheets"].as_array().unwrap().len(), 1);
    let helper = data["sheets"].as_array().unwrap().iter().find(|sheet| sheet["name"] == "__ov_chart_1").unwrap();
    assert!(helper["cells"].as_array().unwrap().iter().any(|cell| cell["formula"].as_str().is_some_and(|formula| formula.contains("EXACT("))));
}

#[test]
fn source_name_styles_drive_native_combo_plots_and_survive_the_independent_reader() {
    let named = json!([{ "name":"Cost", "style":{"name":"Expenses","type":"line","rightAxis":true,"color":"#ff0000","border":{"width":3}} },
        { "name":"Not active", "style":{"futureSetting":{"retain":true}} }]);
    let data = reopen(workbook("combination", json!({"namedSeries":named,
        "allSeriesStyle":{"label":{"fontSize":16}},"series":{"2":{"label":{"color":"#123456"}}},
        "rightYAxis":{"scaleType":"linear","title":"Expenses"}}), &[1,2]));
    let chart = &data["sheets"][0]["charts"][0];
    let expense = chart["series"].as_array().unwrap().iter().find(|series| series["name"] == "Expenses").unwrap();
    assert_eq!(expense["options"]["rightAxis"], true);
    assert_eq!(expense["options"]["color"], "#ff0000");
    assert_eq!(expense["values"], "'Data'!$C$2:$C$3");
    assert_eq!(expense["options"]["label"]["fontSize"], 16.0);
    assert_eq!(chart["authored"]["options"]["namedSeries"], named);
    assert!(chart["unsupportedReasons"].as_array().is_none_or(|reasons| reasons.iter().all(|reason| reason != "CHART_AUTHORED_SOURCE_STALE")));
}

#[test]
fn named_bubble_groups_on_two_axes_remain_one_native_chart_type() {
    let named = json!([{ "name":"A", "style":{"name":"North","rightAxis":true,"color":"#ff0000"} },
        { "name":"B", "style":{"color":"#00ff00"} }]);
    let data = reopen(workbook("bubble", json!({"mapping":{"bubble":{"xIndex":1,"yIndex":2,"seriesIndex":0}},
        "namedSeries":named}), &[2]));
    let chart = &data["sheets"][0]["charts"][0];
    assert_eq!(chart["kind"], "bubble");
    let north = chart["series"].as_array().unwrap().iter().find(|series| series["name"] == "North").unwrap();
    assert_eq!(north["options"]["rightAxis"], true);
    assert_eq!(north["options"]["color"], "#ff0000");
    let south = chart["series"].as_array().unwrap().iter().find(|series| series["name"] == "B").unwrap();
    assert_eq!(south["options"]["rightAxis"], false);
    assert_eq!(south["options"]["color"], "#00ff00");
    assert_eq!(chart["authored"]["options"]["namedSeries"], named);
}

#[test]
fn native_linear_gradient_alpha_and_plot_layout_reopen_without_changing_author_paint_order() {
    let options = json!({"autoGradientFill":false,"plotArea":{"left":0.1,"top":0.2,"width":0.7,"height":0.6},
        "series":{"1":{"color":{"type":"linear-gradient","start":{"x":0,"y":0},"end":{"x":1,"y":0},
            "stops":[{"offset":1,"color":"blue"},{"offset":0,"color":"rgba(255,0,0,0.5)","opacity":0.4}]},"fillOpacity":0.5}},
        "customSettings":{"keep":"original"}});
    let data = reopen(workbook("bar",options.clone(),&[1]));
    let chart = &data["sheets"][0]["charts"][0];
    let paint = &chart["series"][0]["options"]["color"];
    assert_eq!(paint["type"],"linear-gradient");
    assert_eq!(paint["start"]["x"],0.0);
    assert_eq!(paint["end"]["x"],1.0);
    assert_eq!(paint["stops"][0]["offset"],0.0);
    assert_eq!(paint["stops"][0]["color"],"rgba(255, 0, 0, 0.1)");
    assert_eq!(paint["stops"][1]["color"],"rgba(0, 0, 255, 0.5)");
    assert!(chart["series"][0]["options"]["fillOpacity"].is_null());
    assert_eq!(chart["options"]["plotArea"],options["plotArea"]);
    assert_eq!(chart["authored"]["options"],options);
    assert_eq!(chart["authored"]["options"]["series"]["1"]["color"]["stops"][0]["offset"],1);
    assert!(chart["unsupportedReasons"].as_array().is_none_or(|reasons| reasons.iter().all(|reason| reason != "CHART_AUTHORED_SOURCE_STALE")));
}

#[test]
fn native_patterns_and_explicit_point_texture_reset_keep_independent_authored_configuration() {
    let options = json!({"series":{"1":{"color":{"type":"linear-gradient","start":{"x":0,"y":0},"end":{"x":1,"y":0},
        "stops":[{"offset":0,"color":"red"},{"offset":1,"color":"blue"}]},"fillOpacity":0.4,
        "fillPattern":{"preset":"diagCross","foregroundColor":"rgba(255,0,0,0.5)","backgroundColor":"rgba(0,0,255,0.25)"},
        "dataPoints":{"0":{"fillPattern":null}}}}});
    let data = reopen(workbook("bar",options.clone(),&[1]));
    let chart = &data["sheets"][0]["charts"][0];
    let style = &chart["series"][0]["options"];
    assert_eq!(style["fillPattern"]["preset"],"diagCross");
    assert_eq!(style["fillPattern"]["foregroundColor"],"rgba(255, 0, 0, 0.2)");
    assert_eq!(style["fillPattern"]["backgroundColor"],"rgba(0, 0, 255, 0.1)");
    assert_eq!(style["dataPoints"]["0"]["color"]["type"],"linear-gradient");
    assert_eq!(chart["authored"]["options"],options);
    assert!(chart["authored"]["options"]["series"]["1"]["dataPoints"]["0"]["fillPattern"].is_null());
}

#[test]
fn automatic_native_gradient_stops_remain_derived_and_do_not_enter_the_author_manifest() {
    let options = json!({"autoGradientFill":true});
    let mut model = workbook("bar",options.clone(),&[1]);
    model["sheets"][0]["cells"]["1:1"]["value"] = json!(-5);
    let data = reopen(model);
    let chart = &data["sheets"][0]["charts"][0];
    assert_eq!(chart["series"][0]["options"]["color"]["type"],"linear-gradient");
    assert_eq!(chart["series"][0]["options"]["dataPoints"]["0"]["color"]["type"],"linear-gradient");
    assert_eq!(chart["authored"]["options"],options);
    assert!(chart["authored"]["options"]["series"].is_null());
}
