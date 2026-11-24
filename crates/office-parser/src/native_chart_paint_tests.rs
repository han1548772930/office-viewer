//! Draft native-reader contracts. Original source XML remains independently traceable.

use super::*;
use office_io_common::chart_xml;

fn shape(body: &str) -> Node {
    chart_xml::parse(&format!("<c:spPr xmlns:c=\"{CHART_NS}\" xmlns:a=\"{DRAWING_NS}\" xmlns:x=\"urn:extension\">{body}</c:spPr>")).unwrap()
}

#[test]
fn linear_gradient_preserves_native_stop_order_transforms_and_opacity() {
    let native = shape("<a:gradFill><a:gsLst><a:gs pos=\"100000\"><a:srgbClr val=\"0000FF\"/></a:gs><a:gs pos=\"0\"><a:srgbClr val=\"FF0000\"><a:alpha val=\"50000\"/></a:srgbClr></a:gs></a:gsLst><a:lin ang=\"16200000\" scaled=\"1\"/></a:gradFill>");
    let paint = read(Some(&native),&HashMap::new()).unwrap();
    let ChartColor::LinearGradient(gradient) = paint.color.unwrap() else { panic!("linear fill was flattened"); };
    assert_eq!(gradient.stops[0].offset,1.0);
    assert_eq!(gradient.stops[1].offset,0.0);
    assert_eq!(gradient.stops[1].color,"rgba(255, 0, 0, 0.5)");
    assert_eq!((gradient.start.y,gradient.end.y),(1.0,0.0));
    assert_eq!(paint.opacity,None);
}

#[test]
fn native_pattern_keeps_preset_both_alpha_colors_and_theme_resolution() {
    let native = shape("<a:pattFill prst=\"diagCross\"><a:fgClr><a:schemeClr val=\"accent1\"><a:alpha val=\"50000\"/></a:schemeClr></a:fgClr><a:bgClr><a:srgbClr val=\"0000FF\"><a:alpha val=\"25000\"/></a:srgbClr></a:bgClr></a:pattFill>");
    let paint = read(Some(&native),&HashMap::from([("accent1".into(),"#ff0000".into())])).unwrap();
    let pattern = paint.pattern.unwrap();
    assert_eq!(pattern.preset,"diagCross");
    assert_eq!(pattern.foreground_color,"rgba(255, 0, 0, 0.5)");
    assert_eq!(pattern.background_color,"rgba(0, 0, 255, 0.25)");
}

#[test]
fn namespace_collisions_cannot_replace_fills_colors_or_color_transforms() {
    let native = shape("<x:gradFill><x:lin/></x:gradFill><a:solidFill><x:srgbClr val=\"000000\"/><a:srgbClr val=\"FF0000\"><x:alpha val=\"0\"/></a:srgbClr></a:solidFill>");
    let paint = read(Some(&native),&HashMap::new()).unwrap();
    assert_eq!(paint.color.as_ref().and_then(ChartColor::as_solid),Some("#ff0000"));
    assert_eq!(paint.opacity,Some(1.0));
}

#[test]
fn transparent_native_fill_does_not_inherit_the_palette() {
    let paint = read(Some(&shape("<a:noFill/>")),&HashMap::new()).unwrap();
    assert_eq!(paint.color.as_ref().and_then(ChartColor::as_solid),Some("transparent"));
}

#[test]
fn invalid_stops_angles_and_path_gradients_report_specific_mapping_errors() {
    let stops = "<a:gsLst><a:gs pos=\"0\"><a:srgbClr val=\"FF0000\"/></a:gs><a:gs pos=\"100000\"><a:srgbClr val=\"0000FF\"/></a:gs></a:gsLst>";
    let path = shape(&format!("<a:gradFill>{stops}<a:path path=\"circle\"/></a:gradFill>"));
    assert_eq!(read(Some(&path),&HashMap::new()).err().unwrap(),"CHART_NATIVE_GRADIENT_PATH_UNMAPPED");
    let invalid = shape(&format!("<a:gradFill>{stops}<a:lin ang=\"NaN\"/></a:gradFill>"));
    assert!(read(Some(&invalid),&HashMap::new()).is_err());
    let invalid = shape("<a:gradFill><a:gsLst><a:gs pos=\"100001\"><a:srgbClr val=\"FF0000\"/></a:gs><a:gs pos=\"0\"><a:srgbClr val=\"0000FF\"/></a:gs></a:gsLst><a:lin/></a:gradFill>");
    assert!(read(Some(&invalid),&HashMap::new()).is_err());
}

#[test]
fn inner_edge_factor_plot_layout_maps_and_default_relative_layout_is_retained_as_a_difference() {
    let body = "<c:layout><c:manualLayout><c:layoutTarget val=\"inner\"/><c:xMode val=\"edge\"/><c:yMode val=\"edge\"/><c:x val=\"0.1\"/><c:y val=\"0.2\"/><c:w val=\"0.7\"/><c:h val=\"0.6\"/></c:manualLayout></c:layout>";
    let plot = chart_xml::parse(&format!("<c:plotArea xmlns:c=\"{CHART_NS}\">{body}</c:plotArea>")).unwrap();
    assert_eq!(plot_area(&plot).unwrap(),Some(PlotArea { left:0.1,top:0.2,width:0.7,height:0.6 }));
    let outer = body.replace("val=\"inner\"","val=\"outer\"");
    let root = chart_xml::parse(&format!("<c:chartSpace xmlns:c=\"{CHART_NS}\"><c:chart><c:plotArea>{outer}</c:plotArea></c:chart></c:chartSpace>")).unwrap();
    let plot = root.child("chart").unwrap().child("plotArea").unwrap();
    assert_eq!(plot_area(plot).unwrap(),None);
    assert_eq!(source_differences(&root,false),["CHART_NATIVE_DEFAULT_RELATIVE_OR_OUTER_LAYOUT"]);
}

#[test]
fn unscaled_tile_and_patterned_stroke_differences_are_deduplicated() {
    let native = shape("<a:gradFill><a:lin ang=\"2700000\" scaled=\"0\"/><a:tileRect l=\"1000\"/></a:gradFill><a:ln><a:pattFill prst=\"cross\"/></a:ln>");
    assert_eq!(source_differences(&native,false),["CHART_NATIVE_GRADIENT_TILE_GEOMETRY","CHART_NATIVE_PATTERNED_LINE_STROKE","CHART_NATIVE_UNSCALED_GRADIENT_GEOMETRY"]);
}
