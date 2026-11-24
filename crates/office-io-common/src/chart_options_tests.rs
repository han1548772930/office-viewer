//! Draft authored paint DTO contracts, including inactive settings.

use super::*;
use serde_json::json;

#[test]
fn missing_pattern_inherits_but_explicit_null_survives_serialization() {
    let inherited: Series = serde_json::from_value(json!({})).unwrap();
    let reset: Series = serde_json::from_value(json!({"fillPattern":null,"dataPoints":{"2":{"fillPattern":null}}})).unwrap();
    assert!(inherited.fill_pattern.is_none());
    assert_eq!(reset.fill_pattern, Some(None));
    assert_eq!(serde_json::to_value(&reset).unwrap()["fillPattern"], Value::Null);
    assert_eq!(reset.data_points[&2].fill_pattern, Some(None));
    assert!(serde_json::to_value(&inherited).unwrap().get("fillPattern").is_none());
}

#[test]
fn every_fixed_pattern_and_both_colors_validate_at_the_authored_boundary() {
    assert_eq!(FILL_PATTERN_PRESETS.len(), 54);
    for preset in FILL_PATTERN_PRESETS {
        FillPattern { preset: (*preset).into(), foreground_color: "rgba(255,0,0,0.5)".into(), background_color: "transparent".into() }.validate().unwrap();
    }
    for pattern in [json!({"preset":"future","foregroundColor":"red","backgroundColor":"white"}),
        json!({"preset":"cross","foregroundColor":"broken","backgroundColor":"white"})] {
        let options: Options = serde_json::from_value(json!({"namedSeries":[{"name":"inactive","style":{"fillPattern":pattern}}]})).unwrap();
        assert!(options.validate_paints().is_err());
    }
}

#[test]
fn inactive_series_and_point_gradients_are_validated_without_activation() {
    let gradient = json!({"type":"linear-gradient","start":{"x":0,"y":0},"end":{"x":1,"y":0},
        "stops":[{"offset":0,"color":"red"},{"offset":1,"color":"blue","opacity":1.1}]});
    for options in [json!({"series":{"999":{"color":gradient}}}),
        json!({"namedSeries":[{"name":"future","style":{"dataPoints":{"500":{"color":gradient}}}}]})] {
        assert!(serde_json::from_value::<Options>(options).unwrap().validate_paints().is_err());
    }
}

#[test]
fn plot_area_requires_finite_positive_dimensions_inside_the_chart_frame() {
    PlotArea { left:0.1,top:0.2,width:0.8,height:0.7 }.validate().unwrap();
    for area in [PlotArea { left:0.0,top:0.0,width:0.0,height:1.0 },
        PlotArea { left:0.2,top:0.1,width:0.9,height:0.5 },
        PlotArea { left:f64::NAN,top:0.0,width:1.0,height:1.0 }] { assert!(area.validate().is_err()); }
    assert!(serde_json::from_value::<PlotArea>(json!({"left":0,"top":0,"width":1,"height":1,"renderer":true})).is_err());
}
