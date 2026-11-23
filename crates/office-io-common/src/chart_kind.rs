//! Shared native DrawingML chart taxonomy. Local snapshot aliases keep their
//! existing semantics; native barDir/grouping/3D nodes carry the observable type.

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct NativeChartKind {
    pub plot: &'static str,
    pub grouping: Option<&'static str>,
    pub horizontal: bool,
    pub three_d: bool,
    pub pie: bool,
    pub of_pie_type: Option<&'static str>,
}

pub const NATIVE_CHART_PLOTS: &[&str] = &[
    "barChart", "bar3DChart", "lineChart", "line3DChart", "areaChart", "area3DChart",
    "pieChart", "pie3DChart", "doughnutChart", "ofPieChart", "scatterChart", "radarChart",
    "bubbleChart", "stockChart",
];

pub fn native_chart_kind(model: &str) -> Option<NativeChartKind> {
    let (plot, grouping, horizontal, three_d, pie, of_pie_type) = match model {
        "bar" | "column" => ("barChart", Some("clustered"), false, false, false, None),
        "stacked-column" | "columnStacked" => ("barChart", Some("stacked"), false, false, false, None),
        "columnPercentStacked" => ("barChart", Some("percentStacked"), false, false, false, None),
        "column3D" => ("bar3DChart", Some("clustered"), false, true, false, None),
        "columnStacked3D" => ("bar3DChart", Some("stacked"), false, true, false, None),
        "columnPercentStacked3D" => ("bar3DChart", Some("percentStacked"), false, true, false, None),
        "bar-horizontal" => ("barChart", Some("clustered"), true, false, false, None),
        "barStacked" => ("barChart", Some("stacked"), true, false, false, None),
        "barPercentStacked" => ("barChart", Some("percentStacked"), true, false, false, None),
        "bar3D" => ("bar3DChart", Some("clustered"), true, true, false, None),
        "barStacked3D" => ("bar3DChart", Some("stacked"), true, true, false, None),
        "barPercentStacked3D" => ("bar3DChart", Some("percentStacked"), true, true, false, None),
        "line" => ("lineChart", Some("standard"), false, false, false, None),
        "lineStacked" => ("lineChart", Some("stacked"), false, false, false, None),
        "linePercentStacked" => ("lineChart", Some("percentStacked"), false, false, false, None),
        "line3D" => ("line3DChart", Some("standard"), false, true, false, None),
        "area" => ("areaChart", Some("standard"), false, false, false, None),
        "areaStacked" => ("areaChart", Some("stacked"), false, false, false, None),
        "areaPercentStacked" => ("areaChart", Some("percentStacked"), false, false, false, None),
        "area3D" => ("area3DChart", Some("standard"), false, true, false, None),
        "areaStacked3D" => ("area3DChart", Some("stacked"), false, true, false, None),
        "areaPercentStacked3D" => ("area3DChart", Some("percentStacked"), false, true, false, None),
        "pie" => ("pieChart", None, false, false, true, None),
        "pie3D" => ("pie3DChart", None, false, true, true, None),
        "doughnut" => ("doughnutChart", None, false, false, true, None),
        "pieOfPie" => ("ofPieChart", None, false, false, true, Some("pie")),
        "barOfPie" => ("ofPieChart", None, false, false, true, Some("bar")),
        "scatter" => ("scatterChart", None, false, false, false, None),
        "radar" => ("radarChart", None, false, false, false, None),
        "bubble" => ("bubbleChart", None, false, false, false, None),
        "candlestick" => ("stockChart", None, false, false, false, None),
        _ => return None,
    };
    Some(NativeChartKind { plot, grouping, horizontal, three_d, pie, of_pie_type })
}

pub fn model_chart_kind(plot: &str, grouping: Option<&str>, direction: Option<&str>, of_pie_type: Option<&str>) -> Option<&'static str> {
    match plot {
        "barChart" | "bar3DChart" => {
            let horizontal = match direction.unwrap_or("col") { "col" => false, "bar" => true, _ => return None };
            let group = match grouping.unwrap_or("clustered") {
                "standard" | "clustered" => 0, "stacked" => 1, "percentStacked" => 2, _ => return None,
            };
            Some(match (plot == "bar3DChart", horizontal, group) {
                (false, false, 0) => "bar", (false, false, 1) => "stacked-column", (false, false, _) => "columnPercentStacked",
                (false, true, 0) => "bar-horizontal", (false, true, 1) => "barStacked", (false, true, _) => "barPercentStacked",
                (true, false, 0) => "column3D", (true, false, 1) => "columnStacked3D", (true, false, _) => "columnPercentStacked3D",
                (true, true, 0) => "bar3D", (true, true, 1) => "barStacked3D", (true, true, _) => "barPercentStacked3D",
            })
        }
        "lineChart" | "areaChart" | "area3DChart" => Some(match (plot, grouping.unwrap_or("standard")) {
            ("lineChart", "standard") => "line", ("lineChart", "stacked") => "lineStacked", ("lineChart", "percentStacked") => "linePercentStacked",
            ("areaChart", "standard") => "area", ("areaChart", "stacked") => "areaStacked", ("areaChart", "percentStacked") => "areaPercentStacked",
            ("area3DChart", "standard") => "area3D", ("area3DChart", "stacked") => "areaStacked3D", ("area3DChart", "percentStacked") => "areaPercentStacked3D",
            _ => return None,
        }),
        "line3DChart" if grouping.is_none_or(|group| group == "standard") => Some("line3D"),
        "pieChart" => Some("pie"), "pie3DChart" => Some("pie3D"), "doughnutChart" => Some("doughnut"),
        "ofPieChart" => match of_pie_type.unwrap_or("pie") { "pie" => Some("pieOfPie"), "bar" => Some("barOfPie"), _ => None },
        "scatterChart" => Some("scatter"), "radarChart" => Some("radar"), "bubbleChart" => Some("bubble"), "stockChart" => Some("candlestick"),
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn native_direction_grouping_and_depth_preserve_each_observable_chart_variant() {
        for model in ["bar", "column", "stacked-column", "columnStacked", "columnPercentStacked", "column3D", "columnStacked3D", "columnPercentStacked3D",
            "bar-horizontal", "barStacked", "barPercentStacked", "bar3D", "barStacked3D", "barPercentStacked3D", "line", "lineStacked", "linePercentStacked", "line3D",
            "area", "areaStacked", "areaPercentStacked", "area3D", "areaStacked3D", "areaPercentStacked3D", "pie", "pie3D", "doughnut", "pieOfPie", "barOfPie", "scatter", "radar", "bubble", "candlestick"] {
            let native = native_chart_kind(model).unwrap();
            let reopened = model_chart_kind(native.plot, native.grouping, Some(if native.horizontal { "bar" } else { "col" }), native.of_pie_type).unwrap();
            assert_eq!(native_chart_kind(reopened), Some(native), "{model}");
        }
        assert_eq!(model_chart_kind("barChart", Some("percentStacked"), Some("bar"), None), Some("barPercentStacked"));
        assert_eq!(model_chart_kind("area3DChart", Some("stacked"), None, None), Some("areaStacked3D"));
        assert_eq!(model_chart_kind("barChart", None, Some("sideways"), None), None);
        assert_eq!(model_chart_kind("lineChart", Some("unexpected"), None, None), None);
        assert_eq!(native_chart_kind("treemap"), None);
    }
}
