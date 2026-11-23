//! Formula-backed dates for native category and XY chart axes. The workbook,
//! rather than an authored chart option, owns the serial epoch.

use serde_json::Value;
use office_io_common::chart_kind::NativeChartKind;
use office_io_common::excel_date::{ChartDate, DateZone, chart_date, max_date_serial, valid_date_serial};
use crate::ExportError;
use super::ChartInput;

pub(super) fn enabled(chart: &ChartInput, kind: NativeChartKind) -> bool {
    chart.options.mapping.as_ref().and_then(|mapping| mapping.use_date_axis) == Some(true)
        && !kind.pie && kind.plot != "radarChart"
}

pub(super) fn validate_axis(chart: &ChartInput, date1904: bool) -> Result<(), ExportError> {
    if let Some(axis) = &chart.x_axis {
        if axis.scale_type != "linear" { return Err(ExportError::invalid("CHART_DATE_AXIS_SCALE: date axes require a linear elapsed-time scale")); }
        if [axis.min, axis.max].into_iter().flatten().any(|value| !valid_date_serial(value, date1904)) {
            return Err(ExportError::invalid("CHART_DATE_AXIS_LIMIT: date axis limits must belong to the workbook's date system"));
        }
    }
    Ok(())
}

pub(super) fn read(value: Option<&Value>, date1904: bool) -> Result<Option<ChartDate>, ExportError> {
    match value {
        None | Some(Value::Null) => Ok(None),
        Some(Value::String(text)) if text.trim().is_empty() => Ok(None),
        Some(value) => chart_date(value, date1904).map(Some).ok_or_else(|| ExportError::invalid(
            "CHART_DATE_AXIS_VALUE: date categories require a valid Excel serial or strict ISO calendar/clock text")),
    }
}

/// Date text keeps a formula binding to the original source. Its syntax layout
/// is recorded when preparing the helper, just as grouped helpers record their
/// current groups. Native edits within that layout recalculate; changing the
/// date text layout requires regenerating the helper in the workbench.
pub(super) fn expression(reference: &str, parsed: Option<ChartDate>, date1904: bool) -> String {
    let limit = max_date_serial(date1904) + 1.0;
    let text = format!("TRIM({reference})");
    let numeric = numeric_expression(reference, &text, limit);
    let Some(layout) = parsed.and_then(|date| date.iso) else { return numeric; };
    let year = format!("VALUE(LEFT({text},4))");
    let month = format!("VALUE(MID({text},6,2))");
    let day = format!("VALUE(MID({text},9,2))");
    let mut conditions = vec![
        format!("LEN({text})={}", 10 + usize::from(layout.time) * 6 + usize::from(layout.seconds) * 3
            + if layout.fractional_digits > 0 { layout.fractional_digits + 1 } else { 0 }
            + match layout.zone { DateZone::None => 0, DateZone::Utc => 1, DateZone::Offset => 6 }),
        format!("MID({text},5,1)=\"-\""), format!("MID({text},8,1)=\"-\""),
        exact_digits(&format!("LEFT({text},4)"), 4),
        exact_digits(&format!("MID({text},6,2)"), 2), exact_digits(&format!("MID({text},9,2)"), 2),
        format!("{year}>={}", if date1904 { 1904 } else { 1900 }), format!("{year}<=9999"),
    ];
    let native_date = format!("DATE({year},{month},{day})");
    let calendar = format!("AND(YEAR({native_date})={year},MONTH({native_date})={month},DAY({native_date})={day})");
    let date = if date1904 {
        conditions.push(calendar); native_date
    } else {
        let zero = format!("AND({year}=1900,{month}=1,{day}=0)");
        let ghost = format!("AND({year}=1900,{month}=2,{day}=29)");
        conditions.push(format!("OR({zero},{ghost},{calendar})"));
        format!("IF({zero},0,IF({ghost},60,{native_date}))")
    };
    let mut serial = date;
    if layout.time {
        conditions.extend([format!("OR(MID({text},11,1)=\"T\",MID({text},11,1)=\" \" )"),
            format!("MID({text},14,1)=\":\""), exact_digits(&format!("MID({text},12,2)"), 2), exact_digits(&format!("MID({text},15,2)"), 2)]);
        let hours = format!("VALUE(MID({text},12,2))");
        let minutes = format!("VALUE(MID({text},15,2))");
        conditions.extend([format!("{hours}<24"), format!("{minutes}<60")]);
        let seconds = if layout.seconds {
            conditions.extend([format!("MID({text},17,1)=\":\""), exact_digits(&format!("MID({text},18,2)"), 2), format!("VALUE(MID({text},18,2))<60")]);
            format!("VALUE(MID({text},18,2))")
        } else { "0".into() };
        let fraction = if layout.fractional_digits > 0 {
            let digits = format!("MID({text},21,{})", layout.fractional_digits);
            conditions.extend([format!("MID({text},20,1)=\".\""), exact_digits(&digits, layout.fractional_digits)]);
            format!("VALUE({digits})/10^{}", layout.fractional_digits)
        } else { "0".into() };
        serial = format!("({serial}+({hours}*3600+{minutes}*60+{seconds}+{fraction})/86400)");
    }
    match layout.zone {
        DateZone::None => {},
        DateZone::Utc => conditions.push(format!("RIGHT({text},1)=\"Z\"")),
        DateZone::Offset => {
            let sign = format!("MID({text},LEN({text})-5,1)");
            let hours = format!("MID({text},LEN({text})-4,2)");
            let minutes = format!("RIGHT({text},2)");
            conditions.extend([format!("OR({sign}=\"+\",{sign}=\"-\")"), format!("MID({text},LEN({text})-2,1)=\":\""),
                exact_digits(&hours, 2), exact_digits(&minutes, 2), format!("VALUE({hours})<24"), format!("VALUE({minutes})<60")]);
            serial = format!("({serial}-IF({sign}=\"-\",-1,1)*(VALUE({hours})*60+VALUE({minutes}))/1440)");
        },
    }
    let iso = format!("IFERROR(IF(AND({}),IF(AND({serial}>=0,{serial}<{limit}),{serial},NA()),NA()),NA())", conditions.join(","));
    // Numeric edits stay valid in either date system. Booleans and blanks are
    // always missing, including when Y's invalid-value strategy is zero.
    format!("IF(ISNUMBER({reference}),{numeric},{iso})")
}

fn exact_digits(expression: &str, count: usize) -> String {
    format!("EXACT({expression},TEXT(VALUE({expression}),\"{}\"))", "0".repeat(count))
}

fn numeric_expression(reference: &str, text: &str, limit: f64) -> String {
    let mut remaining = text.to_owned();
    for character in ["0", "1", "2", "3", "4", "5", "6", "7", "8", "9", ".", "+", "-", "e", "E"] {
        remaining = format!("SUBSTITUTE({remaining},\"{character}\",\"\")");
    }
    let value = format!("VALUE({text})");
    format!("IFERROR(IF(ISNUMBER({reference}),IF(AND({reference}>=0,{reference}<{limit}),{reference},NA()),IF(OR(ISBLANK({reference}),{reference}=\"\",ISLOGICAL({reference})),NA(),IF(AND({remaining}=\"\",{value}>=0,{value}<{limit}),{value},NA()))),NA())")
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;

    #[test]
    fn zero_y_strategy_does_not_turn_missing_or_invalid_dates_into_epoch_zero() {
        assert_eq!(read(Some(&Value::Null), false).unwrap(), None);
        assert_eq!(read(Some(&json!("  ")), true).unwrap(), None);
        assert!(read(Some(&json!(true)), false).is_err());
        assert!(read(Some(&json!("2026-02-30")), false).is_err());
        assert!(expression("'Data'!$A$2", None, false).contains("NA()"));
    }

    #[test]
    fn iso_helpers_reference_source_calendar_clock_fraction_and_zone_components() {
        let source = "'Long worksheet name'!$A$2";
        let date = read(Some(&json!("1904-01-01T08:30:00.125+08:30")), true).unwrap();
        let formula = expression(source, date, true);
        assert!(formula.contains("DATE(VALUE(LEFT("));
        assert!(formula.contains("YEAR(DATE("));
        assert!(formula.contains("/10^3"));
        assert!(formula.contains("/1440"));
        assert!(formula.contains(source));
        assert!(formula.len() <= 8192);
        assert!(!formula.contains("DATEVALUE("));
    }
}
