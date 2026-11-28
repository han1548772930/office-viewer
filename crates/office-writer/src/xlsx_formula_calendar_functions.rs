//! Project source-calendar operations to standard spreadsheet formulas.
//! This module builds expressions; the workbook's formula engine and Office
//! remain the evaluators. LET captures each authored argument once.

const GREGORIAN_CYCLE: i32 = 146_097;

fn call(name: &str, arguments: &[String]) -> String {
    format!("{name}({})", arguments.join(","))
}
fn bind(bindings: &[(&str, String)], result: &str) -> String {
    let values = bindings.iter().map(|(name, value)| format!("{name},{value}"))
        .chain(std::iter::once(result.to_owned())).collect::<Vec<_>>();
    format!("_xlfn.LET({})", values.join(","))
}
fn maximum(source: bool) -> i32 { if source { 2_957_003 } else { 2_958_465 } }
fn translated(value: &str, source: bool, target: bool) -> String {
    if source == target { value.to_owned() }
    else if source { format!("({value}+IF({value}<-1401,1461,1462))") }
    else { format!("({value}-IF({value}<60,1461,1462))") }
}

/// A typed date reference may also contain text, booleans or an error after an
/// Office edit. Only numeric results participate in epoch conversion.
pub(super) fn convert(expression: &str, source: bool, target: bool, prefix: &str) -> String {
    if source == target { return expression.to_owned(); }
    let value = format!("{prefix}v");
    bind(&[(&value, expression.to_owned())], &format!("IF(ISNUMBER({value}),{},{value})", translated(&value, source, target)))
}

/// Native 1904 date functions may reject an early negative serial. A complete
/// 400-year Gregorian cycle keeps the month/day/week pattern and supplies a
/// positive calculation date. The 1900 fictitious day is accounted for once.
fn native_fields(value: &str, source: bool, native: bool) -> (String, String) {
    let converted = translated(value, source, native);
    let shift = if native { format!("IF({converted}<0,400,0)") }
        else { format!("IF({converted}<61,400,0)") };
    let extra = if native { String::new() } else { format!("+IF({converted}<60,1,0)") };
    let serial = format!("({converted}+IF({shift}>0,{GREGORIAN_CYCLE}{extra},0))");
    (serial, shift)
}
fn field(expression: &str, name: &str, source: bool, native: bool, prefix: &str) -> String {
    let value = format!("{prefix}s"); let whole = format!("{prefix}i");
    let (serial, shift) = native_fields(&whole, source, native);
    let ordinary = match name {
        "YEAR" => format!("YEAR({serial})-{shift}"),
        "MONTH" => format!("MONTH({serial})"),
        "DAY" => format!("DAY({serial})"),
        _ => unreachable!("calendar field"),
    };
    let result = if source { ordinary } else {
        let zero = if name == "YEAR" { 1900 } else if name == "MONTH" { 1 } else { 0 };
        if name == "DAY" { format!("IF({whole}=0,0,IF({whole}=60,29,{ordinary}))") }
        else { format!("IF({whole}=0,{zero},{ordinary})") }
    };
    bind(&[(&value, format!("--({expression})")), (&whole, format!("INT({value})"))],
        &format!("IF(OR({value}<0,{value}>={}),#NUM!,{result})", maximum(source) + 1))
}

fn date(arguments: &[String], source: bool, native: bool, prefix: &str) -> String {
    let year = format!("{prefix}y"); let month = format!("{prefix}m"); let day = format!("{prefix}d");
    let normalized = format!("{prefix}ny"); let month_index = format!("{prefix}nm");
    let shift = format!("{prefix}shift"); let first = format!("{prefix}first"); let result = format!("{prefix}result");
    let threshold = if native { 1904 } else { 1900 };
    let correction = if native { GREGORIAN_CYCLE } else { GREGORIAN_CYCLE + 1 };
    let first_serial = format!("(DATE({normalized}+{shift},{month_index},1)-IF({shift}>0,{correction},0))");
    let inner = bind(&[
        (&normalized, format!("{year}+IF({year}<1900,1900,0)+INT(({month}-1)/12)")),
        (&month_index, format!("MOD({month}-1,12)+1")),
        (&shift, format!("IF({normalized}<{threshold},400,0)")),
    ], &format!("IF(OR({normalized}<0,{normalized}>9999),#NUM!,{})", bind(&[
        (&first, first_serial),
        (&result, format!("{}+{day}-1", translated(&first, native, source))),
    ], &format!("IF(OR({result}<0,{result}>={}),#NUM!,{result})", maximum(source) + 1))));
    bind(&[(&year, format!("INT({})", arguments[0])), (&month, format!("INT({})", arguments[1])), (&day, format!("INT({})", arguments[2]))],
        &format!("IF(OR({year}<0,{year}>9999,ABS({month})>120000,ABS({day})>1000000),#NUM!,{inner})"))
}

fn weekday(arguments: &[String], source: bool, prefix: &str) -> String {
    let value = format!("{prefix}s"); let kind = format!("{prefix}t"); let day = format!("{prefix}w");
    let first = format!("IF({kind}=2,1,MOD({kind}-10,7))");
    bind(&[
        (&value, format!("--({})", arguments[0])),
        (&kind, format!("INT({})", arguments.get(1).map(String::as_str).unwrap_or("1"))),
        (&day, format!("MOD(INT({value})+{},7)", if source { 5 } else { 6 })),
    ], &format!("IF(OR({value}<0,{value}>{},ISNA(MATCH({kind},{{1,2,3,11,12,13,14,15,16,17}},0))),#NUM!,IF(OR({kind}=1,{kind}=17),{day}+1,IF({kind}=3,MOD({day}+6,7),MOD({day}-{first},7)+1)))", maximum(source)))
}
fn iso_week(expression: &str, source: bool, native: bool, prefix: &str) -> String {
    let value = format!("{prefix}s");
    let (serial, _) = native_fields(&format!("INT({value})"), source, native);
    bind(&[(&value, format!("--({expression})"))], &format!("IF(OR({value}<0,{value}>={}),#NUM!,ISOWEEKNUM({serial}))", maximum(source) + 1))
}
fn weeknum(arguments: &[String], source: bool, native: bool, prefix: &str) -> String {
    let value = format!("{prefix}s"); let kind = format!("{prefix}t"); let year = format!("{prefix}y");
    let first = format!("{prefix}first"); let first_day = format!("{prefix}wd"); let start = format!("{prefix}start");
    let ordinary = bind(&[
        (&year, field(&value, "YEAR", source, native, &format!("{prefix}field_"))),
        (&first, date(&[year.clone(), "1".into(), "1".into()], source, native, &format!("{prefix}date_"))),
        (&first_day, format!("MOD({first}+{},7)", if source { 5 } else { 6 })),
        (&start, format!("IF({kind}=1,0,IF({kind}=2,1,MOD({kind}-10,7)))")),
    ], &format!("INT((INT({value})-({first}-MOD({first_day}-{start},7)))/7)+1"));
    bind(&[(&value, format!("--({})", arguments[0])), (&kind, format!("INT({})", arguments.get(1).map(String::as_str).unwrap_or("1")))],
        &format!("IF({kind}=21,{},IF(ISNA(MATCH({kind},{{1,2,11,12,13,14,15,16,17}},0)),#NUM!,{ordinary}))", iso_week(&value, source, native, &format!("{prefix}iso_"))))
}
fn month_end(arguments: &[String], end: bool, source: bool, native: bool, prefix: &str) -> String {
    let value = format!("{prefix}s"); let offset = format!("{prefix}offset"); let year = format!("{prefix}y");
    let month = format!("{prefix}m"); let day = format!("{prefix}d"); let normalized = format!("{prefix}ny");
    let normalized_month = format!("{prefix}nm"); let last = format!("{prefix}last"); let native_year = format!("{prefix}nativey");
    let final_day = if end { last.clone() } else { format!("MIN({day},{last})") };
    let result = date(&[normalized.clone(), normalized_month.clone(), final_day], source, native, &format!("{prefix}date_"));
    bind(&[
        (&value, format!("--({})", arguments[0])), (&offset, format!("TRUNC({})", arguments[1])),
        (&year, field(&value, "YEAR", source, native, &format!("{prefix}year_"))),
        (&month, field(&value, "MONTH", source, native, &format!("{prefix}month_"))),
        (&day, field(&value, "DAY", source, native, &format!("{prefix}day_"))),
        (&normalized, format!("{year}+INT(({month}-1+{offset})/12)")),
        (&normalized_month, format!("MOD({month}-1+{offset},12)+1")),
    ], &format!("IF(OR({normalized}<{}, {normalized}>9999),#NUM!,{})", if source { 1904 } else { 1900 }, bind(&[
        (&native_year, format!("{normalized}+IF({normalized}<1904,400,0)")),
        (&last, format!("DAY(EOMONTH(DATE({native_year},{normalized_month},1),0))")),
    ], &result)))
}
fn clock(expression: &str, name: &str, prefix: &str) -> String {
    let value = format!("{prefix}s"); let seconds = format!("{prefix}seconds");
    let result = match name {
        "HOUR" => format!("INT({seconds}/3600)"),
        "MINUTE" => format!("INT(MOD({seconds},3600)/60)"),
        "SECOND" => format!("MOD({seconds},60)"),
        _ => unreachable!("clock field"),
    };
    bind(&[(&value, format!("--({expression})")), (&seconds, format!("INT(MOD({value},1)*86400+0.00000001)"))], &result)
}

/// Financial functions consume dates in the workbook calendar. Ordinary
/// numeric parameters remain source quantities; date-returning coupon methods
/// convert the result back into the source frame.
fn native_date_arguments(name: &str) -> Option<(&'static [usize], bool)> {
    Some(match name {
        "COUPDAYBS" | "COUPDAYS" | "COUPDAYSNC" | "COUPNUM" | "DISC" | "DURATION" | "INTRATE" | "MDURATION" |
        "PRICE" | "PRICEDISC" | "RECEIVED" | "YIELD" | "YIELDDISC" => (&[0, 1], false),
        "COUPNCD" | "COUPPCD" => (&[0, 1], true),
        "PRICEMAT" | "YIELDMAT" => (&[0, 1, 2], false),
        "ACCRINT" => (&[0, 1, 2], false),
        "ACCRINTM" => (&[0, 1], false),
        "ODDFPRICE" | "ODDFYIELD" => (&[0, 1, 2, 3], false),
        "ODDLPRICE" | "ODDLYIELD" => (&[0, 1, 2], false),
        "TBILLEQ" | "TBILLPRICE" | "TBILLYIELD" => (&[0, 1], false),
        _ => return None,
    })
}

pub(super) fn project(name: &str, authored_name: &str, arguments: &[String], source: bool, native: bool, prefix: &str) -> String {
    let arity = match name {
        "DATE" | "TIME" => Some((3, 3)),
        "YEAR" | "MONTH" | "DAY" | "HOUR" | "MINUTE" | "SECOND" | "ISOWEEKNUM" | "DATEVALUE" | "TIMEVALUE" | "TO_DATE" => Some((1, 1)),
        "WEEKDAY" | "WEEKNUM" | "EPOCHTODATE" => Some((1, 2)),
        "EDATE" | "EOMONTH" | "DAYS" => Some((2, 2)),
        "NOW" | "TODAY" => Some((0, 0)),
        _ => None,
    };
    if arity.is_some_and(|(minimum, maximum)| arguments.len() < minimum || arguments.len() > maximum) { return "#VALUE!".into(); }
    match name {
        "DATE" => date(arguments, source, native, prefix),
        "YEAR" | "MONTH" | "DAY" => field(&arguments[0], name, source, native, prefix),
        "WEEKDAY" => weekday(arguments, source, prefix),
        "WEEKNUM" => weeknum(arguments, source, native, prefix),
        "ISOWEEKNUM" => iso_week(&arguments[0], source, native, prefix),
        "EDATE" | "EOMONTH" => month_end(arguments, name == "EOMONTH", source, native, prefix),
        "HOUR" | "MINUTE" | "SECOND" => clock(&arguments[0], name, prefix),
        "NOW" | "TODAY" | "DATEVALUE" => convert(&call(name, arguments), native, source, prefix),
        "TO_DATE" => arguments[0].clone(),
        "EPOCHTODATE" => {
            let value = format!("{prefix}v"); let unit = format!("{prefix}u");
            bind(&[(&value, arguments[0].clone()), (&unit, format!("INT({})", arguments.get(1).map(String::as_str).unwrap_or("1")))],
                &format!("IF(OR(NOT(ISNUMBER({value})),{value}<0,{unit}<1,{unit}>3),#NUM!,{}+TRUNC({value}*IF({unit}=1,1000,IF({unit}=2,1,0.001)))/86400000)", if source { 24_107 } else { 25_569 }))
        }
        _ => {
            if let Some((indices, date_result)) = native_date_arguments(name) {
                let mapped = arguments.iter().enumerate().map(|(index, argument)| if indices.contains(&index) {
                    convert(argument, source, native, &format!("{prefix}arg{index}_"))
                } else { argument.clone() }).collect::<Vec<_>>();
                let result = call(authored_name, &mapped);
                if date_result { convert(&result, native, source, &format!("{prefix}result_")) } else { result }
            } else { call(authored_name, arguments) }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn conversions_are_piecewise_and_capture_the_expression_once() {
        let forward = convert("NOW()", false, true, "_test_");
        assert_eq!(forward.matches("NOW()").count(), 1);
        assert!(forward.contains("<60,1461,1462"));
        assert!(forward.contains("ISNUMBER"));
        assert!(convert("A1", true, false, "_test_").contains("<-1401,1461,1462"));
        assert_eq!(convert("A1", true, true, "_test_"), "A1");
    }
    #[test]
    fn date_uses_the_normalized_month_origin_before_adding_the_day() {
        let expression = project("DATE", "DATE", &["1900".into(), "1".into(), "60".into()], false, true, "_test_");
        assert!(expression.contains("DATE(_test_ny+_test_shift,_test_nm,1)"));
        assert!(expression.contains("+_test_d-1"));
        assert!(expression.contains("146097"));
        assert_eq!(project("DATE", "DATE", &["1900".into()], false, true, "_test_"), "#VALUE!");
    }
    #[test]
    fn fields_keep_both_pseudo_dates_and_use_a_positive_native_calendar() {
        let day = project("DAY", "DAY", &["A1".into()], false, true, "_test_");
        assert!(day.contains("IF(_test_i=0,0,IF(_test_i=60,29,"));
        assert!(day.contains("146097"));
        let year = project("YEAR", "YEAR", &["A1".into()], false, true, "_test_");
        assert!(year.contains("IF(_test_i=0,1900,"));
        assert!(year.contains("-IF("));
    }
    #[test]
    fn weekdays_use_the_source_serial_axis_and_preserve_one_argument_evaluation() {
        let expression = project("WEEKDAY", "WEEKDAY", &["RAND()".into(), "2".into()], false, true, "_test_");
        assert_eq!(expression.matches("RAND()").count(), 1);
        assert!(expression.contains("MOD(INT(_test_s)+6,7)"));
        assert!(project("WEEKDAY", "WEEKDAY", &["A1".into()], true, false, "_test_").contains("MOD(INT(_test_s)+5,7)"));
    }
    #[test]
    fn financial_dates_convert_only_the_date_arguments_and_result() {
        let expression = project("COUPNCD", "COUPNCD", &["A1".into(), "B1".into(), "2".into(), "0".into()], false, true, "_test_");
        assert!(expression.contains("COUPNCD(_xlfn.LET"));
        assert!(expression.contains(",2,0)"));
        assert!(expression.contains("<-1401,1461,1462"));
        assert_eq!(project("SUM", "SUM", &["A1:A9".into()], false, true, "_test_"), "SUM(A1:A9)");
    }
}
