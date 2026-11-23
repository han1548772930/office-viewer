//! The Office IO date boundary shares Sheets' explicit 1900/1904 serial rules.
//! No host time zone, locale date parser or native Date normalization is used.

use serde_json::Value;

pub const DATE_1904_OFFSET: f64 = 1462.0;
pub const MAX_DATE_1900_SERIAL: f64 = 2_958_465.0;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DateZone { None, Utc, Offset }

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct IsoDateLayout {
    pub time: bool,
    pub seconds: bool,
    pub fractional_digits: usize,
    pub zone: DateZone,
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ChartDate {
    pub serial: f64,
    /// None denotes a numeric serial, including strictly numeric text.
    pub iso: Option<IsoDateLayout>,
}

pub fn max_date_serial(date1904: bool) -> f64 {
    MAX_DATE_1900_SERIAL - if date1904 { DATE_1904_OFFSET } else { 0.0 }
}

pub fn valid_date_serial(value: f64, date1904: bool) -> bool {
    value.is_finite() && value >= 0.0 && value < max_date_serial(date1904) + 1.0
}

pub fn chart_date(value: &Value, date1904: bool) -> Option<ChartDate> {
    if let Some(serial) = value.as_f64() {
        return valid_date_serial(serial, date1904).then_some(ChartDate { serial, iso: None });
    }
    let text = value.as_str()?.trim();
    if numeric_text(text) {
        let serial = text.parse::<f64>().ok()?;
        return valid_date_serial(serial, date1904).then_some(ChartDate { serial, iso: None });
    }
    iso_date(text, date1904)
}

fn numeric_text(text: &str) -> bool {
    let bytes = text.as_bytes();
    let mut index = usize::from(bytes.first() == Some(&b'+'));
    let digits = |index: &mut usize| {
        let start = *index;
        while bytes.get(*index).is_some_and(u8::is_ascii_digit) { *index += 1; }
        *index - start
    };
    let mut count = digits(&mut index);
    if bytes.get(index) == Some(&b'.') { index += 1; count += digits(&mut index); }
    if count == 0 { return false; }
    if matches!(bytes.get(index), Some(b'e' | b'E')) {
        index += 1;
        if matches!(bytes.get(index), Some(b'+' | b'-')) { index += 1; }
        if digits(&mut index) == 0 { return false; }
    }
    index == bytes.len()
}

fn digits(text: &[u8]) -> Option<u32> {
    if text.is_empty() || !text.iter().all(u8::is_ascii_digit) { return None; }
    Some(text.iter().fold(0, |number, byte| number * 10 + u32::from(*byte - b'0')))
}

fn iso_date(text: &str, date1904: bool) -> Option<ChartDate> {
    let bytes = text.as_bytes();
    if bytes.len() < 10 || bytes[4] != b'-' || bytes[7] != b'-' { return None; }
    let year = digits(&bytes[..4])?;
    let month = digits(&bytes[5..7])?;
    let day = digits(&bytes[8..10])?;
    let mut index = 10;
    let mut layout = IsoDateLayout { time: false, seconds: false, fractional_digits: 0, zone: DateZone::None };
    let (mut hours, mut minutes, mut seconds, mut fraction) = (0, 0, 0, 0.0);
    if bytes.len() > index {
        if !matches!(bytes[index], b'T' | b' ') || bytes.len() < 16 || bytes[13] != b':' { return None; }
        layout.time = true;
        hours = digits(&bytes[11..13])?; minutes = digits(&bytes[14..16])?;
        index = 16;
        if bytes.get(index) == Some(&b':') {
            seconds = digits(bytes.get(index + 1..index + 3)?)?;
            index += 3; layout.seconds = true;
            if bytes.get(index) == Some(&b'.') {
                index += 1; let start = index;
                while bytes.get(index).is_some_and(u8::is_ascii_digit) { index += 1; }
                layout.fractional_digits = index - start;
                if !(1..=9).contains(&layout.fractional_digits) { return None; }
                fraction = f64::from(digits(&bytes[start..index])?) / 10f64.powi(layout.fractional_digits as i32);
            }
        }
    }
    if hours > 23 || minutes > 59 || seconds > 59 { return None; }
    let mut zone_minutes = 0.0;
    match bytes.get(index) {
        None => {},
        Some(b'Z') if layout.time && index + 1 == bytes.len() => { layout.zone = DateZone::Utc; index += 1; },
        Some(b'+' | b'-') if layout.time && index + 6 == bytes.len() && bytes[index + 3] == b':' => {
            let zone_hours = digits(&bytes[index + 1..index + 3])?;
            let zone_minute = digits(&bytes[index + 4..index + 6])?;
            if zone_hours > 23 || zone_minute > 59 { return None; }
            zone_minutes = f64::from(zone_hours * 60 + zone_minute) * if bytes[index] == b'-' { -1.0 } else { 1.0 };
            layout.zone = DateZone::Offset; index += 6;
        },
        _ => return None,
    }
    if index != bytes.len() { return None; }
    let base = date_serial(year, month, day, date1904)?;
    let serial = base + (f64::from(hours * 3600 + minutes * 60 + seconds) + fraction) / 86_400.0 - zone_minutes / 1440.0;
    valid_date_serial(serial, date1904).then_some(ChartDate { serial, iso: Some(layout) })
}

fn date_serial(year: u32, month: u32, day: u32, date1904: bool) -> Option<f64> {
    let minimum_year = if date1904 { 1904 } else { 1900 };
    if year < minimum_year || year > 9999 || !(1..=12).contains(&month) { return None; }
    if !date1904 && year == 1900 {
        if month == 1 && day == 0 { return Some(0.0); }
        if month == 2 && day == 29 { return Some(60.0); }
    }
    let leap = year % 4 == 0 && (year % 100 != 0 || year % 400 == 0);
    let months = [31, if leap { 29 } else { 28 }, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
    if day == 0 || day > months[(month - 1) as usize] { return None; }
    let before_year = |year: u32| { let previous = year - 1; previous * 365 + previous / 4 - previous / 100 + previous / 400 };
    let calendar_days = before_year(year) - before_year(1900) + months[..(month - 1) as usize].iter().sum::<u32>() + day;
    let serial = f64::from(calendar_days + u32::from(year > 1900 || month > 2));
    Some(serial - if date1904 { DATE_1904_OFFSET } else { 0.0 })
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;

    #[test]
    fn the_shared_io_and_workbench_fixture_has_one_date_contract() {
        let cases: Vec<Value> = serde_json::from_str(include_str!("../../../tests/fixtures/charts/date-serials.json")).unwrap();
        for case in cases {
            let actual = chart_date(&case["value"], case["date1904"].as_bool().unwrap()).map(|date| date.serial);
            assert_eq!(actual, case["serial"].as_f64(), "{case}");
        }
    }

    #[test]
    fn epochs_pseudo_dates_and_the_last_calendar_day_match_sheets() {
        for (text, serial) in [("1900-01-00", 0.0), ("1900-01-01", 1.0), ("1900-02-28", 59.0), ("1900-02-29", 60.0), ("1900-03-01", 61.0), ("1904-01-01", 1462.0), ("9999-12-31", MAX_DATE_1900_SERIAL)] {
            assert_eq!(chart_date(&json!(text), false).unwrap().serial, serial, "{text}");
        }
        assert_eq!(chart_date(&json!("1904-01-01"), true).unwrap().serial, 0.0);
        assert_eq!(chart_date(&json!("9999-12-31"), true).unwrap().serial, max_date_serial(true));
        assert!(chart_date(&json!("1900-02-29"), true).is_none());
        assert!(chart_date(&json!("1903-12-31"), true).is_none());
    }

    #[test]
    fn strict_iso_offsets_fractions_and_numeric_text_do_not_use_locale_parsing() {
        let parsed = chart_date(&json!("1904-01-01T08:30:00.125+08:30"), true).unwrap();
        assert!((parsed.serial - 0.125 / 86_400.0).abs() < 1e-12);
        assert_eq!(parsed.iso.unwrap().fractional_digits, 3);
        assert_eq!(chart_date(&json!(" +6.0e1 "), false).unwrap().serial, 60.0);
        assert_eq!(chart_date(&json!("1900-02-29T12:00Z"), false).unwrap().serial, 60.5);
        for text in ["2026-02-30", "2026-11-24Z", "2026-11-24T24:00", "2026-11-24T12:00+24:00", "2026-11-24T12:00:00.", "11/24/2026", "0x10", "-1", "+", "NaN"] {
            assert!(chart_date(&json!(text), false).is_none(), "{text}");
        }
        assert!(chart_date(&json!(true), false).is_none());
        assert!(chart_date(&json!(MAX_DATE_1900_SERIAL + 1.0), false).is_none());
    }
}
