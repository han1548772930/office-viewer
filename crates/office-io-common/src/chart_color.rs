//! CSS colors at the authored-model boundary, converted to native sRGB.

pub fn rgba(value: &str) -> Result<([u8; 3], f64), String> {
    let color = csscolorparser::parse(value).map_err(|error| format!("invalid chart color {value}: {error}"))?;
    let [red, green, blue, _] = color.to_rgba8();
    let alpha = f64::from(color.to_array()[3]);
    Ok(([red, green, blue], alpha))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn css_formats_preserve_rgb_and_fractional_alpha_for_drawingml() {
        for value in ["red", "#f00", "#ff0000", "rgb(255, 0, 0)", "hsl(0, 100%, 50%)"] {
            assert_eq!(rgba(value).unwrap(), ([255, 0, 0], 1.0), "{value}");
        }
        let (rgb, alpha) = rgba("rgba(0, 128, 255, 0.37)").unwrap();
        assert_eq!(rgb, [0, 128, 255]);
        assert!((alpha - 0.37).abs() < 0.00001);
        assert!(rgba("not-a-color").is_err());
    }
}
