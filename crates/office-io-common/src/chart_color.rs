//! CSS colors at the authored-model boundary, converted to native sRGB.

use serde::{Deserialize, Serialize};

pub const MAX_GRADIENT_STOPS: usize = 256;

/// The serializable chart-color contract. Renderer-specific paint objects do
/// not cross the Office DTO boundary.
#[derive(Debug, Clone, PartialEq, Deserialize, Serialize)]
#[serde(untagged)]
pub enum ChartColor {
    Solid(String),
    LinearGradient(ChartLinearGradient),
}

#[derive(Debug, Clone, Copy, PartialEq, Deserialize, Serialize)]
pub enum GradientKind {
    #[serde(rename = "linear-gradient")]
    Linear,
}

#[derive(Debug, Clone, Copy, PartialEq, Deserialize, Serialize)]
#[serde(deny_unknown_fields)]
pub struct GradientPoint { pub x: f64, pub y: f64 }

#[derive(Debug, Clone, PartialEq, Deserialize, Serialize)]
#[serde(deny_unknown_fields)]
pub struct GradientStop {
    pub offset: f64,
    pub color: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub opacity: Option<f64>,
}

#[derive(Debug, Clone, PartialEq, Deserialize, Serialize)]
#[serde(deny_unknown_fields)]
pub struct ChartLinearGradient {
    #[serde(rename = "type")]
    pub kind: GradientKind,
    pub start: GradientPoint,
    pub end: GradientPoint,
    pub stops: Vec<GradientStop>,
}

impl From<String> for ChartColor { fn from(value: String) -> Self { Self::Solid(value) } }
impl From<&str> for ChartColor { fn from(value: &str) -> Self { Self::Solid(value.into()) } }

impl ChartColor {
    pub fn as_solid(&self) -> Option<&str> { match self { Self::Solid(value) => Some(value), Self::LinearGradient(_) => None } }

    pub fn validate(&self) -> Result<(), String> {
        match self { Self::Solid(value) => rgba(value).map(|_| ()), Self::LinearGradient(value) => value.validate() }
    }

    /// Legends and inherited plain marker/border colors use the first stop by
    /// offset, matching the editor. This does not replace the gradient fill.
    pub fn representative(&self) -> Result<String, String> {
        match self {
            Self::Solid(value) => { rgba(value)?; Ok(value.clone()) }
            Self::LinearGradient(value) => {
                value.validate()?;
                let stop = value.stops.iter().min_by(|left, right| left.offset.total_cmp(&right.offset)).ok_or("chart gradient is missing its stops")?;
                let (rgb, alpha) = rgba(&stop.color)?;
                Ok(css_rgba(rgb, alpha * stop.opacity.unwrap_or(1.0)))
            }
        }
    }
}

impl ChartLinearGradient {
    pub fn validate(&self) -> Result<(), String> {
        if ![self.start.x, self.start.y, self.end.x, self.end.y].iter().all(|value| value.is_finite()) {
            return Err("chart gradient coordinates must be finite".into());
        }
        if !(2..=MAX_GRADIENT_STOPS).contains(&self.stops.len()) {
            return Err(format!("chart gradient requires 2..{MAX_GRADIENT_STOPS} stops"));
        }
        for stop in &self.stops {
            if !stop.offset.is_finite() || !(0.0..=1.0).contains(&stop.offset)
                || stop.opacity.is_some_and(|opacity| !opacity.is_finite() || !(0.0..=1.0).contains(&opacity)) {
                return Err("chart gradient stop offsets and opacity must be between 0 and 1".into());
            }
            rgba(&stop.color)?;
        }
        Ok(())
    }
}

pub fn css_rgba(rgb: [u8; 3], alpha: f64) -> String {
    if alpha == 1.0 { format!("#{:02x}{:02x}{:02x}", rgb[0], rgb[1], rgb[2]) }
    else { format!("rgba({}, {}, {}, {})", rgb[0], rgb[1], rgb[2], alpha) }
}

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

    #[test]
    fn typed_gradients_preserve_author_order_and_renderer_objects_are_rejected() {
        let raw = serde_json::json!({"type":"linear-gradient","start":{"x":0.0,"y":0.0},"end":{"x":1.0,"y":0.0},
            "stops":[{"offset":1.0,"color":"blue"},{"offset":0.0,"color":"red","opacity":0.5}]});
        let color: ChartColor = serde_json::from_value(raw.clone()).unwrap();
        color.validate().unwrap();
        assert_eq!(serde_json::to_value(&color).unwrap(), raw);
        assert_eq!(color.as_solid(), None);
        assert_eq!(color.representative().unwrap(), "rgba(255, 0, 0, 0.5)");
        let solid: ChartColor = serde_json::from_value(serde_json::json!("#123456")).unwrap();
        assert_eq!(solid.as_solid(), Some("#123456"));
        for raw in [serde_json::json!({"type":"linear","x":0,"colorStops":[]}),
            serde_json::json!({"type":"radial-gradient","start":{"x":0,"y":0},"end":{"x":1,"y":0},"stops":[]})] {
            assert!(serde_json::from_value::<ChartColor>(raw).is_err());
        }
    }

    #[test]
    fn stop_limits_finite_coordinates_and_unknown_fields_are_enforced() {
        let mut gradient = ChartLinearGradient { kind:GradientKind::Linear, start:GradientPoint { x:0.0,y:0.0 }, end:GradientPoint { x:1.0,y:0.0 },
            stops:vec![GradientStop { offset:0.0,color:"red".into(),opacity:None },GradientStop { offset:1.0,color:"blue".into(),opacity:None }] };
        gradient.validate().unwrap();
        gradient.stops.resize(MAX_GRADIENT_STOPS,gradient.stops[0].clone());
        gradient.validate().unwrap();
        gradient.stops.push(gradient.stops[0].clone());
        assert!(gradient.validate().is_err());
        gradient.stops.truncate(2); gradient.start.x = f64::INFINITY;
        assert!(gradient.validate().is_err());
        gradient.start.x = 0.0; gradient.stops[0].offset = -0.1;
        assert!(gradient.validate().is_err());
        assert!(serde_json::from_value::<GradientPoint>(serde_json::json!({"x":0,"y":0,"native":true})).is_err());
        assert!(serde_json::from_value::<GradientStop>(serde_json::json!({"offset":0,"color":"red","native":true})).is_err());
    }
}
