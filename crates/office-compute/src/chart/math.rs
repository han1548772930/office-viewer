use crate::ComputeError;

pub(super) fn finite(value: f64, label: &str) -> Result<f64, ComputeError> {
    if value.is_finite() {
        Ok(value)
    } else {
        Err(ComputeError::new("COMPUTE_NUMERIC_RANGE", format!("{label} exceeds finite binary64")))
    }
}

pub(super) fn interpolate(first: f64, last: f64, ratio: f64) -> Result<f64, ComputeError> {
    if ratio <= 0.0 { return Ok(first); }
    if ratio >= 1.0 { return Ok(last); }
    let difference = last - first;
    finite(if difference.is_finite() {
        first + difference * ratio
    } else {
        first * (1.0 - ratio) + last * ratio
    }, "Chart interpolation")
}

#[derive(Default, Clone, Copy)]
pub(super) struct CompensatedSum {
    total: f64,
    compensation: f64,
}

impl CompensatedSum {
    pub fn add(&mut self, value: f64) {
        let addend = value - self.compensation;
        let next = self.total + addend;
        self.compensation = next - self.total - addend;
        self.total = next;
    }

    pub fn value(self) -> f64 { self.total }
}

#[derive(Clone, Copy)]
pub(super) struct Moments {
    pub mean: f64,
    pub scale: f64,
    pub normalized_mean: f64,
    pub normalized_variance: f64,
}

/// Iterator clones rescan borrowed input; no full-sized sample Vec is cloned.
pub(super) fn moments<I: Iterator<Item = f64> + Clone>(samples: I) -> Result<Moments, ComputeError> {
    let mut scale: f64 = 0.0;
    let mut count = 0;
    for value in samples.clone() {
        scale = scale.max(finite(value, "Chart sample")?.abs());
        count += 1;
    }
    if count == 0 || scale == 0.0 {
        return Ok(Moments { mean: 0.0, scale: 1.0, normalized_mean: 0.0,
            normalized_variance: 0.0 });
    }
    let mut total = CompensatedSum::default();
    for value in samples.clone() { total.add(value / scale); }
    let normalized_mean = (total.value() / count as f64).clamp(-1.0, 1.0);
    total = CompensatedSum::default();
    for value in samples { total.add((value / scale - normalized_mean).powi(2)); }
    Ok(Moments { mean: finite(normalized_mean * scale, "Chart mean")?, scale,
        normalized_mean, normalized_variance: (total.value() / count as f64).clamp(0.0, 1.0) })
}

/// Combine factors in binary mantissa/exponent form. A finite final product is
/// retained even when a left-to-right intermediate would overflow/underflow.
pub(super) fn product(factors: &[f64], label: &str) -> Result<f64, ComputeError> {
    let mut negative = false;
    let mut zero = false;
    let mut mantissa = 1.0;
    let mut exponent: i32 = 0;
    for &factor in factors {
        finite(factor, label)?;
        if factor == 0.0 { zero = true; continue; }
        negative ^= factor.is_sign_negative();
        let mut absolute = factor.abs();
        let mut shift = 0;
        if absolute < f64::MIN_POSITIVE { absolute *= 2.0_f64.powi(52); shift = -52; }
        let bits = absolute.to_bits();
        let power = ((bits >> 52) & 0x7ff) as i32 - 1023 + shift;
        let part = f64::from_bits((1023_u64 << 52) | (bits & ((1_u64 << 52) - 1)));
        mantissa *= part;
        exponent += power;
        if mantissa >= 2.0 { mantissa *= 0.5; exponent += 1; }
    }
    if zero { return Ok(0.0); }
    if exponent > 1023 { return finite(f64::INFINITY, label); }
    let value = if exponent < -1022 {
        mantissa * f64::MIN_POSITIVE * 2.0_f64.powi(exponent + 1022)
    } else {
        mantissa * 2.0_f64.powi(exponent)
    };
    finite(if negative { -value } else { value }, label)
}

pub(super) fn reserve<T>(values: &mut Vec<T>, count: usize) -> Result<(), ComputeError> {
    values.try_reserve_exact(count).map_err(|_| ComputeError::memory())
}
