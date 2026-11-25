//! Literal/display positions and zero-length note ranks share one run scan.
use office_io_common::document_note_source::NOTE_SAFE_INTEGER;
use serde_json::Value;

use super::super::{get, note_reference};
use crate::ExportError;

struct Atom { start: usize, end: usize }

pub(in super::super) type NoteBoundary = (usize, usize);

#[derive(Clone, Copy)]
pub(in super::super) enum EventScope {
    Marker,
    ZeroRun,
    Field(Option<(NoteBoundary, NoteBoundary)>),
}

pub(in super::super) fn optional_reference_count(value: &Value, name: &str) -> Result<Option<usize>, ExportError> {
    value.get(name).map(|value| value.as_u64().filter(|value| *value <= NOTE_SAFE_INTEGER)
        .and_then(|value| usize::try_from(value).ok())
        .ok_or_else(|| ExportError::invalid(format!("{name} must be a nonnegative safe integer.")))).transpose()
}

pub(in super::super) struct ZeroRun {
    pub run_index: usize,
    pub offset: usize,
    pub rank: usize,
}

pub(in super::super) struct NotePositions {
    atoms: Vec<Atom>,
    length: usize,
    pub zero_runs: Vec<ZeroRun>,
}

impl NotePositions {
    pub fn from_runs(runs: &[Value]) -> Result<Self, ExportError> {
        let mut result = Self { atoms: Vec::new(), length: 0, zero_runs: Vec::new() };
        for (run_index, run) in runs.iter().enumerate() {
            let start = result.length;
            result.length = start.checked_add(get(run, "text").as_str().unwrap_or_default().encode_utf16().count())
                .filter(|length| *length as u64 <= NOTE_SAFE_INTEGER)
                .ok_or_else(|| ExportError::invalid("正文 UTF-16 长度超出安全范围。"))?;
            if start == result.length { result.zero_runs.push(ZeroRun { run_index, offset: start, rank: result.atoms.len() }); }
            if note_reference(run).is_some() {
                result.atoms.push(Atom { start, end: result.length });
            }
        }
        Ok(result)
    }

    pub fn length(&self) -> usize { self.length }
    pub fn reference_count(&self) -> usize { self.atoms.len() }

    /** Explicit ranks address the individual zero-length objects at a boundary.
     * An unranked endpoint belongs before those following objects. */
    pub fn boundary_rank(&self, offset: usize, explicit: Option<usize>) -> Result<usize, ExportError> {
        if offset > self.length { return Err(ExportError::invalid("批注或控件端点超出段落。")); }
        let before = self.atoms.partition_point(|atom| atom.end < offset || atom.end == offset && atom.start < atom.end);
        if self.atoms.get(before).is_some_and(|atom| atom.start < offset && offset < atom.end) {
            return Err(ExportError::invalid("批注或控件端点不能位于注释引用标记内部。"));
        }
        let tied = self.atoms[before..].partition_point(|atom| atom.start == offset && atom.end == offset);
        let rank = explicit.unwrap_or(before);
        if rank < before || rank > before + tied { return Err(ExportError::invalid("批注端点的注释引用序号与当前正文位置不一致。")); }
        Ok(rank)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;

    // Unrun contract drafts for the unified verification stage.
    #[test]
    fn adjacent_empty_references_keep_before_between_and_after_boundaries() {
        let runs = vec![
            json!({ "text": "😀" }),
            json!({ "text": "", "footnoteId": "first" }),
            json!({ "text": "", "endnoteId": "second" }),
            json!({ "text": "X" }),
        ];
        let positions = NotePositions::from_runs(&runs).unwrap();
        assert_eq!(positions.zero_runs.iter().map(|reference| (reference.run_index, reference.offset, reference.rank)).collect::<Vec<_>>(),
            vec![(1, 2, 0), (2, 2, 1)]);
        for rank in 0..=2 { assert_eq!(positions.boundary_rank(2, Some(rank)).unwrap(), rank); }
        assert_eq!(positions.boundary_rank(2, None).unwrap(), 0);
        assert_eq!(positions.boundary_rank(3, None).unwrap(), 2);
        assert!(positions.boundary_rank(2, Some(3)).is_err());
        assert!(positions.boundary_rank(3, Some(0)).is_err());
    }

    #[test]
    fn multi_digit_and_custom_reference_labels_are_atomic() {
        let positions = NotePositions::from_runs(&[
            json!({ "text": "A" }), json!({ "text": "12", "footnoteId": "number" }),
            json!({ "text": "😀*", "endnoteId": "custom" }),
        ]).unwrap();
        assert_eq!(positions.boundary_rank(1, Some(0)).unwrap(), 0);
        assert_eq!(positions.boundary_rank(3, Some(1)).unwrap(), 1);
        assert_eq!(positions.boundary_rank(6, Some(2)).unwrap(), 2);
        for offset in [2, 4, 5] { assert!(positions.boundary_rank(offset, None).is_err()); }
        assert!(positions.boundary_rank(1, Some(1)).is_err());
    }
}
