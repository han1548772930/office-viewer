//! Read only color-filter referenced dxfs; reuse OPC theme/palette resources.
use std::{collections::{BTreeMap, BTreeSet}, io::{BufReader, Cursor}};
use office_io_common::{OfficeFilterNode, filter_color::{self, FilterColor}, conditional_format_xml::{self as xml, Element, XmlSelectionLimits}};
use zip::ZipArchive;
use crate::{xlsx_color_resources::{self, ColorResources}, xlsx_parts};

#[derive(Debug, Default)]
pub(crate) struct FilterColors {
    dxfs: BTreeMap<usize, Element>,
    colors: ColorResources,
    pub(crate) error: Option<String>,
}

pub(crate) fn referenced_ids<'a>(roots: impl Iterator<Item = &'a OfficeFilterNode>) -> Result<BTreeSet<usize>, String> {
    fn visit(node: &OfficeFilterNode, ids: &mut BTreeSet<usize>) -> Result<(), String> {
        if node.is_standard("colorFilter") { if let Some(id) = filter_color::dxf_index(node)? { ids.insert(id); } }
        if ids.len() > 4096 { return Err("AUTOFILTER_COLOR_LIMIT: more than 4096 referenced differential styles".into()); }
        for child in &node.children { visit(child, ids)?; } Ok(())
    }
    let mut ids = BTreeSet::new(); for root in roots { visit(root, &mut ids)?; } Ok(ids)
}

impl FilterColors {
    pub(crate) fn failed(error: String) -> Self { Self { error: Some(error), ..Self::default() } }
    pub(crate) fn read(archive: &mut ZipArchive<Cursor<&[u8]>>, ids: &BTreeSet<usize>) -> Result<Self, String> {
        if ids.is_empty() { return Ok(Self::default()); }
        let colors = xlsx_color_resources::read_color_resources(archive)?;
        let book = xlsx_parts::workbook_parts(archive)?;
        let rels = xlsx_parts::relationships(&xlsx_parts::required_xml_part(archive, &xlsx_parts::relationship_part_path(&book.workbook_path))?)?;
        let mut styles = rels.iter().filter(|rel| rel.kind == "styles");
        let style = styles.next().filter(|rel| !rel.external).ok_or("AUTOFILTER_COLOR_REQUIRED_RESOURCE: missing internal styles relationship")?;
        if styles.next().is_some() { return Err("AUTOFILTER_COLOR_REQUIRED_RESOURCE: multiple styles relationships".into()); }
        let path = xlsx_parts::resolve_target(&book.workbook_path, &style.target).ok_or("AUTOFILTER_COLOR_REQUIRED_RESOURCE: invalid styles target")?;
        let entry = archive.by_name(&path).map_err(|error| error.to_string())?;
        let mut index = 0usize; let mut selected_ids = Vec::new();
        let selected = xml::extract(BufReader::new(entry), XmlSelectionLimits::default(), |path, _| {
            if path.len() != 3 || !path[0].is_main("styleSheet") || !path[1].is_main("dxfs") || !path[2].is_main("dxf") { return Ok(false); }
            let current = index; index = index.checked_add(1).ok_or("AUTOFILTER_COLOR_LIMIT: differential-style index overflow")?;
            if ids.contains(&current) { selected_ids.push(current); Ok(true) } else { Ok(false) }
        })?;
        if !selected.root.is_main("styleSheet") { return Err("AUTOFILTER_COLOR_REQUIRED_RESOURCE: invalid stylesheet root".into()); }
        let dxfs = selected_ids.into_iter().zip(selected.fragments).collect::<BTreeMap<_, _>>();
        if ids.iter().any(|id| !dxfs.contains_key(id)) { return Err("AUTOFILTER_COLOR_REQUIRED_RESOURCE: referenced dxfId is missing".into()); }
        Ok(Self { dxfs, colors, error: None })
    }

    pub(crate) fn decode(&self, node: &OfficeFilterNode) -> Result<FilterColor, String> {
        if let Some(error) = &self.error { return Err(error.clone()); }
        let dxf = filter_color::dxf_index(node)?.and_then(|index| self.dxfs.get(&index));
        filter_color::decode(node, dxf, self.colors.theme.as_ref(), self.colors.indexed_colors.as_deref())
    }
    pub(crate) fn signature(&self, node: &OfficeFilterNode) -> Result<Option<String>, String> { Ok(Some(self.decode(node)?.signature())) }
}
