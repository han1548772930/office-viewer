//! Native table styles and theme references share the existing source registry.
use super::{BTreeMap, BTreeSet, Content, Element, Source, SourceStyles, STYLE_DEPTH_LIMIT, STYLE_LIMIT,
    WORD_NAMESPACE, style_chain, style_closure, theme_requirements, validate_resource};

fn references<'a>(root: &'a Element, output: &mut Vec<(&'a str, &'static str)>) -> Result<(), String> {
    if root.namespace == WORD_NAMESPACE {
        let kind = match root.local.as_str() { "tblStyle" => Some("table"), "pStyle" => Some("paragraph"), "rStyle" => Some("character"), _ => None };
        if let Some(kind) = kind {
            let id = root.attribute("val").filter(|id| !id.trim().is_empty()).ok_or("Native table style reference requires a nonempty ID")?;
            output.push((id, kind));
        }
    }
    for child in root.children() { references(child, output)?; }
    Ok(())
}

fn table_closure(source: &Source, id: &str, output: &mut BTreeMap<String, Element>, active: &mut BTreeSet<String>, depth: usize) -> Result<(), String> {
    if depth >= STYLE_DEPTH_LIMIT { return Err("Table source style dependencies exceed their depth budget".into()); }
    if !active.insert(id.into()) { return Err("Native table styles contain a circular resource reference".into()); }
    // style_chain validates both the expected type and basedOn cycles before
    // import; native conditional formatting remains inside the table style.
    for style in style_chain(source, id, "table", &mut BTreeSet::new())? {
        let style_id = style.attribute("styleId").ok_or("Native table style requires a styleId")?.to_owned();
        if output.contains_key(&style_id) { continue; }
        if output.len() >= STYLE_LIMIT { return Err("Table source styles exceed their resource count budget".into()); }
        validate_resource(&style, WORD_NAMESPACE)?;
        let mut nested = Vec::new();
        references(&style, &mut nested)?;
        let nested = nested.into_iter().map(|(id, kind)| (id.to_owned(), kind)).collect::<Vec<_>>();
        // next/link can refer to text styles. They use the same registry and
        // are remapped in styles.xml together with nested pStyle/rStyle.
        let linked = ["next", "link"].into_iter().filter_map(|name| style.child(name).and_then(|node| node.attribute("val")).map(str::to_owned)).collect::<Vec<_>>();
        output.insert(style_id, style);
        for (id, kind) in nested {
            if kind == "table" { table_closure(source, &id, output, active, depth + 1)?; }
            else {
                style_chain(source, &id, kind, &mut BTreeSet::new())?;
                style_closure(source, &id, output, depth + 1)?;
            }
        }
        for id in linked { style_closure(source, &id, output, depth + 1)?; }
    }
    active.remove(id);
    Ok(())
}

pub(super) fn remap_references(root: &mut Element, mapping: &BTreeMap<String, String>) {
    if root.namespace == WORD_NAMESPACE && matches!(root.local.as_str(), "tblStyle" | "pStyle" | "rStyle" | "basedOn" | "next" | "link" | "styleLink" | "numStyleLink") {
        if let Some(mapped) = root.attribute("val").and_then(|id| mapping.get(id)).cloned() { root.set_word_attribute("val", mapped); }
    }
    for content in &mut root.content { if let Content::Element(child) = content { remap_references(child, mapping); } }
}

impl SourceStyles {
    pub(in super::super) fn restore_table_properties(&mut self, root: &mut Element) -> Result<(), String> {
        let mut needed = Vec::new();
        references(root, &mut needed)?;
        let required_theme = theme_requirements(root);
        if needed.is_empty() && required_theme == (false, false) && !super::numbering_source::has_live_numbering(root) { return Ok(()); }
        let source = self.source.as_ref().ok_or_else(|| self.source_error.clone().unwrap_or_else(|| "Native table styles/theme refer to source resources that were not retained".into()))?;
        let mut closure = BTreeMap::new();
        for (id, kind) in needed {
            if kind == "table" { table_closure(source, id, &mut closure, &mut BTreeSet::new(), 0)?; }
            else {
                style_chain(source, id, kind, &mut BTreeSet::new())?;
                style_closure(source, id, &mut closure, 0)?;
            }
        }
        self.import_dependencies(&closure, required_theme, Some(root))?;
        remap_references(root, &self.remap);
        self.numbering_resources.remap(root)?;
        Ok(())
    }
}
