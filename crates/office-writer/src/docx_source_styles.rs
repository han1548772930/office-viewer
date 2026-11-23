use std::collections::{BTreeMap, BTreeSet, HashSet};

use serde_json::Value;

use super::{get, source_xml::{Attribute, Content, Element, DRAWING_NAMESPACE, WORD_NAMESPACE, XMLNS_NAMESPACE, XML_NAMESPACE, canonicalize_office_namespaces, parse_xml}};
use crate::{ExportError, xml};

#[path = "docx_source_story.rs"]
mod story;

#[cfg(test)]
#[path = "docx_source_story_tests.rs"]
mod story_tests;

const STYLE_LIMIT: usize = 4_096;
const STYLE_DEPTH_LIMIT: usize = 64;

struct Theme {
    xml: String,
    has_fonts: bool,
    has_colors: bool,
    has_minor_latin: bool,
}

struct Source {
    styles: BTreeMap<String, Element>,
    defaults: Vec<Element>,
    paragraph_defaults: Vec<Element>,
    default_paragraph: Option<String>,
    namespaces: BTreeMap<String, String>,
    theme: Result<Option<Theme>, String>,
}

pub(super) struct SourceStyles {
    generated: String,
    source: Option<Source>,
    source_error: Option<String>,
    imported: BTreeMap<String, Element>,
    remap: BTreeMap<String, String>,
    used_ids: HashSet<String>,
    next_id: u32,
    theme_used: bool,
}

#[derive(Clone, Default)]
struct Properties(Vec<Element>);

fn enabled(element: &Element) -> bool { !matches!(element.attribute("val"), Some("0" | "false" | "off")) }

fn toggle_property(local: &str) -> bool {
    matches!(local, "b" | "bCs" | "i" | "iCs" | "caps" | "smallCaps" | "strike" | "dstrike" | "outline" | "shadow" | "emboss" | "imprint" | "vanish")
}

impl Properties {
    fn merge(&mut self, properties: &[Element], style: bool) {
        for property in properties {
            if property.is_word("rStyle") { continue; }
            let existing = self.0.iter().position(|candidate| candidate.namespace == property.namespace && candidate.local == property.local);
            let mut next = property.clone();
            if style && property.namespace == WORD_NAMESPACE && toggle_property(&property.local) {
                if !enabled(property) { continue; }
                let prior = existing.is_some_and(|index| enabled(&self.0[index]));
                next.set_word_attribute("val", if prior { "0" } else { "1" });
            }
            if property.is_word("rFonts") || property.is_word("lang") {
                if let Some(index) = existing {
                    let mut merged = self.0[index].clone();
                    for attribute in &next.attributes {
                        if next.is_word("rFonts") && attribute.namespace == WORD_NAMESPACE {
                            let competing = match attribute.local.as_str() {
                                "ascii" => Some("asciiTheme"), "asciiTheme" => Some("ascii"),
                                "hAnsi" => Some("hAnsiTheme"), "hAnsiTheme" => Some("hAnsi"),
                                "eastAsia" => Some("eastAsiaTheme"), "eastAsiaTheme" => Some("eastAsia"),
                                "cs" => Some("cstheme"), "cstheme" => Some("cs"), _ => None,
                            };
                            if let Some(competing) = competing.filter(|name| next.attribute(name).is_none()) { merged.attributes.retain(|candidate| !(candidate.namespace == WORD_NAMESPACE && candidate.local == competing)); }
                        }
                        if let Some(prior) = merged.attributes.iter_mut().find(|prior| prior.namespace == attribute.namespace && prior.local == attribute.local) { *prior = attribute.clone(); }
                        else { merged.attributes.push(attribute.clone()); }
                    }
                    next = merged;
                }
            }
            if let Some(index) = existing { self.0[index] = next; }
            else { self.0.push(next); }
        }
    }
}

/// Properties are detached from their style/default containers when merging.
/// Retain the complete scope at each detached root so two source contexts can
/// use the same extension prefix without silently changing its binding.
fn scoped_children(element: &Element, inherited: &BTreeMap<String, String>) -> Vec<Element> {
    let mut scope = inherited.clone();
    scope.extend(namespaces(element));
    element.children().map(|child| {
        let mut child = child.clone();
        attach_namespaces(&mut child, &scope);
        child
    }).collect()
}

fn child_properties(element: &Element, local: &str) -> Vec<Element> {
    element.child(local).map(|properties| scoped_children(properties, &namespaces(element))).unwrap_or_default()
}

fn properties(style: &Element) -> Vec<Element> { child_properties(style, "rPr") }

fn namespaces(root: &Element) -> BTreeMap<String, String> {
    let mut result = root.attributes.iter().filter(|attribute| attribute.namespace == XMLNS_NAMESPACE)
        .map(|attribute| (attribute.name.strip_prefix("xmlns:").unwrap_or("").to_owned(), attribute.value.clone())).collect::<BTreeMap<_, _>>();
    result.insert("w".into(), WORD_NAMESPACE.into());
    result
}

fn validate_resource(element: &Element, primary: &str) -> Result<(), String> {
    let extension = |namespace: &str| namespace.starts_with("http://schemas.microsoft.com/office/word/") || namespace.starts_with("http://schemas.microsoft.com/office/drawing/");
    if element.namespace != primary && !extension(&element.namespace) { return Err(format!("Source resource contains unsupported namespace {:?}.", element.namespace)); }
    for attribute in &element.attributes {
        if !matches!(attribute.namespace.as_str(), "" | XML_NAMESPACE | XMLNS_NAMESPACE | "http://schemas.openxmlformats.org/markup-compatibility/2006") && attribute.namespace != primary && !extension(&attribute.namespace) {
            return Err("Source style or theme depends on a relationship or extension resource that was not retained.".into());
        }
    }
    for child in element.children() { validate_resource(child, primary)?; }
    Ok(())
}

fn read_theme(resource: &Value) -> Result<Option<Theme>, String> {
    if resource.is_null() { return Ok(None); }
    let xml = get(resource, "xml").as_str().ok_or("Source theme has no XML.")?;
    let mut root = parse_xml(xml, true)?;
    canonicalize_office_namespaces(&mut root)?;
    if root.namespace != DRAWING_NAMESPACE || root.local != "theme" { return Err("Source theme has an invalid root element.".into()); }
    validate_resource(&root, DRAWING_NAMESPACE)?;
    let elements = root.children().find(|child| child.namespace == DRAWING_NAMESPACE && child.local == "themeElements").ok_or("Source theme has no themeElements.")?;
    let fonts = elements.children().find(|child| child.namespace == DRAWING_NAMESPACE && child.local == "fontScheme");
    let has_minor_latin = fonts.and_then(|fonts| fonts.children().find(|child| child.namespace == DRAWING_NAMESPACE && child.local == "minorFont"))
        .and_then(|fonts| fonts.children().find(|child| child.namespace == DRAWING_NAMESPACE && child.local == "latin"))
        .is_some_and(|font| font.attributes.iter().any(|attribute| attribute.namespace.is_empty() && attribute.local == "typeface" && !attribute.value.is_empty()));
    Ok(Some(Theme { xml: format!("{}{}", xml::DECL, root.render(None)),
        has_fonts: fonts.is_some(),
        has_colors: elements.children().any(|child| child.namespace == DRAWING_NAMESPACE && child.local == "clrScheme"),
        has_minor_latin,
    }))
}

fn source_defaults(root: &Element, theme: &Result<Option<Theme>, String>) -> Result<Vec<Element>, String> {
    let declaration = root.child("docDefaults").and_then(|defaults| defaults.child("rPrDefault"));
    let original = declaration.and_then(|defaults| defaults.child("rPr"));
    if let Some(original) = original { validate_resource(original, WORD_NAMESPACE)?; }
    let mut scope = namespaces(root);
    if let Some(declaration) = root.child("docDefaults") { scope.extend(namespaces(declaration)); }
    if let Some(declaration) = declaration { scope.extend(namespaces(declaration)); }
    let mut defaults = original.map(|properties| scoped_children(properties, &scope)).unwrap_or_default();
    // Keep the reader's Word-compatible distinction between absent run defaults
    // (12 pt / minor theme fonts) and declared defaults with no size (10 pt).
    // Generated application-wide defaults must not replace that source context.
    if declaration.is_none() {
        let mut fonts = Element::word("rFonts");
        if theme.as_ref().ok().and_then(Option::as_ref).is_some_and(|theme| theme.has_minor_latin) {
            for (name, value) in [("asciiTheme", "minorHAnsi"), ("hAnsiTheme", "minorHAnsi"), ("eastAsiaTheme", "minorEastAsia"), ("cstheme", "minorBidi")] { fonts.set_word_attribute(name, value); }
        } else {
            fonts.set_word_attribute("ascii", "Aptos"); fonts.set_word_attribute("hAnsi", "Aptos");
        }
        defaults.push(fonts);
    }
    for name in ["sz", "szCs"] {
        if !defaults.iter().any(|property| property.is_word(name)) {
            let mut size = Element::word(name);
            size.set_word_attribute("val", if declaration.is_some() { "20" } else { "24" });
            defaults.push(size);
        }
    }
    Ok(defaults)
}

fn read_source(resource: &Value) -> Result<Source, String> {
    if get(resource, "version").as_u64() != Some(1) { return Err("Source style resource has an unsupported version.".into()); }
    if let Some(errors) = resource.get("resourceErrors") {
        let errors = errors.as_array().ok_or("Source resource errors have an invalid shape.")?;
        if errors.iter().any(|error| !error.is_string()) { return Err("Source resource errors contain a non-text value.".into()); }
        if let Some(error) = errors.first().and_then(Value::as_str) { return Err(format!("Original Office resource could not be retained: {error}")); }
    }
    let absent_styles = format!("<w:styles xmlns:w=\"{WORD_NAMESPACE}\"/>");
    let styles = get(resource, "styles");
    let xml = if styles.is_null() { absent_styles.as_str() } else { get(styles, "xml").as_str().ok_or("Original styles.xml was not retained.")? };
    let mut root = parse_xml(xml, true)?;
    canonicalize_office_namespaces(&mut root)?;
    if !root.is_word("styles") { return Err("Original styles.xml has an invalid root element.".into()); }
    let mut styles = BTreeMap::new();
    let mut default_paragraph = None;
    for style in root.children().filter(|child| child.is_word("style")) {
        let id = style.attribute("styleId").filter(|id| !id.is_empty() && id.len() <= 256).ok_or("Source style has an invalid style ID.")?;
        if style.attribute("type") == Some("paragraph") && style.attribute("default").is_some_and(|value| matches!(value, "1" | "true" | "on")) {
            if default_paragraph.is_some() { return Err("Original styles.xml has multiple default paragraph styles.".into()); }
            default_paragraph = Some(id.to_owned());
        }
        let mut retained = style.clone();
        attach_namespaces(&mut retained, &namespaces(&root));
        if styles.insert(id.to_owned(), retained).is_some() { return Err("Original styles.xml has duplicate style IDs.".into()); }
        if styles.len() > STYLE_LIMIT { return Err("Original styles.xml exceeds the supported style count.".into()); }
    }
    if default_paragraph.is_none() && styles.get("Normal").is_some_and(|style| style.attribute("type") == Some("paragraph")) { default_paragraph = Some("Normal".into()); }
    let theme = read_theme(get(resource, "theme"));
    let defaults = source_defaults(&root, &theme)?;
    let mut scope = namespaces(&root);
    let paragraph = root.child("docDefaults").and_then(|defaults| {
        scope.extend(namespaces(defaults));
        defaults.child("pPrDefault")
    }).and_then(|defaults| {
        scope.extend(namespaces(defaults));
        defaults.child("pPr")
    });
    let paragraph_defaults = if let Some(paragraph) = paragraph {
        validate_resource(paragraph, WORD_NAMESPACE)?;
        scoped_children(paragraph, &scope)
    } else { Vec::new() };
    Ok(Source { styles, default_paragraph, defaults, paragraph_defaults, namespaces: namespaces(&root), theme })
}

fn source_style(source: &Source, id: &str) -> Result<Element, String> {
    if let Some(style) = source.styles.get(id) { return Ok(style.clone()); }
    // This built-in means the containing paragraph's font. It may legally be
    // referenced without an explicit definition in the original package.
    if id == "DefaultParagraphFont" {
        let mut style = Element::word("style");
        style.set_word_attribute("type", "character"); style.set_word_attribute("styleId", id);
        let mut name = Element::word("name"); name.set_word_attribute("val", id);
        style.content.push(Content::Element(name));
        return Ok(style);
    }
    Err(format!("Original style {id:?} has no definition."))
}

fn style_chain(source: &Source, id: &str, kind: &str, active: &mut BTreeSet<String>) -> Result<Vec<Element>, String> {
    if active.len() >= STYLE_DEPTH_LIMIT || !active.insert(id.into()) { return Err(format!("Original style {id:?} has cyclic or excessive inheritance.")); }
    let style = source_style(source, id)?;
    if style.attribute("type") != Some(kind) { return Err(format!("Original style {id:?} is not a {kind} style.")); }
    if let Some(properties) = style.child("rPr") { validate_resource(properties, WORD_NAMESPACE)?; }
    let mut chain = if let Some(base) = style.child("basedOn").and_then(|value| value.attribute("val")) { style_chain(source, base, kind, active)? } else { Vec::new() };
    chain.push(style);
    active.remove(id);
    Ok(chain)
}

fn style_closure(source: &Source, id: &str, result: &mut BTreeMap<String, Element>, depth: usize) -> Result<(), String> {
    if result.contains_key(id) { return Ok(()); }
    if depth >= STYLE_DEPTH_LIMIT { return Err("Referenced source style dependencies exceed the maximum depth of 64.".into()); }
    if result.len() >= STYLE_LIMIT { return Err("Referenced source styles exceed the supported style count.".into()); }
    let style = source_style(source, id)?;
    if !matches!(style.attribute("type"), Some("character" | "paragraph")) { return Err(format!("Referenced style {id:?} has an unsupported type.")); }
    validate_resource(&style, WORD_NAMESPACE)?;
    if style.child("pPr").is_some_and(|paragraph| paragraph.child("numPr").is_some()) { return Err(format!("Linked style {id:?} uses original numbering IDs that could not be remapped.")); }
    if style.child("rPr").is_some_and(|run| run.child("rStyle").is_some()) { return Err(format!("Original style {id:?} delegates run properties to another rStyle.")); }
    let dependencies = ["basedOn", "next", "link"].into_iter().filter_map(|name| style.child(name).and_then(|child| child.attribute("val")).map(str::to_owned)).collect::<Vec<_>>();
    result.insert(id.into(), style);
    for dependency in dependencies { style_closure(source, &dependency, result, depth + 1)?; }
    Ok(())
}

fn theme_requirements(element: &Element) -> (bool, bool) {
    let fonts = element.attributes.iter().any(|attribute| attribute.namespace == WORD_NAMESPACE && matches!(attribute.local.as_str(), "asciiTheme" | "hAnsiTheme" | "eastAsiaTheme" | "cstheme"));
    let colors = element.attributes.iter().any(|attribute| attribute.namespace == WORD_NAMESPACE && matches!(attribute.local.as_str(), "themeColor" | "themeFill"));
    element.children().fold((fonts, colors), |(fonts, colors), child| { let requirements = theme_requirements(child); (fonts || requirements.0, colors || requirements.1) })
}

fn canonicalize_word(element: &mut Element) {
    if element.namespace == WORD_NAMESPACE { element.name = format!("w:{}", element.local); }
    element.attributes.retain(|attribute| attribute.name != "xmlns:w" || attribute.value == WORD_NAMESPACE);
    for attribute in &mut element.attributes { if attribute.namespace == WORD_NAMESPACE { attribute.name = format!("w:{}", attribute.local); } }
    for content in &mut element.content { if let Content::Element(child) = content { canonicalize_word(child); } }
}

fn attach_namespaces(element: &mut Element, namespaces: &BTreeMap<String, String>) {
    for (prefix, value) in namespaces {
        let name = if prefix.is_empty() { "xmlns".to_owned() } else { format!("xmlns:{prefix}") };
        if !element.attributes.iter().any(|attribute| attribute.name == name) { element.attributes.push(Attribute { name, namespace: XMLNS_NAMESPACE.into(), local: if prefix.is_empty() { "xmlns".into() } else { prefix.clone() }, value: value.clone() }); }
    }
}

impl SourceStyles {
    pub fn new(model: &Value, generated: String) -> Result<Self, ExportError> {
        let root = parse_xml(&generated, true).map_err(ExportError::invalid)?;
        let used_ids = root.children().filter_map(|style| style.attribute("styleId")).map(|id| id.to_ascii_lowercase()).collect();
        let resources = get(get(get(model, "metadata"), "source"), "officeSource");
        // Only the current ODT origin lacks a Word source-style context. A DOCX
        // with historical ODT resources still has its own native Word parts.
        let current_odt = get(get(resources, "origin"), "format").as_str() == Some("odt");
        let (source, source_error) = if resources.is_null() || current_odt { (None, None) } else { match read_source(resources) { Ok(source) => (Some(source), None), Err(error) => (None, Some(error)) } };
        Ok(Self { generated, source, source_error, imported: BTreeMap::new(), remap: BTreeMap::new(), used_ids, next_id: 1, theme_used: false })
    }

    pub fn restore_marker(&mut self, elements: &mut [Element], marker_namespaces: &BTreeMap<String, String>, block: &Value) -> Result<(), String> {
        let references = elements.iter().filter_map(|run| run.child("rPr").and_then(|run| run.child("rStyle")).and_then(|style| style.attribute("val"))).map(str::to_owned).collect::<BTreeSet<_>>();
        let needs_theme = elements.iter().any(|element| theme_requirements(element) != (false, false));
        let Some(source) = &self.source else {
            return if references.is_empty() && !needs_theme && self.source_error.is_none() { Ok(()) } else { Err(self.source_error.clone().unwrap_or_else(|| "Original styles.xml and its inheritance context were not retained.".into())) };
        };
        for (prefix, namespace) in &source.namespaces {
            if prefix != "w" && marker_namespaces.get(prefix).is_some_and(|original| original != namespace) { return Err(format!("Source style namespace prefix {prefix:?} conflicts with the numbering fragment.")); }
        }
        let paragraph_id = get(get(block, "metadata"), "style").get("styleId").and_then(Value::as_str)
            .or_else(|| get(get(get(get(block, "metadata"), "source"), "details"), "style").get("styleId").and_then(Value::as_str))
            .or_else(|| get(get(get(block, "metadata"), "source"), "details").get("styleId").and_then(Value::as_str))
            .or(source.default_paragraph.as_deref());
        let mut base = Properties::default();
        base.merge(&source.defaults, false);
        if let Some(paragraph_id) = paragraph_id { for style in style_chain(source, paragraph_id, "paragraph", &mut BTreeSet::new())? { base.merge(&properties(&style), true); } }
        let mut closure = BTreeMap::new();
        let mut chains = BTreeMap::new();
        for reference in &references {
            chains.insert(reference.clone(), style_chain(source, reference, "character", &mut BTreeSet::new())?);
            style_closure(source, reference, &mut closure, 0)?;
        }
        let mut replacements = Vec::new();
        let mut requirements = (false, false);
        for (index, run) in elements.iter().enumerate().filter(|(_, element)| element.is_word("r")) {
            let mut run_scope = marker_namespaces.clone();
            run_scope.extend(namespaces(run));
            let original = run.child("rPr");
            let reference = original.and_then(|properties| properties.child("rStyle"));
            let mut resolved = base.clone();
            if let Some(chain) = reference.and_then(|reference| reference.attribute("val")).and_then(|id| chains.get(id)) { for style in chain { resolved.merge(&properties(style), true); } }
            if let Some(original) = original { resolved.merge(&scoped_children(original, &run_scope), false); }
            if let Some(reference) = reference { resolved.0.insert(0, reference.clone()); }
            if resolved.0.is_empty() { continue; }
            let mut properties = Element::word("rPr");
            properties.attributes = original.map(|properties| properties.attributes.clone()).unwrap_or_default();
            properties.content = resolved.0.into_iter().map(Content::Element).collect();
            canonicalize_word(&mut properties);
            attach_namespaces(&mut properties, &run_scope);
            attach_namespaces(&mut properties, &source.namespaces);
            super::source_xml::validate_namespace_bindings(&properties, &BTreeMap::new())?;
            let needed = theme_requirements(&properties);
            requirements = (requirements.0 || needed.0, requirements.1 || needed.1);
            replacements.push((index, properties));
        }
        for style in closure.values() { let needed = theme_requirements(style); requirements = (requirements.0 || needed.0, requirements.1 || needed.1); }
        self.import_styles(&closure, requirements)?;
        for (index, mut properties) in replacements {
            for content in &mut properties.content { if let Content::Element(reference) = content { if reference.is_word("rStyle") { if let Some(mapped) = reference.attribute("val").and_then(|id| self.remap.get(id)).cloned() { reference.set_word_attribute("val", mapped); } } } }
            let run = &mut elements[index];
            if let Some(existing) = run.content.iter_mut().find(|content| matches!(content, Content::Element(properties) if properties.is_word("rPr"))) { *existing = Content::Element(properties); }
            else { run.content.insert(0, Content::Element(properties)); }
        }
        Ok(())
    }

    fn import_styles(&mut self, closure: &BTreeMap<String, Element>, requirements: (bool, bool)) -> Result<(), String> {
        let source = self.source.as_ref().ok_or_else(|| self.source_error.clone().unwrap_or_else(|| "Original styles.xml and its inheritance context were not retained.".into()))?;
        let requirements = closure.values().fold(requirements, |required, style| {
            let current = theme_requirements(style); (required.0 || current.0, required.1 || current.1)
        });
        if requirements != (false, false) {
            let theme = source.theme.as_ref().map_err(Clone::clone)?.as_ref().ok_or("Referenced source theme was not retained.")?;
            if requirements.0 && !theme.has_fonts || requirements.1 && !theme.has_colors { return Err("Referenced source theme is missing the required font or color scheme.".into()); }
            self.theme_used = true;
        }
        for (id, style) in closure {
            if self.remap.contains_key(id) { continue; }
            let mut mapped = id.clone();
            if self.used_ids.contains(&mapped.to_ascii_lowercase()) {
                loop {
                    mapped = format!("OfficeSourceStyle{}", self.next_id);
                    self.next_id = self.next_id.checked_add(1).ok_or("Imported style IDs exceed the supported range.")?;
                    if !self.used_ids.contains(&mapped.to_ascii_lowercase()) && !source.styles.keys().any(|id| id.eq_ignore_ascii_case(&mapped)) { break; }
                }
            }
            self.used_ids.insert(mapped.to_ascii_lowercase());
            self.remap.insert(id.clone(), mapped);
            self.imported.insert(id.clone(), style.clone());
        }
        Ok(())
    }

    pub fn styles_xml(&self) -> String {
        let Some(source) = &self.source else { return self.generated.clone(); };
        let mut imported = String::new();
        for (id, original) in &self.imported {
            let mut style = original.clone();
            style.attributes.retain(|attribute| !(attribute.namespace == WORD_NAMESPACE && attribute.local == "default"));
            style.set_word_attribute("styleId", self.remap[id].clone());
            for content in &mut style.content { if let Content::Element(reference) = content { if reference.namespace == WORD_NAMESPACE && matches!(reference.local.as_str(), "basedOn" | "next" | "link") {
                if let Some(mapped) = reference.attribute("val").and_then(|id| self.remap.get(id)).cloned() { reference.set_word_attribute("val", mapped); }
            } } }
            canonicalize_word(&mut style);
            imported.push_str(&style.render(Some(&source.namespaces)));
        }
        self.generated.replacen("</w:styles>", &format!("{imported}</w:styles>"), 1)
    }

    pub fn theme_xml(&self) -> Option<&str> {
        self.theme_used.then(|| self.source.as_ref()?.theme.as_ref().ok()?.as_ref().map(|theme| theme.xml.as_str())).flatten()
    }
}
