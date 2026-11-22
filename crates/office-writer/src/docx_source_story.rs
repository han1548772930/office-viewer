use super::{BTreeMap, BTreeSet, Content, Element, Properties, SourceStyles, WORD_NAMESPACE, attach_namespaces, canonicalize_office_namespaces, canonicalize_word, child_properties, get, namespaces, parse_xml, properties, scoped_children, style_chain, style_closure, theme_requirements, validate_resource};
use serde_json::Value;

fn metadata_source_format(provenance: &Value) -> &Value {
    get(get(get(provenance, "metadata"), "source"), "sourceFormat")
}

fn read_properties(format: &Value, local: &str) -> Result<Option<Element>, String> {
    if format.is_null() { return Ok(None); }
    if get(format, "version").as_u64() != Some(1) { return Err("Story source-format fragment has an unsupported version.".into()); }
    let Some(xml) = format.get("propertiesXml") else { return Ok(None); };
    let xml = xml.as_str().ok_or("Story source-format fragment has invalid XML.")?;
    let mut wrapper = Element::word("sourceContext");
    if let Some(declared) = format.get("namespaces") {
        let declared = declared.as_object().ok_or("Story source-format fragment has invalid namespace bindings.")?;
        let mut bindings = BTreeMap::new();
        for (prefix, namespace) in declared {
            if !prefix.is_empty() && !super::super::source_xml::valid_name(prefix) { return Err("Story source-format fragment has an invalid namespace prefix.".into()); }
            bindings.insert(prefix.clone(), namespace.as_str().ok_or("Story source-format fragment has an invalid namespace URI.")?.to_owned());
        }
        attach_namespaces(&mut wrapper, &bindings);
    }
    if !wrapper.attributes.iter().any(|attribute| attribute.name == "xmlns:w") {
        attach_namespaces(&mut wrapper, &BTreeMap::from([("w".into(), WORD_NAMESPACE.into())]));
    }
    let wrapped = format!("{}{xml}</w:sourceContext>", wrapper.opening(None, false));
    let mut wrapper = parse_xml(&wrapped, false)?;
    canonicalize_office_namespaces(&mut wrapper)?;
    let mut children = wrapper.children();
    let mut result = children.next().ok_or("Story source-format fragment has no properties element.")?.clone();
    if children.next().is_some() || !result.is_word(local) || !wrapper.text().trim().is_empty() {
        return Err("Story source-format fragment has an invalid properties root.".into());
    }
    attach_namespaces(&mut result, &namespaces(&wrapper));
    validate_resource(&result, WORD_NAMESPACE)?;
    Ok(Some(result))
}

fn generated_properties(xml: &str, local: &str) -> Result<Element, String> {
    parse_xml(&format!("<w:{local} xmlns:w=\"{WORD_NAMESPACE}\">{xml}</w:{local}>"), false)
}

fn replace_properties(container: &mut Element, local: &str, properties: Element) {
    if let Some(current) = container.content.iter_mut().find(|value| matches!(value, Content::Element(element) if element.is_word(local))) {
        *current = Content::Element(properties);
    } else { container.content.insert(0, Content::Element(properties)); }
}

fn merge_attributes(current: &mut Element, next: &Element) {
    for attribute in &next.attributes {
        if let Some(prior) = current.attributes.iter_mut().find(|prior| prior.namespace == attribute.namespace && prior.local == attribute.local) { *prior = attribute.clone(); }
        else { current.attributes.push(attribute.clone()); }
    }
}

/// Paragraph spacing/indent attributes and tab/border children inherit
/// independently. Do not flatten a style by replacing its entire container.
fn merge_paragraph(current: &mut Properties, incoming: &[Element]) {
    for property in incoming {
        let previous = current.0.iter().position(|candidate| candidate.namespace == property.namespace && candidate.local == property.local);
        let Some(index) = previous else { current.0.push(property.clone()); continue; };
        let prior = &mut current.0[index];
        if property.is_word("ind") || property.is_word("spacing") {
            let conflicts: &[(&str, &str)] = if property.is_word("ind") {
                &[("firstLine", "hanging"), ("firstLineChars", "hangingChars"), ("left", "leftChars"), ("right", "rightChars"), ("firstLine", "firstLineChars"), ("hanging", "hangingChars")]
            } else { &[("before", "beforeLines"), ("after", "afterLines")] };
            for &(first, second) in conflicts {
                if property.attribute(first).is_some() && property.attribute(second).is_none() { prior.attributes.retain(|attribute| !(attribute.namespace == WORD_NAMESPACE && attribute.local == second)); }
                if property.attribute(second).is_some() && property.attribute(first).is_none() { prior.attributes.retain(|attribute| !(attribute.namespace == WORD_NAMESPACE && attribute.local == first)); }
            }
            merge_attributes(prior, property);
        } else if property.is_word("tabs") || property.is_word("pBdr") {
            let prior_scope = namespaces(prior);
            for child in &mut prior.content { if let Content::Element(child) = child { attach_namespaces(child, &prior_scope); } }
            merge_attributes(prior, property);
            for mut next in scoped_children(property, &BTreeMap::new()) {
                let index = prior.content.iter().position(|candidate| matches!(candidate, Content::Element(candidate) if
                    candidate.namespace == next.namespace && if property.is_word("tabs") { candidate.attribute("pos") == next.attribute("pos") } else { candidate.local == next.local }));
                canonicalize_word(&mut next);
                if let Some(index) = index { prior.content[index] = Content::Element(next); } else { prior.content.push(Content::Element(next)); }
            }
        } else if property.is_word("rPr") {
            let reference = property.child("rStyle").or_else(|| prior.child("rStyle")).cloned();
            let mut runs = Properties::default();
            runs.merge(&scoped_children(prior, &BTreeMap::new()), false);
            runs.merge(&scoped_children(property, &BTreeMap::new()), false);
            if let Some(reference) = reference { runs.0.insert(0, reference); }
            merge_attributes(prior, property);
            prior.content = runs.0.into_iter().map(Content::Element).collect();
        } else { *prior = property.clone(); }
    }
}

fn changed(current: &Value, baseline: &Value, fields: &[&str]) -> bool {
    fields.iter().any(|field| get(current, field) != get(baseline, field))
}

fn overlay_properties(result: &mut Element, generated: &Element, selected: &[&str]) {
    for name in selected {
        result.content.retain(|content| !matches!(content, Content::Element(property) if property.is_word(name)));
        for property in generated.children().filter(|property| property.is_word(name)) { result.content.push(Content::Element(property.clone())); }
    }
}

fn explicit_reset(result: &mut Element, local: &str, value: &str) {
    if result.child(local).is_none() {
        let mut property = Element::word(local); property.set_word_attribute("val", value);
        result.content.push(Content::Element(property));
    }
}

impl SourceStyles {
    fn remap_property_references(&self, properties: &mut Element) {
        for content in &mut properties.content {
            if let Content::Element(reference) = content {
                if reference.is_word("rStyle") || reference.is_word("pStyle") {
                    if let Some(mapped) = reference.attribute("val").and_then(|id| self.remap.get(id)).cloned() { reference.set_word_attribute("val", mapped); }
                }
            }
        }
    }

    fn resolved_story_run(&mut self, original: &Element, paragraph_id: Option<&str>) -> Result<Element, String> {
        validate_resource(original, WORD_NAMESPACE)?;
        let reference = original.child("rStyle").and_then(|style| style.attribute("val"));
        let Some(source) = self.source.as_ref() else {
            if self.source_error.is_some() || reference.is_some() || paragraph_id.is_some() || theme_requirements(original) != (false, false) {
                return Err(self.source_error.clone().unwrap_or_else(|| "Story style/theme dependencies were not retained.".into()));
            }
            return Ok(original.clone());
        };
        let mut resolved = Properties::default();
        resolved.merge(&source.defaults, false);
        if let Some(id) = paragraph_id.or(source.default_paragraph.as_deref()) {
            for style in style_chain(source, id, "paragraph", &mut BTreeSet::new())? { resolved.merge(&properties(&style), true); }
        }
        let mut closure = BTreeMap::new();
        if let Some(id) = reference {
            for style in style_chain(source, id, "character", &mut BTreeSet::new())? { resolved.merge(&properties(&style), true); }
            style_closure(source, id, &mut closure, 0)?;
        }
        resolved.merge(&scoped_children(original, &BTreeMap::new()), false);
        if let Some(reference) = original.child("rStyle") { resolved.0.insert(0, reference.clone()); }
        let mut result = Element::word("rPr");
        result.attributes = original.attributes.clone();
        result.content = resolved.0.into_iter().map(Content::Element).collect();
        attach_namespaces(&mut result, &source.namespaces);
        canonicalize_word(&mut result);
        self.import_styles(&closure, theme_requirements(&result))?;
        self.remap_property_references(&mut result);
        Ok(result)
    }

    fn resolved_story_paragraph(&mut self, original: &Element, paragraph_id: Option<&str>) -> Result<Element, String> {
        validate_resource(original, WORD_NAMESPACE)?;
        if original.child("numPr").is_some() { return Err("Story paragraph uses original numbering IDs whose resource mapping remains unimplemented.".into()); }
        let Some(source) = self.source.as_ref() else {
            if self.source_error.is_some() || paragraph_id.is_some() || theme_requirements(original) != (false, false) {
                return Err(self.source_error.clone().unwrap_or_else(|| "Story paragraph style/theme dependencies were not retained.".into()));
            }
            return Ok(original.clone());
        };
        let id = paragraph_id.or(source.default_paragraph.as_deref());
        let mut merged = Properties::default();
        merge_paragraph(&mut merged, &source.paragraph_defaults);
        let mut closure = BTreeMap::new();
        if let Some(id) = id {
            for style in style_chain(source, id, "paragraph", &mut BTreeSet::new())? {
                let properties = child_properties(&style, "pPr");
                merge_paragraph(&mut merged, &properties);
            }
            style_closure(source, id, &mut closure, 0)?;
        }
        merge_paragraph(&mut merged, &scoped_children(original, &BTreeMap::new()));
        if merged.0.iter().any(|property| property.is_word("numPr")) { return Err("Story paragraph defaults use original numbering IDs whose resource mapping remains unimplemented.".into()); }
        if let Some(id) = id {
            let mut reference = Element::word("pStyle"); reference.set_word_attribute("val", id);
            merged.0.retain(|property| !property.is_word("pStyle")); merged.0.insert(0, reference);
        }
        let mut result = Element::word("pPr");
        result.attributes = original.attributes.clone();
        result.content = merged.0.into_iter().map(Content::Element).collect();
        attach_namespaces(&mut result, &source.namespaces);
        canonicalize_word(&mut result);
        self.import_styles(&closure, theme_requirements(&result))?;
        self.remap_property_references(&mut result);
        Ok(result)
    }

    /// Preserve source rPr, including unmodeled properties, until a modeled
    /// property changes. A shared clear command explicitly drops the source rPr.
    pub(in super::super) fn story_run_properties(&mut self, xml: &str, current: &Value, provenance: &Value) -> Result<String, String> {
        let metadata = get(provenance, "metadata");
        let format = metadata_source_format(provenance);
        let generated = generated_properties(xml, "rPr")?;
        if format.is_null() || get(metadata, "officeStoryTextFormatCleared").as_bool() == Some(true) { return Ok(generated.render(None)); }
        let original = read_properties(format, "rPr")?.unwrap_or_else(|| Element::word("rPr"));
        let mut result = self.resolved_story_run(&original, get(format, "paragraphStyleId").as_str())?;
        let baseline = get(metadata, "officeStoryStyleBaseline");
        // Legacy fragments without a normalized baseline keep their advanced
        // properties, while the modeled properties retain application semantics.
        let changed_fields = |fields: &[&str]| baseline.is_null() || changed(current, baseline, fields);
        for (fields, properties) in [
            (&["bold"][..], &["b", "bCs"][..]), (&["italic"][..], &["i", "iCs"][..]),
            (&["underline"][..], &["u"][..]), (&["strike"][..], &["strike", "dstrike"][..]),
            (&["fontSize"][..], &["sz", "szCs"][..]), (&["color"][..], &["color"][..]),
            (&["highlight"][..], &["highlight", "shd"][..]), (&["verticalAlign"][..], &["vertAlign"][..]),
        ] {
            if changed_fields(fields) { overlay_properties(&mut result, &generated, properties); }
        }
        if changed_fields(&["highlight"]) { explicit_reset(&mut result, "highlight", "none"); explicit_reset(&mut result, "shd", "nil"); }
        if changed_fields(&["verticalAlign"]) { explicit_reset(&mut result, "vertAlign", "baseline"); }
        let latin_changed = changed_fields(&["fontFamily"]);
        let asian_changed = changed_fields(&["eastAsiaFontFamily"]) || latin_changed && get(current, "eastAsiaFontFamily").is_null();
        if latin_changed || asian_changed {
            let mut fonts = result.child("rFonts").cloned().unwrap_or_else(|| Element::word("rFonts"));
            for (changed, slots) in [(latin_changed, &["ascii", "hAnsi"][..]), (asian_changed, &["eastAsia"][..])] {
                if !changed { continue; }
                for slot in slots {
                    fonts.attributes.retain(|attribute| !(attribute.namespace == WORD_NAMESPACE && (attribute.local == *slot || attribute.local == format!("{slot}Theme"))));
                    if let Some(value) = generated.child("rFonts").and_then(|fonts| fonts.attribute(slot)) { fonts.set_word_attribute(slot, value); }
                }
            }
            replace_properties(&mut result, "rFonts", fonts);
        }
        super::super::source_xml::validate_namespace_bindings(&result, &BTreeMap::from([("w".into(), WORD_NAMESPACE.into())]))?;
        Ok(result.render(None))
    }

    pub(in super::super) fn story_paragraph_properties(&mut self, xml: &str, current: &Value, provenance: &Value) -> Result<String, String> {
        let metadata = get(provenance, "metadata");
        let format = metadata_source_format(provenance);
        let generated = generated_properties(xml, "pPr")?;
        if format.is_null() || get(metadata, "officeStoryParagraphFormatCleared").as_bool() == Some(true) { return Ok(generated.render(None)); }
        let original = read_properties(format, "pPr")?.unwrap_or_else(|| Element::word("pPr"));
        let paragraph_id = original.child("pStyle").and_then(|style| style.attribute("val")).or_else(|| get(format, "paragraphStyleId").as_str());
        let mut result = self.resolved_story_paragraph(&original, paragraph_id)?;
        let baseline = get(metadata, "officeStoryParagraphStyleBaseline");
        for (fields, properties) in [
            (&["align"][..], &["jc"][..]),
            (&["indent", "indentStart", "indentEnd", "indentFirstLine", "hanging"][..], &["ind"][..]),
            (&["spacingRule", "lineHeight", "spaceBefore", "spaceAfter"][..], &["spacing"][..]),
            (&["wordWrap"][..], &["wordWrap"][..]), (&["snapToGrid"][..], &["snapToGrid"][..]),
        ] {
            if baseline.is_null() || changed(current, baseline, fields) { overlay_properties(&mut result, &generated, properties); }
        }
        if let Some(mut mark) = result.child("rPr").cloned() {
            attach_namespaces(&mut mark, &namespaces(&result));
            let mark = self.resolved_story_run(&mark, paragraph_id)?;
            replace_properties(&mut result, "rPr", mark);
        }
        super::super::source_xml::validate_namespace_bindings(&result, &BTreeMap::from([("w".into(), WORD_NAMESPACE.into())]))?;
        Ok(result.render(None))
    }

    fn restore_story_inline(&mut self, node: &mut Element, paragraph_id: Option<&str>, inherited: &BTreeMap<String, String>) -> Result<(), String> {
        if node.is_word("p") { return self.restore_story_paragraph(node, inherited); }
        let mut scope = inherited.clone();
        scope.extend(namespaces(node));
        if node.is_word("r") {
            let mut original = node.child("rPr").cloned().unwrap_or_else(|| Element::word("rPr"));
            attach_namespaces(&mut original, &scope);
            let resolved = self.resolved_story_run(&original, paragraph_id)?;
            replace_properties(node, "rPr", resolved);
        }
        for content in &mut node.content {
            if let Content::Element(child) = content {
                if !child.is_word("rPr") { self.restore_story_inline(child, paragraph_id, &scope)?; }
            }
        }
        Ok(())
    }

    fn restore_story_paragraph(&mut self, node: &mut Element, inherited: &BTreeMap<String, String>) -> Result<(), String> {
        let mut scope = inherited.clone();
        scope.extend(namespaces(node));
        let mut original = node.child("pPr").cloned().unwrap_or_else(|| Element::word("pPr"));
        attach_namespaces(&mut original, &scope);
        let paragraph_id = original.child("pStyle").and_then(|style| style.attribute("val")).map(str::to_owned);
        let mut resolved = self.resolved_story_paragraph(&original, paragraph_id.as_deref())?;
        if let Some(mut mark) = resolved.child("rPr").cloned() {
            attach_namespaces(&mut mark, &namespaces(&resolved));
            let mark = self.resolved_story_run(&mark, paragraph_id.as_deref())?;
            replace_properties(&mut resolved, "rPr", mark);
        }
        replace_properties(node, "pPr", resolved);
        for content in &mut node.content {
            if let Content::Element(child) = content {
                if !child.is_word("pPr") { self.restore_story_inline(child, paragraph_id.as_deref(), &scope)?; }
            }
        }
        Ok(())
    }

    /// Used only for an unchanged, validated rich story. Source IDs are
    /// remapped through the same registry as preserved body numbering markers.
    pub(in super::super) fn restore_story(&mut self, root: &mut Element) -> Result<(), String> {
        canonicalize_office_namespaces(root)?;
        fn visit(styles: &mut SourceStyles, node: &mut Element, inherited: &BTreeMap<String, String>) -> Result<(), String> {
            if node.is_word("p") { return styles.restore_story_paragraph(node, inherited); }
            if node.is_word("tblStyle") { return Err("Story table-style resource mapping remains unimplemented.".into()); }
            let mut scope = inherited.clone();
            scope.extend(namespaces(node));
            for content in &mut node.content { if let Content::Element(child) = content { visit(styles, child, &scope)?; } }
            Ok(())
        }
        visit(self, root, &BTreeMap::new())?;
        super::super::source_xml::validate_namespace_bindings(root, &BTreeMap::new())
    }
}
