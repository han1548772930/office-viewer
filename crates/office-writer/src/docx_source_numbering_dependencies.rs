//! Resource reachability and validation, separate from numbering evaluation.
use std::collections::VecDeque;
use super::*;

#[derive(Clone, Ord, PartialOrd, Eq, PartialEq)]
enum Resource { Style(String), Number(u32), Abstract(u32), Picture(u32) }

pub(in super::super) fn has_live_numbering(node: &Element) -> bool {
    node.is_word("numId") && !unsigned(node.attribute("val"), "numId").is_ok_and(|id| id == 0)
        || node.is_word("abstractNumId") || node.is_word("lvlPicBulletId") || node.children().any(has_live_numbering)
}

fn source_numbering(source: &Source) -> Result<&SourceNumbering, String> {
    source.numbering.as_ref().map_err(Clone::clone)?.as_ref().ok_or_else(|| "DOCX_NUMBERING_SOURCE_MISSING: referenced original numbering.xml was not retained".into())
}

fn reference(node: &Element) -> Result<&str, String> {
    node.attribute("val").filter(|id| !id.is_empty() && id.len() <= 256).ok_or_else(|| format!("DOCX_NUMBERING_STYLE_REFERENCE: {} requires a nonempty style ID", node.local))
}

fn edges(source: &Source, node: &Element, queue: &mut VecDeque<(Resource, usize)>, depth: usize) -> Result<(), String> {
    if node.namespace == WORD_NAMESPACE {
        let expected = match node.local.as_str() {
            "pStyle" => Some(Some("paragraph")), "rStyle" => Some(Some("character")), "tblStyle" => Some(Some("table")),
            "numStyleLink" => Some(Some("numbering")), "styleLink" | "basedOn" | "next" | "link" => Some(None), _ => None,
        };
        if let Some(expected) = expected {
            let id = reference(node)?;
            if let Some(expected) = expected { style_chain(source, id, expected, &mut BTreeSet::new())?; }
            queue.push_back((Resource::Style(id.into()), depth));
        }
        match node.local.as_str() {
            "numId" => { let id = unsigned(node.attribute("val"), "numId")?; if id != 0 { queue.push_back((Resource::Number(id), depth)); } },
            "abstractNumId" => queue.push_back((Resource::Abstract(unsigned(node.attribute("val"), "abstractNumId")?), depth)),
            "lvlPicBulletId" => queue.push_back((Resource::Picture(unsigned(node.attribute("val"), "lvlPicBulletId")?), depth)), _ => {},
        }
    }
    for child in node.children() { edges(source, child, queue, depth)?; }
    Ok(())
}

fn inherited_number(source: &Source, style: &str) -> Result<u32, String> {
    let chain = style_chain(source, style, "numbering", &mut BTreeSet::new())?;
    for style in chain.iter().rev() {
        if let Some(id) = style.child("pPr").and_then(|properties| properties.child("numPr")).and_then(|properties| properties.child("numId")) {
            let id = unsigned(id.attribute("val"), "numStyleLink numId")?;
            if id != 0 { return Ok(id); }
            break;
        }
    }
    Err("DOCX_NUMBERING_STYLE_LINK: a delegated numbering style has no active numId".into())
}

/// styleLink/pStyle and a style's numPr are legal association back-links.
/// numStyleLink delegates to another abstract definition; its cycles really
/// recurse and must be rejected independently from ordinary resource reachability.
fn validate_delegation(source: &Source, id: u32, active: &mut BTreeSet<u32>) -> Result<(), String> {
    if active.len() >= STYLE_DEPTH_LIMIT || !active.insert(id) { return Err("DOCX_NUMBERING_CYCLE: source numbering contains cyclic or excessive numStyleLink delegation".into()); }
    let numbering = source_numbering(source)?;
    let definition = numbering.abstracts.get(&id).ok_or_else(|| format!("DOCX_NUMBERING_ABSTRACT_MISSING: abstractNum {id} has no retained definition"))?;
    if let Some(link) = definition.child("numStyleLink") {
        let num_id = inherited_number(source, reference(link)?)?;
        let number = numbering.numbers.get(&num_id).ok_or_else(|| format!("DOCX_NUMBERING_INSTANCE_MISSING: numId {num_id} has no retained instance"))?;
        let abstract_id = unsigned(number.child("abstractNumId").and_then(|node| node.attribute("val")), "delegated abstractNumId")?;
        validate_delegation(source, abstract_id, active)?;
    }
    active.remove(&id); Ok(())
}

fn collect_shape_types(source: &SourceNumbering, node: &Element, output: &mut BTreeMap<String, Element>, depth: usize) -> Result<(), String> {
    if depth >= STYLE_DEPTH_LIMIT { return Err("DOCX_NUMBERING_VML: shape type dependencies exceed their depth budget".into()); }
    if node.namespace == VML_NAMESPACE {
        if let Some(id) = plain_attribute(node, "type").and_then(|value| value.strip_prefix('#')) {
            if !output.contains_key(id) {
                let definition = if let Some(definition) = source.shape_types.get(id) { definition.clone() }
                    else if id == "_x0000_t75" {
                        // The Office image shape has a standard built-in type.
                        // Reuse the same template as the domain-list writer.
                        let wrapped = format!("<w:pict xmlns:w=\"{WORD_NAMESPACE}\" xmlns:v=\"{VML_NAMESPACE}\" xmlns:o=\"{OFFICE_NAMESPACE}\">{}</w:pict>",
                            super::super::super::numbering::picture_shape_type_xml());
                        let wrapper = parse_xml(&wrapped, false)?;
                        let mut definition = wrapper.children().next().ok_or("DOCX_NUMBERING_VML: the built-in image shape type is missing")?.clone();
                        attach_namespaces(&mut definition, &namespaces(&wrapper)); definition
                    } else { return Err(format!("DOCX_NUMBERING_VML: shape type {id:?} was not retained")); };
                validate_resource(&definition)?;
                output.insert(id.into(), definition.clone());
                collect_shape_types(source, &definition, output, depth + 1)?;
            }
        }
    }
    for child in node.children() { collect_shape_types(source, child, output, depth)?; }
    Ok(())
}

pub(in super::super) fn expand(source: &Source, initial: &BTreeMap<String, Element>, extra: Option<&Element>) -> Result<Dependencies, String> {
    let mut output = Dependencies::default();
    let mut queue = initial.keys().map(|id| (Resource::Style(id.clone()), 0)).collect::<VecDeque<_>>();
    if let Some(extra) = extra { edges(source, extra, &mut queue, 0)?; }
    let mut visited = BTreeSet::new();
    while let Some((key, depth)) = queue.pop_front() {
        if visited.contains(&key) { continue; }
        if depth >= STYLE_DEPTH_LIMIT { return Err("DOCX_NUMBERING_RESOURCE_LIMIT: source style/numbering dependencies exceed depth 64".into()); }
        visited.insert(key.clone());
        let node = match key {
            Resource::Style(id) => {
                if output.styles.len() >= STYLE_LIMIT { return Err("DOCX_NUMBERING_RESOURCE_LIMIT: too many referenced source styles".into()); }
                let node = source_style(source, &id)?;
                let kind = node.attribute("type").filter(|kind| matches!(*kind, "paragraph" | "character" | "table" | "numbering"))
                    .ok_or("DOCX_NUMBERING_STYLE_REFERENCE: source style has an unsupported type")?;
                style_chain(source, &id, kind, &mut BTreeSet::new())?;
                super::super::validate_resource(&node, WORD_NAMESPACE)?;
                output.styles.insert(id, node.clone()); node
            },
            Resource::Number(id) => {
                let node = source_numbering(source)?.numbers.get(&id).ok_or_else(|| format!("DOCX_NUMBERING_INSTANCE_MISSING: numId {id} has no retained instance"))?.clone();
                if node.child("abstractNumId").is_none() { return Err("DOCX_NUMBERING_ABSTRACT_MISSING: a source numbering instance has no abstractNumId".into()); }
                validate_resource(&node)?; output.numbers.insert(id, node.clone()); node
            },
            Resource::Abstract(id) => {
                let node = source_numbering(source)?.abstracts.get(&id).ok_or_else(|| format!("DOCX_NUMBERING_ABSTRACT_MISSING: abstractNum {id} has no retained definition"))?.clone();
                validate_resource(&node)?; validate_delegation(source, id, &mut BTreeSet::new())?;
                output.abstracts.insert(id, node.clone()); node
            },
            Resource::Picture(id) => {
                let numbering = source_numbering(source)?;
                let node = numbering.pictures.get(&id).ok_or_else(|| format!("DOCX_NUMBERING_PICTURE_MISSING: numPicBullet {id} has no retained definition"))?.clone();
                validate_resource(&node)?;
                collect_shape_types(numbering, &node, &mut output.shape_types, 0)?;
                output.pictures.insert(id, node.clone()); node
            },
        };
        edges(source, &node, &mut queue, depth + 1)?;
    }
    Ok(output)
}
