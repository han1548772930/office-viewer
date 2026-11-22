use quick_xml::{Reader, XmlVersion, events::{BytesStart, Event}};
use serde::Serialize;
use std::io::{Cursor, Read};
use zip::ZipArchive;

#[derive(Debug, Default, Clone, Serialize)]
pub struct StoryText {
    pub text: String,
    pub fields: Vec<StoryField>,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct StoryField {
    pub field: String,
    pub offset: usize,
    pub instruction: String,
    pub cached_result: String,
}

pub fn read_story(archive: &mut ZipArchive<Cursor<&[u8]>>, part: &str) -> Option<StoryText> {
    let entry = archive.by_name(part.trim_start_matches('/')).ok()?;
    const LIMIT: u64 = 16 * 1024 * 1024;
    if entry.size() > LIMIT { return None; }
    let mut xml = String::new();
    entry.take(LIMIT + 1).read_to_string(&mut xml).ok()?;
    if xml.len() as u64 > LIMIT { return None; }
    parse_story(&xml)
}

fn attr(reader: &Reader<&[u8]>, tag: &BytesStart<'_>, key: &[u8]) -> Option<String> {
    let a = tag.attributes().flatten().find(|a| a.key.local_name().as_ref() == key)?;
    Some(a.decoded_and_normalized_value_with(XmlVersion::Implicit1_0, reader.decoder(), 1, quick_xml::escape::resolve_xml_entity).ok()?.into_owned())
}

#[derive(Default)]
struct Pending { instruction: String, cached: String }

fn finish(story: &mut StoryText, pending: Pending, offset: &mut usize) {
    let instruction = pending.instruction.trim();
    let name = instruction.split_whitespace().next().unwrap_or_default().to_ascii_uppercase();
    if matches!(name.as_str(), "PAGE" | "NUMPAGES" | "DATE" | "TIME") {
        story.fields.push(StoryField { field: name, offset: *offset, instruction: instruction.into(), cached_result: pending.cached });
    } else {
        story.text.push_str(&pending.cached);
        *offset += pending.cached.encode_utf16().count();
    }
}

fn append(story: &mut StoryText, pending: &mut Option<Pending>, offset: &mut usize, value: &str) {
    if let Some(field) = pending { field.cached.push_str(value); }
    else { story.text.push_str(value); *offset += value.encode_utf16().count(); }
}

pub(crate) fn parse_story(xml: &str) -> Option<StoryText> {
    let mut reader = Reader::from_str(xml);
    let mut story = StoryText::default();
    let mut pending: Option<Pending> = None;
    let (mut text, mut instr, mut seen_p, mut offset) = (false, false, false, 0);
    loop {
        let event = reader.read_event().ok()?;
        let empty = matches!(&event, Event::Empty(_));
        match event {
            Event::Start(e) | Event::Empty(e) => match e.local_name().as_ref() {
                b"p" => { if seen_p { append(&mut story, &mut pending, &mut offset, "\n"); } seen_p = true; }
                b"t" => text = !empty,
                b"instrText" => instr = !empty,
                b"br" | b"cr" => append(&mut story, &mut pending, &mut offset, "\n"),
                b"tab" => append(&mut story, &mut pending, &mut offset, "\t"),
                b"fldSimple" => {
                    if pending.is_some() { return None; }
                    let field = Pending { instruction: attr(&reader, &e, b"instr")?, ..Default::default() };
                    if empty { finish(&mut story, field, &mut offset); } else { pending = Some(field); }
                }
                b"fldChar" => match attr(&reader, &e, b"fldCharType").as_deref() {
                    Some("begin") if pending.is_none() => pending = Some(Pending::default()),
                    Some("begin") => return None,
                    Some("end") => finish(&mut story, pending.take()?, &mut offset),
                    _ => {}
                },
                _ => {}
            },
            Event::End(e) => match e.local_name().as_ref() {
                b"t" => text = false,
                b"instrText" => instr = false,
                b"fldSimple" => finish(&mut story, pending.take()?, &mut offset),
                _ => {}
            },
            Event::Eof => return pending.is_none().then_some(story),
            _ if text || instr => {
                let s = match event {
                    Event::Text(t) => t.decode().ok()?.into_owned(),
                    Event::CData(t) => t.decode().ok()?.into_owned(),
                    Event::GeneralRef(r) => if let Some(c) = r.resolve_char_ref().ok()? { c.to_string() }
                        else { quick_xml::escape::resolve_xml_entity(&r.decode().ok()?)?.to_owned() },
                    _ => continue,
                };
                if instr { pending.as_mut()?.instruction.push_str(&s); }
                else { append(&mut story, &mut pending, &mut offset, &s); }
            }
            _ => {}
        }
    }
}
