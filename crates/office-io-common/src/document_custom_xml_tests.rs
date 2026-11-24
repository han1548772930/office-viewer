//! Deferred DOC-11/DOC-22 contracts; run during the unified validation phase.
use super::*;

fn source(name: &str, xml: &str) -> DocumentCustomXmlPart {
    DocumentCustomXmlPart { part_name: name.into(), xml: xml.into(), content_type: None, store_item_id: Some("{STORE}".into()),
        properties_part_name: None, properties_xml: None, relationships_xml: None, related_parts: Vec::new(), encoding: None, byte_order_mark: None }
}

#[test]
fn package_names_keep_unicode_spaces_percent_and_ascii_case_identity() {
    let name = canonical_part_name("CustomXml/名 前%20😀.xml").unwrap();
    assert_eq!(name, "CustomXml/%E5%90%8D%20%E5%89%8D%20%F0%9F%98%80.xml");
    assert_eq!(canonical_part_name(&name).unwrap(), name);
    assert_eq!(zip_part_name(&name).unwrap(), "CustomXml/名 前 😀.xml");
    assert_eq!(part_name_key("CustomXml/%6Eame.xml").unwrap(), part_name_key("customxml/NAME.XML").unwrap());
    assert_eq!(part_name_from_zip("customXml/a%.xml").unwrap(), "customXml/a%25.xml");
    for value in ["", "/customXml/a.xml", "x//a.xml", "x/../a.xml", "x/%2e%2e/a.xml", "x/a%2Fb.xml", "x/a%5cb.xml", "x/a%.xml", "x/a%ff.xml", "x/a."] {
        assert!(canonical_part_name(value).is_err(), "{value}");
    }
}

#[test]
fn relationships_resolve_owner_relative_names_and_reject_package_escape() {
    assert_eq!(resolve_target("customXml/nested/item.xml", "../media/图 1.bin").unwrap(), "customXml/media/%E5%9B%BE%201.bin");
    assert_eq!(resolve_target("word/document.xml", "/customXml/item%25.xml").unwrap(), "customXml/item%25.xml");
    assert_eq!(relationship_owner("CUSTOMXML/_RELS/A%20B.XML.RELS").unwrap(), "CUSTOMXML/A%20B.XML");
    assert!(relationship_owner("_rels/.rels").is_none());
    for target in ["../../outside.xml", "//outside/a.xml", "a.xml?x=1", "a.xml#value", "http://example.test/a.xml", "%2e%2e/a.xml"] {
        assert!(resolve_target("customXml/item.xml", target).is_err(), "{target}");
    }
}

#[test]
fn mime_and_shared_resources_use_ascii_tokens_and_case_insensitive_type_semantics() {
    for value in ["Application/XML", "application/vnd.example+xml", "application/octet-stream"] { assert!(valid_content_type(value)); }
    for value in ["application/xml\n", "application/xml\r\n", "application/xmK", "application//xml", "application/xml; charset=utf8"] { assert!(!valid_content_type(value)); }
    let mut a = source("customXml/a.xml", "<a/>");
    let resource = DocumentCustomXmlResource { part_name: "customXml/shared.bin".into(), content_type: "Application/Octet-Stream".into(), data_base64: "AA==".into() };
    a.related_parts.push(resource.clone());
    let mut b = source("customXml/b.xml", "<b/>"); b.store_item_id = Some("{OTHER}".into());
    b.related_parts.push(DocumentCustomXmlResource { content_type: "application/octet-stream".into(), ..resource });
    assert!(validate_parts(&[a.clone(), b.clone()]).is_ok());
    b.related_parts[0].data_base64 = "AQ==".into();
    assert!(validate_parts(&[a, b]).is_err());
}

#[test]
fn literal_declaration_text_in_comments_cdata_and_processing_instructions_is_retained() {
    let xml = "<?xml version='1.0'?><!-- <!DOCTYPE fake> --><?example text='<!ENTITY fake>'?><root><![CDATA[<!DOCTYPE example>]]>&amp;&#x4e2d;</root>";
    validate_xml(xml).unwrap();
    let encoded = encode_xml(&source("customXml/a.xml", xml)).unwrap();
    assert_eq!(String::from_utf8(encoded).unwrap(), xml);
    for invalid in ["<!DOCTYPE root><root/>", "<!DOCTYPE root [<!ENTITY x 'value'>]><root>&x;</root>", "<root>&unknown;</root>"] { assert!(validate_xml(invalid).is_err()); }
}

#[test]
fn namespace_entities_are_normalized_before_identity_and_reserved_prefix_checks() {
    let xml = "<ds:datastoreItem xmlns:ds='http://schemas.openxmlformats.org/officeDocument/2006/custom&#88;ml' ds:itemID='{STORE}'/>";
    assert_eq!(properties_store_id(&xml).unwrap(), "{STORE}");
    assert!(validate_xml("<root xmlns:xml='http://www.w3.org/XML/1998/namesp&#97;ce' xml:lang='zh-CN'/>").is_ok());
    for xml in ["<root xmlns:a='urn:x' xmlns:b='urn:&#x78;' a:value='1' b:value='2'/>",
        "<root xmlns='http://www.w3.org/2000/xmlns/'/>", "<root xmlns:p='' />", "<xmlns:root/>", "<root xmlns:p='urn:x' xmlns:p='urn:y'/>"] {
        assert!(validate_xml(xml).is_err(), "{xml}");
    }
    assert!(validate_xml("<root xmlns:a='urn:X' xmlns:b='urn:x' a:value='1' b:value='2'/>").is_ok());
}

#[test]
fn xml_names_declaration_order_text_and_character_boundaries_are_checked() {
    for xml in ["<0root/>", "<a:b:c/>", "<root a='<'/>", "<root>]]></root>", "<root>&#0;</root>", "<root>\u{fffe}</root>",
        "\u{a0}<root/>", "<root/><?xml version='1.0'?>", "<!--comment--><?xml version='1.0'?><root/>",
        "<?xml encoding='UTF-8' version='1.0'?><root/>", "<?xml version='1.0' version='1.0'?><root/>",
        "<?xml version='1.0'encoding='UTF-8'?><root/>", "<?xml version='1.0' standalone='maybe'?><root/>", "<?XML value?><root/>"] {
        assert!(validate_xml(xml).is_err(), "{xml}");
    }
    validate_xml("<?xml version='1.0' encoding='UTF-8' standalone='yes'?><根 属性='中'><child/></根>").unwrap();
}

#[test]
fn utf16_and_bom_outputs_reopen_with_matching_declarations_and_surrogate_pairs() {
    for (encoding, declaration) in [(CustomXmlEncoding::Utf8, "UTF-8"), (CustomXmlEncoding::Utf16le, "UTF-16LE"), (CustomXmlEncoding::Utf16be, "UTF-16BE")] {
        for bom in [false, true] {
            let mut part = source("customXml/a.xml", "<?xml version='1.0' encoding='UTF8'?><root>中😀</root>");
            part.encoding = Some(encoding); part.byte_order_mark = Some(bom);
            let bytes = encode_xml(&part).unwrap();
            let (xml, actual_encoding, actual_bom) = decode_xml(&bytes).unwrap();
            assert_eq!(actual_encoding, encoding); assert_eq!(actual_bom, bom);
            assert_eq!(xml, part.xml.replace("UTF8", declaration));
        }
    }
    assert!(decode_xml(&[0xff, 0xfe, 0x3c]).is_err());
    assert!(decode_xml(&[0xff, 0xfe, 0x00, 0xd8]).is_err());
}

#[test]
fn encoded_budget_counts_utf16_bom_and_encoding_declaration_expansion() {
    let xml = format!("<r>{}</r>", "x".repeat(MAX_PART_BYTES / 2 - 7));
    let mut part = source("customXml/a.xml", &xml); part.encoding = Some(CustomXmlEncoding::Utf16le);
    assert!(validate_parts(&[part.clone()]).is_ok());
    part.byte_order_mark = Some(true);
    assert!(validate_parts(&[part]).unwrap_err().contains("RESOURCE_LIMIT"));
    let (prefix, suffix) = ("<?xml version='1.0' encoding='UTF8'?><r>", "</r>");
    let xml = format!("{prefix}{}{suffix}", "x".repeat(MAX_PART_BYTES - prefix.len() - suffix.len()));
    assert!(validate_parts(&[source("customXml/a.xml", &xml)]).unwrap_err().contains("RESOURCE_LIMIT"));
}

#[test]
fn canonical_base64_checks_padding_trailing_bits_and_empty_resources() {
    for (value, size) in [("", 0), ("AA==", 1), ("AAA=", 2), ("AAAA", 3)] { assert_eq!(resource_byte_length(value).unwrap(), size); }
    for value in ["AB==", "AAB=", "AA=A", "====", "AA==\n"] { assert!(resource_byte_length(value).is_err(), "{value}"); }
}

#[test]
fn prefixed_relationship_roots_keep_comment_pi_and_external_targets() {
    let xml = format!("<?xml version='1.0'?><!--c--><p:Relationships xmlns:p='{RELATIONSHIP_NAMESPACE}'><?keep x?><p:Relationship Id='r1' Type='urn:resource' Target='a%20b.xml'/><p:Relationship Id='r2' Type='urn:external' Target='https://example.test/schema' TargetMode='External'/></p:Relationships>");
    let edges = relationships(&xml).unwrap(); assert_eq!(edges.len(), 2); assert!(!edges[0].external); assert!(edges[1].external);
    assert!(relationships(&format!("<Relationships xmlns='{RELATIONSHIP_NAMESPACE}'/>")).unwrap().is_empty());
    for body in ["text", "<Relationship Id='r1' Type='urn:x' Target='a'/><Relationship Id='r1' Type='urn:y' Target='b'/>",
        "<Relationship Id='r1' Type='urn:x' Target='a' TargetMode='Unknown'/>"] {
        assert!(relationships(&format!("<Relationships xmlns='{RELATIONSHIP_NAMESPACE}'>{body}</Relationships>")).is_err());
    }
}

#[test]
fn binding_mappings_accept_namespaces_and_reject_attributes_or_injected_nodes() {
    validate_prefix_mappings("xmlns:a='urn:example&amp;more' xmlns='urn:default'").unwrap();
    validate_prefix_mappings("").unwrap();
    for value in ["class='ordinary'", "xmlns:p=''", "xmlns:a='urn:x'/><other/>", "xmlns:a='urn:x'><child/></binding><binding"] {
        assert!(validate_prefix_mappings(value).is_err(), "{value}");
    }
}
