use std::borrow::Cow;

pub const DECL: &str = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";

pub fn escape(value: &str) -> Cow<'_, str> {
    if !value.chars().any(|character| {
        matches!(character, '&' | '<' | '>' | '\"' | '\'')
            || (character < '\u{20}' && !matches!(character, '\t' | '\n' | '\r'))
    }) {
        return Cow::Borrowed(value);
    }
    let mut output = String::with_capacity(value.len() + 16);
    for character in value.chars() {
        match character {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '>' => output.push_str("&gt;"),
            '\"' => output.push_str("&quot;"),
            '\'' => output.push_str("&apos;"),
            character
                if character == '\u{0}'
                    || (character < '\u{20}' && !matches!(character, '\t' | '\n' | '\r')) =>
            {
                output.push('\u{fffd}')
            }
            character => output.push(character),
        }
    }
    Cow::Owned(output)
}

pub fn attr(name: &str, value: &str) -> String {
    format!(" {name}=\"{}\"", escape(value))
}
