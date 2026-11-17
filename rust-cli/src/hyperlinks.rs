//! Bookmark targets can live in field instructions as well as w:hyperlink.
use crate::xml::{Element, Node, W};
use anyhow::Result;
use std::collections::HashMap;

fn token_spans(code: &str) -> Vec<(usize, usize)> {
    let bytes = code.as_bytes();
    let mut tokens = Vec::new();
    let mut i = 0;
    while i < bytes.len() {
        while i < bytes.len() && bytes[i].is_ascii_whitespace() {
            i += 1;
        }
        if i == bytes.len() {
            break;
        }
        let start = i;
        if bytes[i] == b'"' {
            i += 1;
            while i < bytes.len() {
                if bytes[i] == b'"' && bytes[i - 1] != b'\\' {
                    i += 1;
                    break;
                }
                i += 1;
            }
        } else {
            while i < bytes.len() && !bytes[i].is_ascii_whitespace() {
                i += 1;
            }
        }
        tokens.push((start, i));
    }
    tokens
}

pub fn rewrite_instruction(code: &str, names: &HashMap<String, String>) -> String {
    let tokens = token_spans(code);
    let token = |index: usize| {
        tokens
            .get(index)
            .map(|(a, b)| code[*a..*b].trim_matches('"'))
    };
    let Some(command) = token(0) else {
        return code.into();
    };
    let mut hash_prefix = false;
    let selected = if ["REF", "PAGEREF", "NOTEREF"]
        .iter()
        .any(|name| command.eq_ignore_ascii_case(name))
    {
        Some(1)
    } else if command.eq_ignore_ascii_case("HYPERLINK") {
        if let Some(address) = token(1).filter(|t| !t.starts_with('\\') && !t.is_empty()) {
            if address.starts_with('#') {
                hash_prefix = true;
                Some(1)
            } else {
                None
            }
        } else {
            (1..tokens.len())
                .find(|index| token(*index).is_some_and(|t| t.eq_ignore_ascii_case("\\l")))
                .map(|index| index + 1)
        }
    } else {
        None
    };
    let Some(index) = selected else {
        return code.into();
    };
    let Some(target) = token(index) else {
        return code.into();
    };
    let key = if hash_prefix {
        target.trim_start_matches('#')
    } else {
        target
    };
    let Some(mapped) = names.get(key) else {
        return code.into();
    };
    let (start, end) = tokens[index];
    let quoted = code.as_bytes()[start] == b'"';
    let prefix = if hash_prefix { "#" } else { "" };
    let value = if quoted {
        format!("\"{prefix}{mapped}\"")
    } else {
        format!("{prefix}{mapped}")
    };
    format!("{}{}{}", &code[..start], value, &code[end..])
}

pub fn externalize_instruction(code: &str, targets: &HashMap<String, (String, String)>) -> String {
    let spans = token_spans(code);
    let token = |index: usize| {
        spans
            .get(index)
            .map(|(a, b)| code[*a..*b].trim_matches('"'))
    };
    if !token(0).is_some_and(|t| t.eq_ignore_ascii_case("HYPERLINK")) {
        return code.into();
    }
    if let Some(first) = token(1).filter(|s| !s.is_empty() && !s.starts_with('\\')) {
        if let Some(target) = first.strip_prefix('#').and_then(|name| targets.get(name)) {
            let (a, b) = spans[1];
            return format!("{}\"{}#{}\"{}", &code[..a], target.0, target.1, &code[b..]);
        }
        return code.into();
    }
    let Some(index) = (1..spans.len())
        .find(|index| token(*index).is_some_and(|s| s.eq_ignore_ascii_case("\\l")))
        .map(|index| index + 1)
    else {
        return code.into();
    };
    let Some((uri, original)) = token(index).and_then(|name| targets.get(name)) else {
        return code.into();
    };
    let (a, b) = spans[index];
    let mut changed = format!("{}\"{}\"{}", &code[..a], original, &code[b..]);
    if token(1) == Some("") {
        let (a, b) = spans[1];
        changed.replace_range(a..b, &format!("\"{uri}\""));
    } else {
        changed.insert_str(spans[0].1, &format!(" \"{uri}\""));
    }
    changed
}

#[derive(Default)]
struct Field {
    paths: Vec<Vec<usize>>,
    code: String,
    active: bool,
}
pub fn rebase_external_instruction(
    code: &str,
    resolve: &mut impl FnMut(&str) -> Result<String>,
) -> Result<String> {
    let spans = token_spans(code);
    let token = |index: usize| {
        spans
            .get(index)
            .map(|(a, b)| code[*a..*b].trim_matches('"'))
    };
    if !token(0).is_some_and(|t| t.eq_ignore_ascii_case("HYPERLINK")) {
        return Ok(code.into());
    }
    let Some(address) = token(1).filter(|t| !t.is_empty() && !t.starts_with(['\\', '#'])) else {
        return Ok(code.into());
    };
    let resolved = resolve(address)?;
    if resolved == address {
        return Ok(code.into());
    }
    let (start, end) = spans[1];
    Ok(format!("{}\"{resolved}\"{}", &code[..start], &code[end..]))
}

pub fn rewrite_fields(
    root: &mut Element,
    names: &HashMap<String, String>,
    mut resolve: impl FnMut(&str) -> Result<String>,
) -> Result<()> {
    let mut stack = Vec::new();
    let mut actions = Vec::new();
    collect(root, &mut Vec::new(), &mut stack, &mut actions);
    for (paths, original) in actions {
        let changed =
            rebase_external_instruction(&rewrite_instruction(&original, names), &mut resolve)?;
        if changed == original {
            continue;
        }
        for (index, path) in paths.iter().enumerate() {
            if let Some(element) = at_mut(root, path) {
                element.children = if index == 0 {
                    vec![Node::Text(changed.clone())]
                } else {
                    Vec::new()
                };
            }
        }
    }
    Ok(())
}
fn collect(
    element: &Element,
    path: &mut Vec<usize>,
    stack: &mut Vec<Field>,
    actions: &mut Vec<(Vec<Vec<usize>>, String)>,
) {
    if element.is(W, "fldChar") {
        match element.attr(W, "fldCharType") {
            Some("begin") => stack.push(Field {
                active: true,
                ..Field::default()
            }),
            Some("separate") => {
                if let Some(field) = stack.last_mut() {
                    if field.active {
                        actions.push((field.paths.clone(), field.code.clone()));
                        field.active = false;
                    }
                }
            }
            Some("end") => {
                if let Some(field) = stack.pop() {
                    if field.active {
                        actions.push((field.paths, field.code));
                    }
                }
            }
            _ => {}
        }
    }
    if element.is(W, "instrText") {
        if let Some(field) = stack.last_mut().filter(|f| f.active) {
            field.paths.push(path.clone());
            for child in &element.children {
                if let Node::Text(text) = child {
                    field.code.push_str(text);
                }
            }
        }
    }
    for (index, child) in element.children.iter().enumerate() {
        if let Node::Element(child) = child {
            path.push(index);
            collect(child, path, stack, actions);
            path.pop();
        }
    }
}
fn at_mut<'a>(element: &'a mut Element, path: &[usize]) -> Option<&'a mut Element> {
    if path.is_empty() {
        return Some(element);
    }
    match element.children.get_mut(path[0])? {
        Node::Element(child) => at_mut(child, &path[1..]),
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn local_targets_are_remapped_but_other_documents_are_not() {
        let names = HashMap::from([("Target".into(), "s1_Target".into())]);
        assert_eq!(
            rewrite_instruction(" HYPERLINK \\l \"Target\" \\o \"tip\" ", &names),
            " HYPERLINK \\l \"s1_Target\" \\o \"tip\" "
        );
        assert_eq!(
            rewrite_instruction(" REF Target \\h ", &names),
            " REF s1_Target \\h "
        );
        assert_eq!(
            rewrite_instruction(" HYPERLINK \"other.docx\" \\l Target ", &names),
            " HYPERLINK \"other.docx\" \\l Target "
        );
    }
    #[test]
    fn split_instructions_keep_independent_fields_separate() {
        let mut ns = crate::xml::Namespaces::new();
        let mut root=Element::parse(&format!("<w:p xmlns:w=\"{W}\"><w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText> HYPERLINK \\l \"Tar</w:instrText></w:r><w:r><w:instrText>get\" </w:instrText></w:r><w:r><w:fldChar w:fldCharType=\"separate\"/></w:r><w:r><w:t>visible</w:t></w:r><w:r><w:fldChar w:fldCharType=\"end\"/></w:r></w:p>"),&mut ns).unwrap();
        rewrite_fields(
            &mut root,
            &HashMap::from([("Target".into(), "s1_Target".into())]),
            |target| Ok(target.into()),
        )
        .unwrap();
        let result = root.serialize(&ns, true);
        assert!(result.contains("s1_Target"));
        assert!(result.contains("visible"));
    }
    #[test]
    fn omitted_bookmarks_keep_links_to_the_original_document() {
        let targets = HashMap::from([(
            "s1_Target".into(),
            ("file:///source.docx".into(), "Target".into()),
        )]);
        assert_eq!(
            externalize_instruction(" HYPERLINK \\l \"s1_Target\" \\o \"tip\" ", &targets),
            " HYPERLINK \"file:///source.docx\" \\l \"Target\" \\o \"tip\" "
        );
        assert_eq!(
            externalize_instruction(" HYPERLINK \"\" \\l s1_Target ", &targets),
            " HYPERLINK \"file:///source.docx\" \\l \"Target\" "
        );
    }
}
