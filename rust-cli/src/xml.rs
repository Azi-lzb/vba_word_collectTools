//! Namespace-aware XML editing for package imports. Names are keyed by URI,
//! so identically spelled prefixes from different source packages cannot collide.
use anyhow::{Context, Result};
use std::collections::BTreeMap;

pub const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
pub const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
pub const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
pub const REL: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
pub const CT: &str = "http://schemas.openxmlformats.org/package/2006/content-types";
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Name {
    pub uri: String,
    pub local: String,
}
impl Name {
    pub fn new(uri: &str, local: &str) -> Self {
        Self {
            uri: uri.into(),
            local: local.into(),
        }
    }
}

#[derive(Clone, Debug)]
pub enum Node {
    Element(Element),
    Text(String),
}

#[derive(Clone, Debug)]
pub struct Element {
    pub name: Name,
    pub attrs: Vec<(Name, String)>,
    pub children: Vec<Node>,
}

#[derive(Default)]
pub struct Namespaces {
    pub prefixes: BTreeMap<String, String>,
}
impl Namespaces {
    pub fn new() -> Self {
        let mut result = Self::default();
        for (prefix, uri) in [
            ("w", W),
            ("r", R),
            ("a", A),
            ("mc", MC),
            ("xml", "http://www.w3.org/XML/1998/namespace"),
        ] {
            result.prefixes.insert(uri.into(), prefix.into());
        }
        result
    }
    pub fn register(&mut self, uri: &str, preferred: &str) -> String {
        if uri.is_empty() {
            return String::new();
        }
        if let Some(prefix) = self.prefixes.get(uri) {
            return prefix.clone();
        }
        let mut prefix = preferred.to_owned();
        if prefix.is_empty() || self.prefixes.values().any(|p| p == &prefix) {
            prefix = format!("ns{}", self.prefixes.len());
        }
        self.prefixes.insert(uri.into(), prefix.clone());
        prefix
    }
    fn qualified(&self, name: &Name) -> String {
        if name.uri.is_empty() {
            return name.local.clone();
        }
        format!("{}:{}", self.prefixes[&name.uri], name.local)
    }
    pub fn declarations(&self) -> String {
        self.prefixes
            .iter()
            .filter(|(_, prefix)| prefix.as_str() != "xml")
            .map(|(uri, prefix)| format!(" xmlns:{prefix}=\"{}\"", escape(uri)))
            .collect()
    }
}

impl Element {
    pub fn new(uri: &str, local: &str) -> Self {
        Self {
            name: Name::new(uri, local),
            attrs: Vec::new(),
            children: Vec::new(),
        }
    }
    pub fn is(&self, uri: &str, local: &str) -> bool {
        self.name.uri == uri && self.name.local == local
    }
    pub fn attr(&self, uri: &str, local: &str) -> Option<&str> {
        self.attrs
            .iter()
            .find(|(n, _)| n.uri == uri && n.local == local)
            .map(|(_, v)| v.as_str())
    }
    pub fn set(&mut self, uri: &str, local: &str, value: impl Into<String>) {
        let value = value.into();
        if let Some((_, v)) = self
            .attrs
            .iter_mut()
            .find(|(n, _)| n.uri == uri && n.local == local)
        {
            *v = value;
        } else {
            self.attrs.push((Name::new(uri, local), value));
        }
    }
    pub fn remove_attr(&mut self, uri: &str, local: &str) {
        self.attrs.retain(|(n, _)| n.uri != uri || n.local != local);
    }
    pub fn elements(&self) -> impl Iterator<Item = &Element> {
        self.children.iter().filter_map(|n| {
            if let Node::Element(e) = n {
                Some(e)
            } else {
                None
            }
        })
    }
    pub fn child(&self, uri: &str, local: &str) -> Option<&Element> {
        self.elements().find(|e| e.is(uri, local))
    }
    pub fn push(&mut self, element: Element) {
        self.children.push(Node::Element(element));
    }
    pub fn parse(xml: &str, namespaces: &mut Namespaces) -> Result<Self> {
        let doc = roxmltree::Document::parse(xml).context("Invalid XML package part")?;
        Ok(Self::from_node(doc.root_element(), namespaces))
    }
    fn from_node(node: roxmltree::Node<'_, '_>, ns: &mut Namespaces) -> Self {
        let uri = node.tag_name().namespace().unwrap_or("");
        ns.register(uri, node.lookup_prefix(uri).unwrap_or(""));
        let mut result = Self::new(uri, node.tag_name().name());
        for attribute in node.attributes() {
            let uri = attribute.namespace().unwrap_or("");
            ns.register(uri, node.lookup_prefix(uri).unwrap_or(""));
            let mut value = attribute.value().to_owned();
            if (uri == MC && attribute.name() == "Ignorable")
                || (result.name.uri == MC && attribute.name() == "Requires")
            {
                value = attribute
                    .value()
                    .split_whitespace()
                    .map(|prefix| {
                        node.lookup_namespace_uri(Some(prefix))
                            .map(|uri| ns.register(uri, prefix))
                            .unwrap_or_else(|| prefix.into())
                    })
                    .collect::<Vec<_>>()
                    .join(" ");
            }
            result.attrs.push((Name::new(uri, attribute.name()), value));
        }
        for child in node.children() {
            if child.is_element() {
                result.push(Self::from_node(child, ns));
            } else if child.is_text() {
                result
                    .children
                    .push(Node::Text(child.text().unwrap_or("").into()));
            }
        }
        result
    }
    pub fn serialize(&self, ns: &Namespaces, root: bool) -> String {
        let name = ns.qualified(&self.name);
        let mut xml = format!("<{name}");
        if root {
            xml.push_str(&ns.declarations());
        }
        for (name, value) in &self.attrs {
            xml.push_str(&format!(" {}=\"{}\"", ns.qualified(name), escape(value)));
        }
        if self.children.is_empty() {
            xml.push_str("/>");
            return xml;
        }
        xml.push('>');
        for node in &self.children {
            match node {
                Node::Element(e) => xml.push_str(&e.serialize(ns, false)),
                Node::Text(t) => xml.push_str(&escape(t)),
            }
        }
        xml.push_str(&format!("</{name}>"));
        xml
    }
}

pub fn escape(text: &str) -> String {
    text.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
}

pub fn decode(bytes: &[u8]) -> Result<String> {
    if bytes.starts_with(&[0xff, 0xfe]) || bytes.starts_with(&[0xfe, 0xff]) {
        let little = bytes[0] == 0xff;
        anyhow::ensure!(bytes.len() % 2 == 0, "Odd UTF-16 XML byte length");
        let units = bytes[2..]
            .chunks_exact(2)
            .map(|p| {
                if little {
                    u16::from_le_bytes([p[0], p[1]])
                } else {
                    u16::from_be_bytes([p[0], p[1]])
                }
            })
            .collect::<Vec<_>>();
        return String::from_utf16(&units).context("Invalid UTF-16 XML");
    }
    Ok(std::str::from_utf8(bytes)
        .context("Invalid UTF-8 XML")?
        .trim_start_matches('\u{feff}')
        .to_owned())
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn source_prefix_collisions_keep_namespace_identity() {
        let mut namespaces = Namespaces::new();
        let a = Element::parse(
            "<x:root xmlns:x=\"urn:one\"><x:item x:value=\"&amp;\"/></x:root>",
            &mut namespaces,
        )
        .unwrap();
        let b = Element::parse(
            "<x:root xmlns:x=\"urn:two\"><x:item/></x:root>",
            &mut namespaces,
        )
        .unwrap();
        for (element, uri) in [(a, "urn:one"), (b, "urn:two")] {
            let serialized = element.serialize(&namespaces, true);
            let parsed = roxmltree::Document::parse(&serialized).unwrap();
            assert_eq!(parsed.root_element().tag_name().namespace(), Some(uri));
        }
    }
    #[test]
    fn utf16_xml_preserves_non_ascii_text() {
        let text = "<root>中文</root>";
        let mut bytes = vec![0xff, 0xfe];
        for unit in text.encode_utf16() {
            bytes.extend(unit.to_le_bytes());
        }
        assert_eq!(decode(&bytes).unwrap(), text);
    }
}
