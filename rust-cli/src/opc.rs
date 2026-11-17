//! OPC resource graph copying. Relationship IDs belong to their owning part;
//! every imported source receives a separate package path namespace.
use crate::xml::{self, CT, Element, Namespaces, REL};
use anyhow::{Context, Result, bail, ensure};
use std::{
    collections::{BTreeMap, HashMap},
    fs::File,
    io::Read,
    path::{Component, Path, PathBuf},
};

#[derive(Clone, Debug)]
pub struct Relationship {
    pub id: String,
    pub kind: String,
    pub target: String,
    pub external: bool,
}

pub struct SourcePackage {
    pub tag: String,
    pub parts: BTreeMap<String, Vec<u8>>,
    types: BTreeMap<String, String>,
    defaults: BTreeMap<String, String>,
    source_file: Option<PathBuf>,
    hyperlink_base: Option<String>,
}
impl SourcePackage {
    pub fn open(path: &Path, tag: String) -> Result<Self> {
        let mut archive = zip::ZipArchive::new(File::open(path)?)?;
        let mut parts = BTreeMap::new();
        for i in 0..archive.len() {
            let mut entry = archive.by_index(i)?;
            if entry.is_dir() {
                continue;
            }
            let name = normalize("", entry.name())?;
            ensure!(!parts.contains_key(&name), "Duplicate OPC part {name}");
            let mut bytes = Vec::new();
            entry.read_to_end(&mut bytes)?;
            parts.insert(name, bytes);
        }
        let mut result = Self {
            tag,
            parts,
            types: BTreeMap::new(),
            defaults: BTreeMap::new(),
            source_file: Some(std::fs::canonicalize(path)?),
            hyperlink_base: None,
        };
        let content = result.xml("[Content_Types].xml", &mut Namespaces::new())?;
        for element in content.elements() {
            if element.name.local == "Override" {
                result.types.insert(
                    element
                        .attr("", "PartName")
                        .unwrap_or("")
                        .trim_start_matches('/')
                        .into(),
                    element.attr("", "ContentType").unwrap_or("").into(),
                );
            } else if element.name.local == "Default" {
                result.defaults.insert(
                    element
                        .attr("", "Extension")
                        .unwrap_or("")
                        .to_ascii_lowercase(),
                    element.attr("", "ContentType").unwrap_or("").into(),
                );
            }
        }
        if result.parts.contains_key("docProps/app.xml") {
            let properties = result.xml("docProps/app.xml", &mut Namespaces::new())?;
            result.hyperlink_base = properties
                .elements()
                .find(|e| e.name.local == "HyperlinkBase")
                .map(|e| {
                    e.children
                        .iter()
                        .filter_map(|n| {
                            if let xml::Node::Text(t) = n {
                                Some(t.as_str())
                            } else {
                                None
                            }
                        })
                        .collect::<String>()
                })
                .filter(|s| !s.is_empty());
        }
        Ok(result)
    }
    pub fn xml(&self, name: &str, ns: &mut Namespaces) -> Result<Element> {
        let bytes = self
            .parts
            .get(name)
            .with_context(|| format!("Missing OPC part {name}"))?;
        Element::parse(&xml::decode(bytes)?, ns)
    }
    pub fn relationships(&self, owner: &str) -> Result<Vec<Relationship>> {
        let rels = relationship_path(owner);
        if !self.parts.contains_key(&rels) {
            return Ok(Vec::new());
        }
        let root = self.xml(&rels, &mut Namespaces::new())?;
        let mut result = Vec::new();
        for rel in root.elements() {
            ensure!(
                rel.is(REL, "Relationship"),
                "Unexpected relationship element"
            );
            result.push(Relationship {
                id: rel.attr("", "Id").context("Relationship has no Id")?.into(),
                kind: rel
                    .attr("", "Type")
                    .context("Relationship has no Type")?
                    .into(),
                target: rel
                    .attr("", "Target")
                    .context("Relationship has no Target")?
                    .into(),
                external: rel.attr("", "TargetMode") == Some("External"),
            });
        }
        Ok(result)
    }
    fn content_type(&self, name: &str) -> Result<String> {
        if let Some(t) = self.types.get(name) {
            return Ok(t.clone());
        }
        let extension = name
            .rsplit_once('.')
            .map(|(_, x)| x.to_ascii_lowercase())
            .unwrap_or_default();
        self.defaults
            .get(&extension)
            .cloned()
            .with_context(|| format!("No content type for {name}"))
    }
    pub fn external_target(&self, target: &str) -> Result<String> {
        let Some(source) = &self.source_file else {
            return Ok(target.into());
        };
        if has_scheme(target)
            && !(target.len() > 2
                && target.as_bytes()[1] == b':'
                && matches!(target.as_bytes()[2], b'/' | b'\\'))
        {
            return Ok(target.into());
        }
        let split = target.find(['?', '#']).unwrap_or(target.len());
        let (path, suffix) = target.split_at(split);
        let base = self.hyperlink_base.as_deref();
        if let Some(base) =
            base.filter(|b| b.contains("://") && !b.to_ascii_lowercase().starts_with("file:"))
        {
            return join_web_uri(base, path, suffix);
        }
        let parent = source.parent().unwrap_or(Path::new("."));
        let base_path = if let Some(base) = base {
            let decoded = decode_file_uri(base)?;
            let candidate = PathBuf::from(decoded);
            if candidate.is_absolute() {
                candidate
            } else {
                parent.join(candidate)
            }
        } else {
            parent.to_path_buf()
        };
        let target_path = if path.is_empty() && base.is_none() {
            source.clone()
        } else {
            base_path.join(decode_percent(path)?)
        };
        Ok(format!(
            "{}{}",
            file_uri(&lexical_path(&target_path)),
            suffix
        ))
    }
}

#[derive(Default)]
pub struct Resources {
    pub parts: BTreeMap<String, Vec<u8>>,
    pub types: BTreeMap<String, String>,
    pub relationships: BTreeMap<String, Vec<Relationship>>,
    ids: HashMap<(String, String, String, String), String>,
}
impl Resources {
    pub fn import_relationship(
        &mut self,
        source: &SourcePackage,
        source_owner: &str,
        id: &str,
        destination_owner: &str,
    ) -> Result<String> {
        let key = (
            source.tag.clone(),
            source_owner.into(),
            id.into(),
            destination_owner.into(),
        );
        if let Some(value) = self.ids.get(&key) {
            return Ok(value.clone());
        }
        let relation = source
            .relationships(source_owner)?
            .into_iter()
            .find(|r| r.id == id)
            .with_context(|| format!("Unresolved relationship {source_owner}#{id}"))?;
        let mut copied = relation.clone();
        copied.id = format!("{}_r{}", source.tag, self.ids.len() + 1);
        if !copied.external {
            let (target, fragment) = split_fragment(&relation.target);
            let source_target = normalize(source_owner, target)?;
            let destination = self.copy_part(source, &source_target)?;
            copied.target = format!("/{}{}", encode_path(&destination), fragment);
        } else {
            copied.target = source.external_target(&copied.target)?;
        }
        let new_id = copied.id.clone();
        self.relationships
            .entry(destination_owner.into())
            .or_default()
            .push(copied);
        self.ids.insert(key, new_id.clone());
        Ok(new_id)
    }
    pub fn copy_part(&mut self, source: &SourcePackage, name: &str) -> Result<String> {
        let destination = format!("word/imports/{}/{}", source.tag, name);
        if self.parts.contains_key(&destination) {
            return Ok(destination);
        }
        let bytes = source
            .parts
            .get(name)
            .with_context(|| format!("Missing referenced resource {name}"))?;
        self.parts.insert(destination.clone(), bytes.clone());
        self.types
            .insert(destination.clone(), source.content_type(name)?);
        let mut copied = Vec::new();
        for mut relation in source.relationships(name)? {
            if !relation.external {
                let (target, fragment) = split_fragment(&relation.target);
                let child = normalize(name, target)?;
                let dest = self.copy_part(source, &child)?;
                relation.target = format!("/{}{}", encode_path(&dest), fragment);
            } else {
                relation.target = source.external_target(&relation.target)?;
            }
            copied.push(relation);
        }
        if !copied.is_empty() {
            self.relationships.insert(destination.clone(), copied);
        }
        Ok(destination)
    }
    pub fn add_relationship(&mut self, owner: &str, id: &str, kind: &str, target: &str) {
        self.relationships
            .entry(owner.into())
            .or_default()
            .push(Relationship {
                id: id.into(),
                kind: kind.into(),
                target: target.into(),
                external: false,
            });
    }
    pub fn add_external(&mut self, owner: &str, kind: &str, target: &str) -> String {
        let id = format!(
            "sourceLink{}",
            self.relationships.get(owner).map_or(0, Vec::len) + 1
        );
        self.relationships
            .entry(owner.into())
            .or_default()
            .push(Relationship {
                id: id.clone(),
                kind: kind.into(),
                target: target.into(),
                external: true,
            });
        id
    }
    pub fn finish_relationships(&mut self, ns: &mut Namespaces) {
        ns.register(REL, "pr");
        for (owner, rels) in &self.relationships {
            let mut root = Element::new(REL, "Relationships");
            for rel in rels {
                let mut element = Element::new(REL, "Relationship");
                element.set("", "Id", &rel.id);
                element.set("", "Type", &rel.kind);
                element.set("", "Target", &rel.target);
                if rel.external {
                    element.set("", "TargetMode", "External");
                }
                root.push(element);
            }
            self.parts.insert(
                relationship_path(owner),
                root.serialize(ns, true).into_bytes(),
            );
        }
    }
    pub fn content_types(&self, ns: &mut Namespaces) -> String {
        ns.register(CT, "ct");
        let mut root = Element::new(CT, "Types");
        let mut rel = Element::new(CT, "Default");
        rel.set("", "Extension", "rels");
        rel.set(
            "",
            "ContentType",
            "application/vnd.openxmlformats-package.relationships+xml",
        );
        root.push(rel);
        for (path, kind) in &self.types {
            let mut element = Element::new(CT, "Override");
            element.set("", "PartName", format!("/{path}"));
            element.set("", "ContentType", kind);
            root.push(element);
        }
        root.serialize(ns, true)
    }
}

fn split_fragment(target: &str) -> (&str, &str) {
    target
        .find('#')
        .map(|i| (&target[..i], &target[i..]))
        .unwrap_or((target, ""))
}
pub fn relationship_path(owner: &str) -> String {
    if owner.is_empty() {
        return "_rels/.rels".into();
    }
    match owner.rsplit_once('/') {
        Some((dir, name)) => format!("{dir}/_rels/{name}.rels"),
        None => format!("_rels/{owner}.rels"),
    }
}
pub fn normalize(owner: &str, target: &str) -> Result<String> {
    let mut decoded = Vec::new();
    let bytes = target.as_bytes();
    let mut i = 0;
    while i < bytes.len() {
        if bytes[i] == b'%' && i + 2 < bytes.len() {
            let value = u8::from_str_radix(std::str::from_utf8(&bytes[i + 1..i + 3])?, 16)
                .context("Invalid package URI escape")?;
            decoded.push(value);
            i += 3;
        } else {
            decoded.push(bytes[i]);
            i += 1;
        }
    }
    let target = String::from_utf8(decoded)?;
    ensure!(!target.contains('\\'), "Backslash in package URI");
    let mut segments: Vec<&str> = if target.starts_with('/') {
        Vec::new()
    } else {
        owner
            .rsplit_once('/')
            .map(|(dir, _)| dir.split('/').collect())
            .unwrap_or_default()
    };
    for part in target.split('/') {
        match part {
            "" | "." => {}
            ".." => {
                if segments.pop().is_none() {
                    bail!("Package relationship escapes root");
                }
            }
            p => segments.push(p),
        }
    }
    ensure!(!segments.is_empty(), "Empty package part URI");
    Ok(segments.join("/"))
}
fn encode_path(path: &str) -> String {
    path.bytes()
        .map(|b| {
            if b.is_ascii_alphanumeric() || b"/-_.~".contains(&b) {
                (b as char).to_string()
            } else {
                format!("%{b:02X}")
            }
        })
        .collect()
}

fn has_scheme(uri: &str) -> bool {
    uri.split_once(':').is_some_and(|(scheme, _)| {
        !scheme.is_empty()
            && scheme.as_bytes()[0].is_ascii_alphabetic()
            && scheme
                .bytes()
                .all(|b| b.is_ascii_alphanumeric() || b"+-.".contains(&b))
    })
}
fn decode_percent(text: &str) -> Result<String> {
    let mut result = Vec::new();
    let bytes = text.as_bytes();
    let mut i = 0;
    while i < bytes.len() {
        if bytes[i] == b'%' && i + 2 < bytes.len() {
            if let Ok(value) = u8::from_str_radix(std::str::from_utf8(&bytes[i + 1..i + 3])?, 16) {
                result.push(value);
                i += 3;
                continue;
            }
        }
        result.push(bytes[i]);
        i += 1;
    }
    String::from_utf8(result).context("Relative hyperlink is not a UTF-8 URI")
}
fn decode_file_uri(text: &str) -> Result<String> {
    if let Some(path) = text.strip_prefix("file:///") {
        return decode_percent(if cfg!(windows) { path } else { &text[7..] });
    }
    if let Some(path) = text.strip_prefix("file://") {
        return decode_percent(&format!("//{path}"));
    }
    decode_percent(text)
}
fn lexical_path(path: &Path) -> PathBuf {
    let mut result = PathBuf::new();
    for component in path.components() {
        match component {
            Component::CurDir => {}
            Component::ParentDir => {
                if result.file_name().is_some() {
                    result.pop();
                }
            }
            other => result.push(other.as_os_str()),
        }
    }
    result
}
pub fn file_uri(path: &Path) -> String {
    let raw = path.to_string_lossy();
    let text = raw
        .strip_prefix("\\\\?\\")
        .unwrap_or(&raw)
        .replace('\\', "/");
    let text = if let Some(unc) = text.strip_prefix("UNC/") {
        format!("//{unc}")
    } else {
        text
    };
    let encoded = text
        .bytes()
        .map(|b| {
            if b.is_ascii_alphanumeric() || b"/:-_.~".contains(&b) {
                (b as char).to_string()
            } else {
                format!("%{b:02X}")
            }
        })
        .collect::<String>();
    if encoded.starts_with("//") {
        format!("file:{encoded}")
    } else if encoded.starts_with('/') {
        format!("file://{encoded}")
    } else {
        format!("file:///{encoded}")
    }
}
fn join_web_uri(base: &str, path: &str, suffix: &str) -> Result<String> {
    if path.starts_with("//") {
        return Ok(format!(
            "{}:{path}{suffix}",
            base.split_once(':').unwrap().0
        ));
    }
    let (scheme, rest) = base
        .split_once("://")
        .context("Invalid hyperlink base URI")?;
    let (authority, base_path) = rest
        .split_once('/')
        .map(|(a, p)| (a, format!("/{p}")))
        .unwrap_or((rest, "/".into()));
    let combined = if path.starts_with('/') {
        path.to_owned()
    } else if path.is_empty() {
        base_path
    } else {
        format!(
            "{}{path}",
            base_path
                .rsplit_once('/')
                .map(|(dir, _)| format!("{dir}/"))
                .unwrap_or("/".into())
        )
    };
    let mut segments = Vec::new();
    for part in combined.split('/') {
        match part {
            "" | "." => {}
            ".." => {
                segments.pop();
            }
            p => segments.push(p),
        }
    }
    Ok(format!("{scheme}://{authority}/{}{suffix}", segments.join("/")).replace(' ', "%20"))
}

#[cfg(test)]
mod tests {
    use super::*;
    fn source(tag: &str, image: u8) -> SourcePackage {
        let mut parts = BTreeMap::new();
        parts.insert("word/media/same.png".into(), vec![image]);
        parts.insert("word/_rels/document.xml.rels".into(),format!("<Relationships xmlns=\"{REL}\"><Relationship Id=\"rId1\" Type=\"{}/image\" Target=\"media/same.png\"/><Relationship Id=\"rId2\" Type=\"{}/hyperlink\" Target=\"https://example.com/?a=1&amp;b=2\" TargetMode=\"External\"/></Relationships>",xml::R,xml::R).into_bytes());
        SourcePackage {
            tag: tag.into(),
            parts,
            types: BTreeMap::new(),
            defaults: BTreeMap::from([("png".into(), "image/png".into())]),
            source_file: None,
            hyperlink_base: None,
        }
    }
    #[test]
    fn same_names_and_ids_from_different_sources_are_isolated() {
        let mut resources = Resources::default();
        let first = source("s1", 1);
        let second = source("s2", 2);
        let a = resources
            .import_relationship(&first, "word/document.xml", "rId1", "word/document.xml")
            .unwrap();
        let b = resources
            .import_relationship(&second, "word/document.xml", "rId1", "word/document.xml")
            .unwrap();
        assert_ne!(a, b);
        assert_eq!(
            resources.parts["word/imports/s1/word/media/same.png"],
            vec![1]
        );
        assert_eq!(
            resources.parts["word/imports/s2/word/media/same.png"],
            vec![2]
        );
        resources
            .import_relationship(&first, "word/document.xml", "rId2", "word/document.xml")
            .unwrap();
        let link = resources.relationships["word/document.xml"].last().unwrap();
        assert!(link.external);
        assert_eq!(link.target, "https://example.com/?a=1&b=2");
    }
    #[test]
    fn package_paths_cannot_escape_root() {
        assert_eq!(
            normalize("word/charts/chart1.xml", "../media/image%201.png").unwrap(),
            "word/media/image 1.png"
        );
        assert!(normalize("word/document.xml", "../../outside").is_err());
    }
    #[test]
    fn recursive_resources_and_cycles_keep_relationship_scope() {
        let mut source = source("s1", 7);
        source
            .defaults
            .insert("xml".into(), "application/xml".into());
        for (part, target) in [("a", "b"), ("b", "a")] {
            source
                .parts
                .insert(format!("word/{part}.xml"), b"<root/>".to_vec());
            source.parts.insert(format!("word/_rels/{part}.xml.rels"),format!("<Relationships xmlns=\"{REL}\"><Relationship Id=\"same\" Type=\"urn:child\" Target=\"{target}.xml\"/><Relationship Id=\"image\" Type=\"urn:image\" Target=\"media/same.png\"/></Relationships>").into_bytes());
        }
        let mut copied = Resources::default();
        copied.copy_part(&source, "word/a.xml").unwrap();
        assert_eq!(copied.parts.len(), 3);
        for owner in ["word/imports/s1/word/a.xml", "word/imports/s1/word/b.xml"] {
            let relations = &copied.relationships[owner];
            assert_eq!(relations.len(), 2);
            for relation in relations {
                assert!(
                    copied
                        .parts
                        .contains_key(relation.target.trim_start_matches('/'))
                );
            }
        }
    }
    #[test]
    fn relative_hyperlinks_are_rebased_instead_of_following_the_output_folder() {
        let mut source = source("s1", 1);
        source.source_file = Some(
            std::env::temp_dir()
                .join("source-folder")
                .join("source.docx"),
        );
        assert!(
            source
                .external_target("target%20file.docx#Bookmark")
                .unwrap()
                .ends_with("/source-folder/target%20file.docx#Bookmark")
        );
        source.hyperlink_base = Some("https://example.test/base/sub/".into());
        assert_eq!(
            source.external_target("../target.docx?x=1").unwrap(),
            "https://example.test/base/target.docx?x=1"
        );
    }
}
