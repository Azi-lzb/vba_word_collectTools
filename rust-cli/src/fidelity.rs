//! Import document fragments without sharing source-local IDs or resource paths.
use crate::{
    model::Report,
    opc::{Resources, SourcePackage, normalize},
    xml::{self, A, Element, Namespaces, Node, R, W},
};
use anyhow::{Context as _, Result};
use std::{
    collections::{BTreeMap, HashMap, HashSet},
    fs::File,
    io::{Read, Write},
    path::Path,
};

#[derive(Default)]
struct Theme {
    colors: HashMap<String, String>,
    fonts: HashMap<String, String>,
}
impl Theme {
    fn load(package: &SourcePackage, ns: &mut Namespaces) -> Result<Self> {
        let mut result = Self::default();
        let relation = package
            .relationships("word/document.xml")?
            .into_iter()
            .find(|r| r.kind.ends_with("/theme"));
        let Some(relation) = relation else {
            return Ok(result);
        };
        let theme = package.xml(&normalize("word/document.xml", &relation.target)?, ns)?;
        let settings = package
            .relationships("word/document.xml")?
            .into_iter()
            .find(|r| r.kind.ends_with("/settings"))
            .map(|r| {
                normalize("word/document.xml", &r.target).and_then(|path| package.xml(&path, ns))
            })
            .transpose()?;
        let language = settings
            .as_ref()
            .and_then(|s| s.child(W, "themeFontLang"))
            .and_then(|e| e.attr(W, "eastAsia"))
            .unwrap_or("zh-CN");
        let script = if language.starts_with("ja") {
            "Jpan"
        } else if language.starts_with("ko") {
            "Hang"
        } else if language.starts_with("zh-TW") || language.starts_with("zh-HK") {
            "Hant"
        } else {
            "Hans"
        };
        if let Some(elements) = theme.child(A, "themeElements") {
            if let Some(scheme) = elements.child(A, "clrScheme") {
                for color in scheme.elements() {
                    if let Some(value) = color
                        .elements()
                        .next()
                        .and_then(|e| e.attr("", "lastClr").or_else(|| e.attr("", "val")))
                    {
                        result.colors.insert(color.name.local.clone(), value.into());
                    }
                }
            }
            if let Some(scheme) = elements.child(A, "fontScheme") {
                for family in ["major", "minor"] {
                    if let Some(font) = scheme.child(A, &format!("{family}Font")) {
                        let latin = font
                            .child(A, "latin")
                            .and_then(|e| e.attr("", "typeface"))
                            .unwrap_or("Arial");
                        let east = font
                            .child(A, "ea")
                            .and_then(|e| e.attr("", "typeface"))
                            .filter(|s| !s.is_empty())
                            .or_else(|| {
                                font.elements()
                                    .find(|e| e.attr("", "script") == Some(script))
                                    .and_then(|e| e.attr("", "typeface"))
                            })
                            .unwrap_or(latin);
                        let cs = font
                            .child(A, "cs")
                            .and_then(|e| e.attr("", "typeface"))
                            .filter(|s| !s.is_empty())
                            .unwrap_or(latin);
                        for (kind, value) in [
                            ("Ascii", latin),
                            ("HAnsi", latin),
                            ("EastAsia", east),
                            ("Bidi", cs),
                        ] {
                            result.fonts.insert(format!("{family}{kind}"), value.into());
                        }
                    }
                }
            }
        }
        if let Some(mapping) = settings
            .as_ref()
            .and_then(|s| s.child(W, "clrSchemeMapping"))
        {
            let original = Theme {
                colors: result.colors.clone(),
                fonts: HashMap::new(),
            };
            for (attribute, theme_name) in [
                ("bg1", "background1"),
                ("bg2", "background2"),
                ("t1", "text1"),
                ("t2", "text2"),
                ("accent1", "accent1"),
                ("accent2", "accent2"),
                ("accent3", "accent3"),
                ("accent4", "accent4"),
                ("accent5", "accent5"),
                ("accent6", "accent6"),
                ("hyperlink", "hyperlink"),
                ("followedHyperlink", "followedHyperlink"),
            ] {
                if let Some(value) = mapping
                    .attr(W, attribute)
                    .and_then(|name| original.color(name))
                {
                    result.colors.insert(theme_name.into(), value);
                }
            }
        }
        Ok(result)
    }
    fn color(&self, name: &str) -> Option<String> {
        let key = match name {
            "text1" | "dark1" => "dk1",
            "text2" | "dark2" => "dk2",
            "background1" | "light1" => "lt1",
            "background2" | "light2" => "lt2",
            "hyperlink" => "hlink",
            "followedHyperlink" => "folHlink",
            x => x,
        };
        self.colors
            .get(name)
            .or_else(|| self.colors.get(key))
            .cloned()
    }
    fn materialize(&self, element: &mut Element) {
        if element.name.uri == W {
            for (theme, target, tint, shade) in [
                ("themeColor", "color", "themeTint", "themeShade"),
                ("themeFill", "fill", "themeFillTint", "themeFillShade"),
            ] {
                if let Some(value) = element.attr(W, theme).and_then(|n| self.color(n)) {
                    let color = adjust_color(&value, element.attr(W, tint), element.attr(W, shade));
                    let target = if element.is(W, "color") && target == "color" {
                        "val"
                    } else {
                        target
                    };
                    element.set(W, target, color);
                    for attribute in [theme, tint, shade] {
                        element.remove_attr(W, attribute);
                    }
                }
            }
            if element.is(W, "rFonts") {
                for (theme, target) in [
                    ("asciiTheme", "ascii"),
                    ("hAnsiTheme", "hAnsi"),
                    ("eastAsiaTheme", "eastAsia"),
                    ("cstheme", "cs"),
                    ("csTheme", "cs"),
                ] {
                    if let Some(font) = element
                        .attr(W, theme)
                        .and_then(|v| self.fonts.get(v))
                        .cloned()
                    {
                        element.set(W, target, font);
                        element.remove_attr(W, theme);
                    }
                }
            }
        }
        if element.is(A, "schemeClr") {
            if let Some(color) = element.attr("", "val").and_then(|n| self.color(n)) {
                element.name.local = "srgbClr".into();
                element.set("", "val", color);
            }
        }
    }
}

fn adjust_color(value: &str, tint: Option<&str>, shade: Option<&str>) -> String {
    if value.len() != 6 {
        return value.into();
    }
    if tint.is_none() && shade.is_none() {
        return value.to_ascii_uppercase();
    }
    // WordprocessingML tint/shade modify HSL luminance, not RGB channels.
    let rgb = (0..3)
        .map(|i| u8::from_str_radix(&value[i * 2..i * 2 + 2], 16).unwrap_or(0) as f64 / 255.0)
        .collect::<Vec<_>>();
    let high = rgb.iter().copied().fold(0.0, f64::max);
    let low = rgb.iter().copied().fold(1.0, f64::min);
    let delta = high - low;
    let mut luminance = (high + low) / 2.0;
    let saturation = if delta == 0.0 {
        0.0
    } else {
        delta
            / if luminance <= 0.5 {
                high + low
            } else {
                2.0 - high - low
            }
    };
    let hue = if delta == 0.0 {
        0.0
    } else if high == rgb[0] {
        ((rgb[1] - rgb[2]) / delta).rem_euclid(6.0) / 6.0
    } else if high == rgb[1] {
        ((rgb[2] - rgb[0]) / delta + 2.0) / 6.0
    } else {
        ((rgb[0] - rgb[1]) / delta + 4.0) / 6.0
    };
    if let Some(shade) = shade.and_then(|v| u8::from_str_radix(v, 16).ok()) {
        luminance *= shade as f64 / 255.0;
    }
    if let Some(tint) = tint.and_then(|v| u8::from_str_radix(v, 16).ok()) {
        luminance = 1.0 - (1.0 - luminance) * tint as f64 / 255.0;
    }
    let q = if luminance < 0.5 {
        luminance * (1.0 + saturation)
    } else {
        luminance + saturation - luminance * saturation
    };
    let p = 2.0 * luminance - q;
    [hue + 1.0 / 3.0, hue, hue - 1.0 / 3.0]
        .iter()
        .map(|h| {
            let t = h.rem_euclid(1.0);
            let channel = if saturation == 0.0 {
                luminance
            } else if t < 1.0 / 6.0 {
                p + (q - p) * 6.0 * t
            } else if t < 0.5 {
                q
            } else if t < 2.0 / 3.0 {
                p + (q - p) * (2.0 / 3.0 - t) * 6.0
            } else {
                p
            };
            format!(
                "{:02X}",
                (channel * 255.0 + 1e-7).floor().clamp(0.0, 255.0) as u8
            )
        })
        .collect()
}

struct Context {
    package: SourcePackage,
    declarations: String,
    default_style: String,
    default_table_style: Option<String>,
    bookmark_names: HashMap<String, String>,
    vml_names: HashMap<String, String>,
    nums: HashMap<String, String>,
    abstracts: HashMap<String, String>,
    pictures: HashMap<String, String>,
    theme: Theme,
}

pub struct Prepared {
    pub report: Report,
    ns: Namespaces,
    resources: Resources,
    styles: Element,
    numbering: Element,
    fonts: Element,
    next_id: u64,
    original_targets: HashMap<String, (String, String)>,
    ignorable_uris: HashSet<String>,
}

impl Prepared {
    pub fn ignorable_prefixes(&self) -> String {
        let mut prefixes = self
            .ignorable_uris
            .iter()
            .filter_map(|uri| self.ns.prefixes.get(uri).cloned())
            .collect::<Vec<_>>();
        prefixes.sort();
        prefixes.join(" ")
    }
    pub fn sanitize_body(&mut self, body: &str) -> Result<String> {
        let mut root = Element::parse(
            &format!("<fragment{}>{body}</fragment>", self.ns.declarations()),
            &mut self.ns,
        )?;
        let mut state = BodyIds::default();
        sanitize_ids(&mut root, &mut state);
        remove_unpaired_bookmarks(&mut root, &state.paired);
        let names = collect_bookmark_names(&root);
        repair_links(
            &mut root,
            &names,
            &self.original_targets,
            &mut self.resources,
        )?;
        Ok(root
            .elements()
            .map(|e| e.serialize(&self.ns, false))
            .collect())
    }
    pub fn new(report: &Report) -> Result<Self> {
        let mut ns = Namespaces::new();
        ns.register(xml::CT, "ct");
        ns.register(xml::REL, "pr");
        let mut result = Self {
            report: report.clone(),
            ns,
            resources: Resources::default(),
            styles: Element::new(W, "styles"),
            numbering: Element::new(W, "numbering"),
            fonts: Element::new(W, "fonts"),
            next_id: 1,
            original_targets: HashMap::new(),
            ignorable_uris: HashSet::new(),
        };
        let mut normal = Element::new(W, "style");
        normal.set(W, "type", "paragraph");
        normal.set(W, "default", "1");
        normal.set(W, "styleId", "Normal");
        let mut name = Element::new(W, "name");
        name.set(W, "val", "Normal");
        normal.push(name);
        result.styles.push(normal);
        let mut contexts = HashMap::new();
        for (index, source) in report.sources.iter().enumerate() {
            let context = result.source(Path::new(&source.file), format!("s{}", index + 1))?;
            let uri = source_uri(Path::new(&source.file))?;
            for (original, mapped) in &context.bookmark_names {
                result
                    .original_targets
                    .insert(mapped.clone(), (uri.clone(), original.clone()));
            }
            contexts.insert(source.file.clone(), context);
        }
        for record in &mut result.report.records {
            // Transform after all namespaces and style IDs have been allocated.
            let context = &contexts[&record.source_file];
            for fragment in [&mut record.question_xml, &mut record.answer_xml] {
                if fragment.is_empty() {
                    continue;
                }
                let mut root = Element::parse(
                    &format!("<fragment{}>{}</fragment>", context.declarations, fragment),
                    &mut result.ns,
                )?;
                crate::hyperlinks::rewrite_fields(&mut root, &context.bookmark_names, |address| {
                    context.package.external_target(address)
                })?;
                transform(
                    &mut root,
                    context,
                    &mut result.resources,
                    &mut result.next_id,
                    "word/document.xml",
                    "word/document.xml",
                    true,
                )?;
                *fragment = root
                    .elements()
                    .map(|e| e.serialize(&result.ns, false))
                    .collect();
            }
        }
        result.report.namespaces = result
            .ns
            .prefixes
            .iter()
            .map(|(uri, prefix)| (prefix.clone(), uri.clone()))
            .collect();
        Ok(result)
    }

    fn source(&mut self, path: &Path, tag: String) -> Result<Context> {
        let package = SourcePackage::open(path, tag)?;
        let document = xml::decode(
            package
                .parts
                .get("word/document.xml")
                .context("Missing main document")?,
        )?;
        let parsed = roxmltree::Document::parse(&document)?;
        if let Some(prefixes) = parsed.root_element().attribute((
            "http://schemas.openxmlformats.org/markup-compatibility/2006",
            "Ignorable",
        )) {
            for prefix in prefixes.split_whitespace() {
                if let Some(uri) = parsed.root_element().lookup_namespace_uri(Some(prefix)) {
                    self.ns.register(uri, prefix);
                    self.ignorable_uris.insert(uri.into());
                }
            }
        }
        let mut bookmark_names = HashMap::new();
        let mut used_names = HashSet::new();
        let mut vml_names = HashMap::new();
        for node in parsed.descendants().filter(|n| n.is_element()) {
            if node.tag_name().namespace() == Some(W) && node.tag_name().name() == "bookmarkStart" {
                if let Some(name) = node.attribute((W, "name")) {
                    if !bookmark_names.contains_key(name) {
                        let original = format!("{}_{}", package.tag, name);
                        let mut candidate = original.chars().take(40).collect::<String>();
                        let mut counter = 1;
                        while used_names.contains(&candidate) {
                            let suffix = format!("_{counter}");
                            candidate = format!(
                                "{}{}",
                                original.chars().take(40 - suffix.len()).collect::<String>(),
                                suffix
                            );
                            counter += 1;
                        }
                        used_names.insert(candidate.clone());
                        bookmark_names.insert(name.into(), candidate);
                    }
                }
            }
            if node.tag_name().namespace() == Some("urn:schemas-microsoft-com:vml") {
                if let Some(id) = node.attribute("id") {
                    vml_names.insert(id.into(), format!("{}_{}", package.tag, id));
                }
            }
        }
        let declarations = parsed
            .root_element()
            .namespaces()
            .filter(|n| n.name() != Some("xml"))
            .map(|n| match n.name() {
                Some(prefix) => format!(" xmlns:{prefix}=\"{}\"", xml::escape(n.uri())),
                None => format!(" xmlns=\"{}\"", xml::escape(n.uri())),
            })
            .collect();
        let theme = Theme::load(&package, &mut self.ns)?;
        let mut styles = if package.parts.contains_key("word/styles.xml") {
            package.xml("word/styles.xml", &mut self.ns)?
        } else {
            Element::new(W, "styles")
        };
        let default = styles
            .elements()
            .find(|e| {
                e.is(W, "style")
                    && e.attr(W, "type") == Some("paragraph")
                    && e.attr(W, "default") == Some("1")
            })
            .and_then(|e| e.attr(W, "styleId"))
            .unwrap_or("Normal")
            .to_owned();
        let default_table_style = styles
            .elements()
            .find(|e| {
                e.is(W, "style")
                    && e.attr(W, "type") == Some("table")
                    && e.attr(W, "default") == Some("1")
            })
            .and_then(|e| e.attr(W, "styleId"))
            .map(|id| format!("{}_{}", package.tag, id));
        let mut defaults_id = "__ImportDefaults".to_owned();
        while styles
            .elements()
            .any(|e| e.attr(W, "styleId") == Some(defaults_id.as_str()))
        {
            defaults_id.push('_');
        }
        let mut context = Context {
            declarations,
            default_style: format!("{}_{}", package.tag, default),
            default_table_style,
            bookmark_names,
            vml_names,
            package,
            theme,
            nums: HashMap::new(),
            abstracts: HashMap::new(),
            pictures: HashMap::new(),
        };
        if context.package.parts.contains_key("word/fontTable.xml") {
            let fonts = context.package.xml("word/fontTable.xml", &mut self.ns)?;
            for mut font in fonts.elements().cloned() {
                let name = font.attr(W, "name").unwrap_or("");
                if self
                    .fonts
                    .elements()
                    .any(|old| old.attr(W, "name") == Some(name))
                {
                    continue;
                }
                transform(
                    &mut font,
                    &context,
                    &mut self.resources,
                    &mut self.next_id,
                    "word/fontTable.xml",
                    "word/fontTable.xml",
                    false,
                )?;
                self.fonts.push(font);
            }
        }
        let mut numbering = if context.package.parts.contains_key("word/numbering.xml") {
            context.package.xml("word/numbering.xml", &mut self.ns)?
        } else {
            Element::new(W, "numbering")
        };
        for element in numbering.elements() {
            let (attribute, map) = match element.name.local.as_str() {
                "num" => ("numId", &mut context.nums),
                "abstractNum" => ("abstractNumId", &mut context.abstracts),
                "numPicBullet" => ("numPicBulletId", &mut context.pictures),
                _ => continue,
            };
            if let Some(id) = element.attr(W, attribute) {
                map.insert(id.into(), self.next_id.to_string());
                self.next_id += 1;
            }
        }
        let defaults = styles.child(W, "docDefaults").cloned();
        let mut base = Element::new(W, "style");
        base.set(W, "type", "paragraph");
        base.set(W, "styleId", &defaults_id);
        let mut base_name = Element::new(W, "name");
        base_name.set(W, "val", &defaults_id);
        base.push(base_name);
        if let Some(defaults) = &defaults {
            for (wrapper, property) in [("pPrDefault", "pPr"), ("rPrDefault", "rPr")] {
                if let Some(p) = defaults
                    .child(W, wrapper)
                    .and_then(|e| e.child(W, property))
                {
                    base.push(p.clone());
                }
            }
        }
        transform(
            &mut base,
            &context,
            &mut self.resources,
            &mut self.next_id,
            "word/styles.xml",
            "word/styles.xml",
            false,
        )?;
        self.styles.push(base);
        let mut found_default = false;
        for node in &mut styles.children {
            let Node::Element(style) = node else { continue };
            if !style.is(W, "style") {
                continue;
            }
            if style.attr(W, "styleId") == Some(default.as_str()) {
                found_default = true;
            }
            if style.attr(W, "type") == Some("paragraph") && style.child(W, "basedOn").is_none() {
                let mut based = Element::new(W, "basedOn");
                based.set(W, "val", &defaults_id);
                style.push(based);
            }
            transform(
                style,
                &context,
                &mut self.resources,
                &mut self.next_id,
                "word/styles.xml",
                "word/styles.xml",
                false,
            )?;
            self.styles.push(style.clone());
        }
        if !found_default {
            let mut default = Element::new(W, "style");
            default.set(W, "type", "paragraph");
            default.set(W, "styleId", &default_id(&context));
            let mut based = Element::new(W, "basedOn");
            based.set(W, "val", &defaults_id);
            default.push(based);
            transform(
                &mut default,
                &context,
                &mut self.resources,
                &mut self.next_id,
                "word/styles.xml",
                "word/styles.xml",
                false,
            )?;
            self.styles.push(default);
        }
        for node in &mut numbering.children {
            if let Node::Element(element) = node {
                transform(
                    element,
                    &context,
                    &mut self.resources,
                    &mut self.next_id,
                    "word/numbering.xml",
                    "word/numbering.xml",
                    false,
                )?;
                self.numbering.push(element.clone());
            }
        }
        Ok(context)
    }

    pub fn finish(mut self, output: &Path) -> Result<()> {
        let ignorable = self.ignorable_prefixes();
        if !ignorable.is_empty() {
            for root in [&mut self.styles, &mut self.numbering, &mut self.fonts] {
                root.set(
                    "http://schemas.openxmlformats.org/markup-compatibility/2006",
                    "Ignorable",
                    &ignorable,
                );
            }
        }
        let style_order = [
            "name",
            "aliases",
            "basedOn",
            "next",
            "link",
            "autoRedefine",
            "hidden",
            "uiPriority",
            "semiHidden",
            "unhideWhenUsed",
            "qFormat",
            "locked",
            "personal",
            "personalCompose",
            "personalReply",
            "rsid",
            "pPr",
            "rPr",
            "tblPr",
            "trPr",
            "tcPr",
            "tblStylePr",
        ];
        for node in &mut self.styles.children {
            if let Node::Element(style) = node {
                style.children.sort_by_key(|node| match node {
                    Node::Element(e) => style_order
                        .iter()
                        .position(|name| *name == e.name.local)
                        .unwrap_or(style_order.len()),
                    Node::Text(_) => style_order.len() + 1,
                });
            }
        }
        self.numbering.children.sort_by_key(|node| match node {
            Node::Element(e) => match e.name.local.as_str() {
                "numPicBullet" => 0,
                "abstractNum" => 1,
                "num" => 2,
                _ => 3,
            },
            Node::Text(_) => 4,
        });
        let mut archive = zip::ZipArchive::new(File::open(output)?)?;
        let mut original = BTreeMap::new();
        for index in 0..archive.len() {
            let mut entry = archive.by_index(index)?;
            let mut bytes = Vec::new();
            entry.read_to_end(&mut bytes)?;
            original.insert(entry.name().to_owned(), bytes);
        }
        drop(archive);
        let content = Element::parse(
            &xml::decode(&original["[Content_Types].xml"])?,
            &mut self.ns,
        )?;
        for element in content.elements() {
            if let (Some(path), Some(kind)) = (
                element.attr("", "PartName"),
                element.attr("", "ContentType"),
            ) {
                self.resources
                    .types
                    .insert(path.trim_start_matches('/').into(), kind.into());
            }
        }
        self.resources.parts.insert(
            "word/styles.xml".into(),
            self.styles.serialize(&self.ns, true).into_bytes(),
        );
        self.resources.add_relationship(
            "word/document.xml",
            "rId1",
            &format!("{R}/styles"),
            "styles.xml",
        );
        if !self.numbering.children.is_empty() {
            self.resources.parts.insert(
                "word/numbering.xml".into(),
                self.numbering.serialize(&self.ns, true).into_bytes(),
            );
            self.resources.types.insert(
                "word/numbering.xml".into(),
                "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"
                    .into(),
            );
            self.resources.add_relationship(
                "word/document.xml",
                "numbering",
                &format!("{R}/numbering"),
                "numbering.xml",
            );
        }
        if !self.fonts.children.is_empty() {
            self.resources.parts.insert(
                "word/fontTable.xml".into(),
                self.fonts.serialize(&self.ns, true).into_bytes(),
            );
            self.resources.types.insert(
                "word/fontTable.xml".into(),
                "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"
                    .into(),
            );
            self.resources.add_relationship(
                "word/document.xml",
                "fonts",
                &format!("{R}/fontTable"),
                "fontTable.xml",
            );
        }
        self.resources.finish_relationships(&mut self.ns);
        original.extend(self.resources.parts.clone());
        original.insert(
            "[Content_Types].xml".into(),
            self.resources.content_types(&mut self.ns).into_bytes(),
        );
        let mut temporary =
            tempfile::NamedTempFile::new_in(output.parent().unwrap_or(Path::new(".")))?;
        {
            let mut zip = zip::ZipWriter::new(&mut temporary);
            for (path, bytes) in original {
                zip.start_file(
                    path,
                    zip::write::SimpleFileOptions::default()
                        .compression_method(zip::CompressionMethod::Deflated),
                )?;
                zip.write_all(&bytes)?;
            }
            zip.finish()?;
        }
        temporary.persist(output).map_err(|error| error.error)?;
        Ok(())
    }
}

#[derive(Default)]
struct BodyIds {
    next: u64,
    bookmarks: HashMap<String, Vec<String>>,
    names: HashSet<String>,
    anchors: HashMap<String, String>,
    vml: HashMap<String, String>,
    vml_used: HashSet<String>,
    paired: HashSet<String>,
}
fn sanitize_ids(element: &mut Element, state: &mut BodyIds) {
    if element.name.uri == "urn:schemas-microsoft-com:vml" {
        if let Some(original) = element.attr("", "id").map(str::to_owned) {
            let mut id = original.clone();
            while state.vml_used.contains(&id) {
                state.next += 1;
                id = format!("{original}_{}", state.next);
            }
            state.vml_used.insert(id.clone());
            state.vml.insert(original, id.clone());
            element.set("", "id", id);
        }
    }
    if element.is(W, "instrText") {
        for child in &mut element.children {
            if let Node::Text(text) = child {
                *text = crate::hyperlinks::rewrite_instruction(text, &state.anchors);
            }
        }
    }
    if element.is(W, "bookmarkStart") {
        if let Some(original) = element.attr(W, "id").map(str::to_owned) {
            state.next += 1;
            let id = state.next.to_string();
            state
                .bookmarks
                .entry(original)
                .or_default()
                .push(id.clone());
            element.set(W, "id", id);
        }
        if let Some(original) = element.attr(W, "name").map(str::to_owned) {
            let mut name = original.chars().take(40).collect::<String>();
            let mut counter = 1;
            while state.names.contains(&name) {
                let suffix = format!("_{counter}");
                name = format!(
                    "{}{}",
                    original.chars().take(40 - suffix.len()).collect::<String>(),
                    suffix
                );
                counter += 1;
            }
            state.names.insert(name.clone());
            state.anchors.insert(original, name.clone());
            element.set(W, "name", name);
        }
    } else if element.is(W, "bookmarkEnd") {
        if let Some(id) = element
            .attr(W, "id")
            .and_then(|id| state.bookmarks.get_mut(id))
            .and_then(Vec::pop)
        {
            state.paired.insert(id.clone());
            element.set(W, "id", id);
        } else {
            element.set(W, "id", "-1");
        }
    }
    if element.is(W, "hyperlink") && element.attr(R, "id").is_none() {
        if let Some(anchor) = element.attr(W, "anchor").map(str::to_owned) {
            let name = state
                .anchors
                .get(&anchor)
                .cloned()
                .unwrap_or_else(|| anchor.chars().take(40).collect());
            element.set(W, "anchor", name);
        }
    }
    if element.name.uri == "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
        && element.name.local == "docPr"
    {
        state.next += 1;
        element.set("", "id", state.next.to_string());
    }
    for (name, value) in &mut element.attrs {
        if name.uri == W && name.local == "instr" {
            *value = crate::hyperlinks::rewrite_instruction(value, &state.anchors);
        }
        if name.local == "type" && value.starts_with('#') {
            if let Some(mapped) = state.vml.get(&value[1..]) {
                *value = format!("#{mapped}");
            }
        }
        if name.local == "shapeid" || name.local == "ShapeID" {
            if let Some(mapped) = state.vml.get(value) {
                *value = mapped.clone();
            }
        }
        if name.uri == "urn:schemas-microsoft-com:office:office" && name.local == "spid" {
            state.next += 1;
            *value = format!("_x0000_s{}", state.next);
        }
        if name.local == "paraId"
            || (name.uri == "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
                && ["anchorId", "editId"].contains(&name.local.as_str()))
        {
            state.next += 1;
            *value = format!("{:08X}", state.next);
        }
    }
    let sdt_properties = element.is(W, "sdtPr");
    for child in &mut element.children {
        if let Node::Element(child) = child {
            if sdt_properties && child.is(W, "id") {
                state.next += 1;
                child.set(W, "val", state.next.to_string());
            }
            sanitize_ids(child, state);
        }
    }
}

fn transform(
    element: &mut Element,
    context: &Context,
    resources: &mut Resources,
    next: &mut u64,
    source_owner: &str,
    destination_owner: &str,
    paragraph_defaults: bool,
) -> Result<()> {
    context.theme.materialize(element);
    let external_anchor = element.is(W, "hyperlink") && element.attr(R, "id").is_some();
    if element.is(W, "instrText") {
        for child in &mut element.children {
            if let Node::Text(text) = child {
                *text = crate::hyperlinks::rewrite_instruction(text, &context.bookmark_names);
            }
        }
    }
    for (name, value) in &mut element.attrs {
        if name.uri == W && name.local == "instr" {
            *value = crate::hyperlinks::rewrite_instruction(value, &context.bookmark_names);
            *value = crate::hyperlinks::rebase_external_instruction(value, &mut |address| {
                context.package.external_target(address)
            })?;
        }
        if element.name.uri == "urn:schemas-microsoft-com:vml"
            && name.uri.is_empty()
            && name.local == "id"
        {
            if let Some(mapped) = context.vml_names.get(value) {
                *value = mapped.clone();
            }
        }
        if name.local == "type" && value.starts_with('#') {
            if let Some(mapped) = context.vml_names.get(&value[1..]) {
                *value = format!("#{mapped}");
            }
        }
        if name.local == "shapeid" || name.local == "ShapeID" {
            if let Some(mapped) = context.vml_names.get(value) {
                *value = mapped.clone();
            }
        }
        if name.uri == R && ["id", "embed", "link"].contains(&name.local.as_str()) {
            *value = resources.import_relationship(
                &context.package,
                source_owner,
                value,
                destination_owner,
            )?;
        }
        if name.uri == W {
            if name.local == "styleId"
                || (name.local == "val"
                    && [
                        "pStyle",
                        "rStyle",
                        "tblStyle",
                        "basedOn",
                        "next",
                        "link",
                        "numStyleLink",
                        "styleLink",
                    ]
                    .contains(&element.name.local.as_str()))
            {
                *value = format!("{}_{}", context.package.tag, value);
            }
            let map = match (element.name.local.as_str(), name.local.as_str()) {
                ("num", "numId") | ("numId", "val") => Some(&context.nums),
                ("abstractNum", "abstractNumId") | ("abstractNumId", "val") => {
                    Some(&context.abstracts)
                }
                ("numPicBullet", "numPicBulletId") | ("lvlPicBulletId", "val") => {
                    Some(&context.pictures)
                }
                _ => None,
            };
            if let Some(map) = map {
                if let Some(mapped) = map.get(value) {
                    *value = mapped.clone();
                }
            }
            if (name.local == "anchor" && !external_anchor)
                || (name.local == "name" && element.name.local == "bookmarkStart")
            {
                if let Some(mapped) = context.bookmark_names.get(value) {
                    *value = mapped.clone();
                }
            }
            if name.local == "id"
                && ["bookmarkStart", "bookmarkEnd"].contains(&element.name.local.as_str())
            {
                *value = format!(
                    "{}",
                    value.parse::<u64>().unwrap_or(0)
                        + context
                            .package
                            .tag
                            .trim_start_matches('s')
                            .parse::<u64>()
                            .unwrap_or(0)
                            * 1000000
                );
            }
        }
    }
    if element.is(W, "style") {
        element.remove_attr(W, "default");
        for child in &mut element.children {
            if let Node::Element(e) = child {
                if e.is(W, "name") {
                    if let Some(name) = e.attr(W, "val") {
                        e.set(W, "val", format!("{} {name}", context.package.tag));
                    }
                }
                if e.is(W, "aliases") {
                    if let Some(names) = e.attr(W, "val") {
                        let aliases = names
                            .split(',')
                            .map(|name| format!("{} {}", context.package.tag, name.trim()))
                            .collect::<Vec<_>>()
                            .join(",");
                        e.set(W, "val", aliases);
                    }
                }
            }
        }
    }
    if element.name.local == "docPr"
        && element.name.uri
            == "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    {
        element.set("", "id", next.to_string());
        *next += 1;
    }
    for child in &mut element.children {
        if let Node::Element(child) = child {
            transform(
                child,
                context,
                resources,
                next,
                source_owner,
                destination_owner,
                paragraph_defaults,
            )?;
        }
    }
    if paragraph_defaults && element.is(W, "p") {
        let position = element
            .children
            .iter()
            .position(|n| matches!(n,Node::Element(e) if e.is(W,"pPr")));
        let index = if let Some(i) = position {
            i
        } else {
            element
                .children
                .insert(0, Node::Element(Element::new(W, "pPr")));
            0
        };
        if let Node::Element(properties) = &mut element.children[index] {
            if properties.child(W, "pStyle").is_none() {
                let mut style = Element::new(W, "pStyle");
                style.set(W, "val", &context.default_style);
                properties.children.insert(0, Node::Element(style));
            }
        }
    }
    if paragraph_defaults && element.is(W, "tbl") {
        if let Some(style) = &context.default_table_style {
            let index = element
                .children
                .iter()
                .position(|n| matches!(n,Node::Element(e) if e.is(W,"tblPr")))
                .unwrap_or_else(|| {
                    element
                        .children
                        .insert(0, Node::Element(Element::new(W, "tblPr")));
                    0
                });
            if let Node::Element(properties) = &mut element.children[index] {
                if properties.child(W, "tblStyle").is_none() {
                    let mut reference = Element::new(W, "tblStyle");
                    reference.set(W, "val", style);
                    properties.children.insert(0, Node::Element(reference));
                }
            }
        }
    }
    Ok(())
}

fn default_id(context: &Context) -> String {
    context
        .default_style
        .strip_prefix(&format!("{}_", context.package.tag))
        .unwrap_or("Normal")
        .to_owned()
}

fn source_uri(path: &Path) -> Result<String> {
    let full = std::fs::canonicalize(path)?;
    let text = full.to_string_lossy();
    let text = text
        .strip_prefix("\\\\?\\")
        .unwrap_or(&text)
        .replace('\\', "/");
    let text = if let Some(unc) = text.strip_prefix("UNC/") {
        format!("//{unc}")
    } else {
        text
    };
    let escaped = text
        .bytes()
        .map(|b| {
            if b.is_ascii_alphanumeric() || b"/:-_.~".contains(&b) {
                (b as char).to_string()
            } else {
                format!("%{b:02X}")
            }
        })
        .collect::<String>();
    Ok(if escaped.starts_with("//") {
        format!("file:{escaped}")
    } else if escaped.starts_with('/') {
        format!("file://{escaped}")
    } else {
        format!("file:///{escaped}")
    })
}

fn remove_unpaired_bookmarks(element: &mut Element, paired: &HashSet<String>) {
    element.children.retain(|node|!matches!(node,Node::Element(e) if (e.is(W,"bookmarkStart") || e.is(W,"bookmarkEnd")) && !e.attr(W,"id").is_some_and(|id|paired.contains(id))));
    for child in &mut element.children {
        if let Node::Element(child) = child {
            remove_unpaired_bookmarks(child, paired);
        }
    }
}
fn collect_bookmark_names(element: &Element) -> HashSet<String> {
    let mut result = HashSet::new();
    if element.is(W, "bookmarkStart") {
        if let Some(name) = element.attr(W, "name") {
            result.insert(name.into());
        }
    }
    for child in element.elements() {
        result.extend(collect_bookmark_names(child));
    }
    result
}
fn repair_links(
    element: &mut Element,
    names: &HashSet<String>,
    targets: &HashMap<String, (String, String)>,
    resources: &mut Resources,
) -> Result<()> {
    let omitted = targets
        .iter()
        .filter(|(name, _)| !names.contains(*name))
        .map(|(name, target)| (name.clone(), target.clone()))
        .collect::<HashMap<_, _>>();
    repair_links_inner(element, names, &omitted, resources)
}
fn repair_links_inner(
    element: &mut Element,
    names: &HashSet<String>,
    omitted: &HashMap<String, (String, String)>,
    resources: &mut Resources,
) -> Result<()> {
    if element.is(W, "hyperlink") && element.attr(R, "id").is_none() {
        if let Some(anchor) = element.attr(W, "anchor").map(str::to_owned) {
            if !names.contains(&anchor)
                && !["_top", "_bottom", "_startOfDoc", "_endOfDoc"].contains(&anchor.as_str())
            {
                let (uri, original) = omitted.get(&anchor).with_context(|| {
                    format!("Hyperlink target '{anchor}' does not exist in the source document")
                })?;
                let id =
                    resources.add_external("word/document.xml", &format!("{R}/hyperlink"), uri);
                element.set(R, "id", id);
                element.set(W, "anchor", original);
            }
        }
    }
    if element.is(W, "fldSimple") {
        if let Some(code) = element.attr(W, "instr") {
            let changed = crate::hyperlinks::externalize_instruction(code, omitted);
            element.set(W, "instr", changed);
        }
    }
    if element.is(W, "instrText") {
        for child in &mut element.children {
            if let Node::Text(text) = child {
                *text = crate::hyperlinks::externalize_instruction(text, omitted);
            }
        }
    }
    for child in &mut element.children {
        if let Node::Element(child) = child {
            repair_links_inner(child, names, omitted, resources)?;
        }
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn theme_tint_matches_word_rendering_with_one_channel_rounding_tolerance() {
        for (source, word) in [("993322", "DB7160"), ("2244AA", "6584E0")] {
            let actual = adjust_color(source, Some("99"), None);
            for index in [0, 2, 4] {
                let a = u8::from_str_radix(&actual[index..index + 2], 16).unwrap();
                let b = u8::from_str_radix(&word[index..index + 2], 16).unwrap();
                assert!(a.abs_diff(b) <= 1);
            }
        }
    }
    #[test]
    fn cloned_bookmarks_get_unique_names_and_unpaired_markers_are_removed() {
        let mut namespaces = Namespaces::new();
        let mut root=Element::parse(&format!("<w:p xmlns:w=\"{W}\"><w:bookmarkEnd w:id=\"99\"/><w:bookmarkStart w:id=\"1\" w:name=\"same\"/><w:bookmarkEnd w:id=\"1\"/><w:hyperlink w:anchor=\"same\"/><w:bookmarkStart w:id=\"1\" w:name=\"same\"/><w:bookmarkEnd w:id=\"1\"/><w:hyperlink w:anchor=\"same\"/></w:p>"),&mut namespaces).unwrap();
        let mut state = BodyIds::default();
        sanitize_ids(&mut root, &mut state);
        remove_unpaired_bookmarks(&mut root, &state.paired);
        let names = collect_bookmark_names(&root);
        assert_eq!(names.len(), 2);
        let links = root
            .elements()
            .filter(|e| e.is(W, "hyperlink"))
            .map(|e| e.attr(W, "anchor").unwrap())
            .collect::<Vec<_>>();
        assert_eq!(links, vec!["same", "same_1"]);
        assert_eq!(
            root.elements().filter(|e| e.is(W, "bookmarkEnd")).count(),
            2
        );
    }
}
