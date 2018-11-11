extern crate zip;
extern crate minidom;
extern crate wpm;

use minidom::Element;
use wpm::{Document, RootEntry, Run, RunElement};

pub fn import(xml_data: &str) -> Result<Document, String> {
    let root : Element = xml_data.parse().unwrap();
    let w_ns = root.attr("xmlns:w").unwrap_or("http://schemas.openxmlformats.org/wordprocessingml/2006/main");

    let body = root.get_child("body", w_ns).expect("no body found");
    let doc = Document {
        entries: body.children()
            .map(|be| {
                match be.name() {
                    "p" => {
                        if let Some(r) = be.get_child("r", w_ns) {
                            RootEntry::Paragraph{
                                run: Some(Run{
                                    elements: r.children()
                                        .map(|re| {
                                            match re.name() {
                                                "t" => {
                                                    RunElement::Text{
                                                        value: String::from(re.text())
                                                    }
                                                },
                                                _ => RunElement::Unknown
                                            }
                                        })
                                        .collect()
                                })
                            }
                        } else {
                            RootEntry::Paragraph{
                                run: None
                            }
                        }
                    }
                    _ => RootEntry::Unknown
                }
            })
            .collect()
    };
    Ok(doc)
}

#[cfg(test)]
mod tests {
    use super::import;
    use super::wpm::{Document, RootEntry, Run, RunElement};

    #[test]
    fn it_works() {
        let xml_data = r#"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:ve="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml">
<w:body>
    <w:p>
        <w:pPr>
            <w:pStyle w:val="Heading1"/>
        </w:pPr>
        <w:r><w:t>Introduction</w:t></w:r>
    </w:p>
</w:body>
</w:document>
"#;
        let result = import(xml_data);
        if let Some(document) = result.ok() {
            assert_eq!(document.entries.len(), 1);
            if let Some(RootEntry::Paragraph{
                run
            }) = document.entries.last() {
                if let Some(r) = run {
                    assert_eq!(r.elements.len(), 1);
                    if let Some(RunElement::Text{
                        value
                    }) = r.elements.last() {
                        assert_eq!(*value, String::from("Introduction"));
                    } else {
                        assert!(false);
                    }
                } else {
                    assert!(false);
                }
            } else {
                assert!(false);
            }
        } else {
            assert!(false);
        }
    }
}
