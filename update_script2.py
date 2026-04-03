import re

filepath = r'e:\ProjectNodeJs\temp_doc_build\generate_proposal.js'
with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Font size of headings
content = re.sub(
    r'function heading1\(text\)\s*{\s*return new Paragraph\({\s*heading: HeadingLevel\.HEADING_1,\s*alignment: AlignmentType\.CENTER,\s*children: \[new TextRun\({ text, font: "Times New Roman", size: 28, bold: true }\)\]\s*}\);\s*}',
    r'function heading1(text) {\n    return new Paragraph({\n        heading: HeadingLevel.HEADING_1,\n        alignment: AlignmentType.CENTER,\n        children: [new TextRun({ text, font: "Times New Roman", size: 24, bold: true })]\n    });\n}',
    content
)

content = re.sub(
    r'function heading2\(text\)\s*{\s*return new Paragraph\({\s*heading: HeadingLevel\.HEADING_2,\s*children: \[new TextRun\({ text, font: "Times New Roman", size: 26, bold: true }\)\]\s*}\);\s*}',
    r'function heading2(text) {\n    return new Paragraph({\n        heading: HeadingLevel.HEADING_2,\n        children: [new TextRun({ text, font: "Times New Roman", size: 24, bold: true })]\n    });\n}',
    content
)

content = re.sub(r'function chapterHeading\(bab, title, size = 26\)', r'function chapterHeading(bab, title, size = 24)', content)
content = re.sub(r'function sectionTitle\(text, size = 26\)', r'function sectionTitle(text, size = 24)', content)

# 2. Numbering dots
content = re.sub(r'heading2\("(\d+)\.(\d+)\.\s+', r'heading2("\1.\2 ', content)
content = re.sub(r'heading3\("(\d+)\.(\d+)\.(\d+)\.\s+', r'heading3("\1.\2.\3 ', content)

# 3. Add Header to docs imports
if 'Header,' not in content:
    content = content.replace('Footer, NumberFormat', 'Header, Footer, NumberFormat')

# 4. Page numbering layout replacement
replacement_hf = """            headers: {
                default: new Header({
                    children: [
                        new Paragraph({
                            alignment: AlignmentType.RIGHT,
                            children: [new TextRun({ children: [PageNumber.CURRENT], font: "Times New Roman", size: 24 })]
                        })
                    ]
                })
            },
            footers: {
                first: new Footer({
                    children: [
                        new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [new TextRun({ children: [PageNumber.CURRENT], font: "Times New Roman", size: 24 })]
                        })
                    ]
                })
            }"""

target_hf = """            footers: {
                default: new Footer({
                    children: [
                        new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [new TextRun({ children: [PageNumber.CURRENT], font: "Times New Roman", size: 24 })]
                        })
                    ]
                })
            }"""

if target_hf in content:
    content = content.replace(target_hf, replacement_hf)

# 5. Add titlePage: true to BAB sections (the ones that use Type.NEXT_PAGE and have PageNumber.CURRENT later)
# A robust way is just to add it unconditionally to all NEXT_PAGE sections. The roman ones don't use titlePage features anyway since their footer is static.
content = re.sub(r'(type:\s*SectionType\.NEXT_PAGE)(?!,?\s*titlePage:)', r'\1, titlePage: true', content)

with open(filepath, 'w', encoding='utf-8') as f:
    f.write(content)

print("Updates applied.")
