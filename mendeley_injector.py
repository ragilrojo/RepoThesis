import zipfile
import re
import json
import base64
import os
import uuid

# Parse simple BibTeX
def parse_bib(file_path):
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    entries = {}
    
    # Simple regex to split bibtex entries
    entry_pattern = re.compile(r'@([a-zA-Z]+)\{([^,]+),\s*(.*?)\n\}', re.DOTALL)
    for match in entry_pattern.finditer(content):
        entry_id = match.group(2).strip()
        fields_str = match.group(3)
        entries[entry_id] = {'type': match.group(1).lower(), 'id': entry_id}
        
        field_pattern = re.compile(r'([a-zA-Z0-9_]+)\s*=\s*[\{"]?(.*?)[\}"]?(?:,|$)', re.DOTALL)
        for f_match in field_pattern.finditer(fields_str):
            entries[entry_id][f_match.group(1).lower()] = re.sub(r'\s+', ' ', f_match.group(2).strip())
            
    return entries

def bib_to_csl(bib_entry):
    csl = {
        "id": bib_entry['id'],
        "type": "article-journal" if bib_entry['type'] == 'article' else "book",
        "title": bib_entry.get('title', ''),
    }
    if 'author' in bib_entry:
        csl['author'] = [{"family": p.strip().split(',')[0]} for p in bib_entry['author'].split(' and ')]
    if 'year' in bib_entry:
        csl['issued'] = {"date-parts": [[int(bib_entry['year'])]]}
    return csl

def generate_mendeley_sdt(csl_json):
    cit_uuid = str(uuid.uuid4())
    j_str = json.dumps({
        "citationID": f"MENDELEY_CITATION_{cit_uuid}",
        "properties": {"noteIndex": 0},
        "isEdited": False,
        "manualOverride": {"isManuallyOverridden": False, "citeprocText": "[1]", "manualOverrideText": ""},
        "citationItems": [{"id": csl_json['id'], "itemData": csl_json}]
    })
    b64_str = base64.b64encode(j_str.encode('utf-8')).decode('utf-8').replace('+', '-').replace('/', '_').rstrip('=')
    tag_val = f"MENDELEY_CITATION_v3_{b64_str}"
    sdt_id = str(uuid.uuid4().int & (1<<31)-1)
    return f'<w:sdt><w:sdtPr><w:rPr><w:color w:val="000000"/></w:rPr><w:tag w:val="{tag_val}"/><w:id w:val="{sdt_id}"/></w:sdtPr><w:sdtContent><w:r><w:rPr><w:color w:val="000000"/></w:rPr><w:t>[1]</w:t></w:r></w:sdtContent></w:sdt>'

def inject_citations(docx_path, bib_path, output_path):
    print("Loading references...")
    bib_entries = parse_bib(bib_path)
    
    try:
        with zipfile.ZipFile(docx_path, 'r') as zin:
            with zipfile.ZipFile(output_path, 'w') as zout:
                for item in zin.infolist():
                    # CRITICAL FIX: Skip any undefined media refs that were injected by Node.js
                    if '.undefined' in item.filename:
                        continue
                        
                    content = zin.read(item.filename)
                    
                    if item.filename == 'word/document.xml':
                        doc_xml = content.decode('utf-8')
                        def replace_cite(match):
                            cite_id = match.group(1)
                            if cite_id in bib_entries:
                                return f"</w:t></w:r>{generate_mendeley_sdt(bib_to_csl(bib_entries[cite_id]))}<w:r><w:t xml:space=\"preserve\">"
                            return match.group(0)
                        
                        new_xml = re.sub(r'\[mendeley_cite:([^\]]+)\]', replace_cite, doc_xml)
                        bib_sdt_id = str(uuid.uuid4().int & (1<<31)-1)
                        bib_sdt = f'<w:sdt><w:sdtPr><w:rPr><w:color w:val="000000"/></w:rPr><w:tag w:val="MENDELEY_BIBLIOGRAPHY"/><w:id w:val="{bib_sdt_id}"/><w:placeholder><w:docPart w:val="DefaultPlaceholder"/></w:placeholder></w:sdtPr><w:sdtContent><w:r><w:rPr><w:color w:val="000000"/></w:rPr><w:t>Daftar Pustaka terhubung dengan Mendeley (Klik Update Mendeley Cite).</w:t></w:r></w:sdtContent></w:sdt>'
                        new_xml = new_xml.replace('[mendeley_bibliography]', f'</w:t></w:r>{bib_sdt}<w:r><w:t xml:space="preserve">')
                        new_xml = new_xml.replace('<w:r><w:t xml:space="preserve"></w:t></w:r>', '')
                        content = new_xml.encode('utf-8')
                    
                    # Clean up Content_Types and Rel files from the .undefined junk
                    elif item.filename == '[Content_Types].xml':
                        content = re.sub(rb'<Override[^>]+PartName="/word/media/[^"]+\.undefined"[^>]+/>', b'', content)
                    elif item.filename == 'word/_rels/document.xml.rels':
                        content = re.sub(rb'<Relationship[^>]+Target="media/[^"]+\.undefined"[^>]+/>', b'', content)

                    zout.writestr(item, content)
        
        print(f"Success! Output saved to {output_path}")
        print("Note: Please insert logo manually in Word to avoid XML conflict.")
    except Exception as e:
        print(f"Error injecting citations: {e}")

if __name__ == '__main__':
    doc_in = "e:/ProjectNodeJs/temp_doc_build/proposal_tesis_ragil.docx"
    doc_out = "e:/ProjectNodeJs/temp_doc_build/proposal_tesis_ragil_mendeley.docx"
    bib_in = "e:/ProjectNodeJs/temp_doc_build/references.bib"
    
    if os.path.exists(doc_in):
        inject_citations(doc_in, bib_in, doc_out)
    else:
        print(f"Error: {doc_in} not found. Please run generate_proposal.js first.")
