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
        entry_type = match.group(1).lower()
        entry_id = match.group(2).strip()
        fields_str = match.group(3)
        
        entry = {
            'type': entry_type,
            'id': entry_id
        }
        
        # Parse fields
        field_pattern = re.compile(r'([a-zA-Z0-9_]+)\s*=\s*[\{"]?(.*?)[\}"]?(?:,|$)', re.DOTALL)
        for f_match in field_pattern.finditer(fields_str):
            key = f_match.group(1).lower()
            val = f_match.group(2).strip()
            # Clean newlines in val
            val = re.sub(r'\s+', ' ', val)
            entry[key] = val
        
        entries[entry_id] = entry
        
    return entries

def bib_to_csl(bib_entry):
    # Mapping BibTeX to CSL JSON loosely
    csl = {
        "id": bib_entry['id'],
        "type": "article-journal" if bib_entry['type'] == 'article' else "book",
        "title": bib_entry.get('title', ''),
    }
    
    # Authors
    if 'author' in bib_entry:
        authors = bib_entry['author'].split(' and ')
        csl['author'] = []
        for author in authors:
            parts = [p.strip() for p in author.split(',')]
            if len(parts) == 2:
                csl['author'].append({"family": parts[0], "given": parts[1]})
            else:
                csl['author'].append({"family": parts[0]})
                
    if 'journal' in bib_entry:
        csl['container-title'] = bib_entry['journal']
    if 'volume' in bib_entry:
        csl['volume'] = bib_entry['volume']
    if 'pages' in bib_entry:
        csl['page'] = bib_entry['pages']
    if 'year' in bib_entry:
        csl['issued'] = {"date-parts": [[int(bib_entry['year'])]]}
    if 'doi' in bib_entry:
        csl['DOI'] = bib_entry['doi']
        
    return csl

def generate_mendeley_sdt(csl_json):
    # Base64 without padding, using URL-safe chars as seen in Mendeley
    j_str = json.dumps({
        "citationItems": [{"id": csl_json['id'], "itemData": csl_json}],
        "properties": {"noteIndex": 0}
    })
    
    b64_str = base64.b64encode(j_str.encode('utf-8')).decode('utf-8')
    b64_str = b64_str.replace('+', '-').replace('/', '_').rstrip('=')
    
    tag_val = f"MENDELEY_CITATION_v3_{b64_str}"
    
    sdt_id = str(uuid.uuid4().int & (1<<31)-1)
    
    # Mendeley XML SDT format
    sdt = f'''<w:sdt><w:sdtPr><w:rPr><w:color w:val="000000"/></w:rPr><w:tag w:val="{tag_val}"/><w:id w:val="{sdt_id}"/></w:sdtPr><w:sdtContent><w:r><w:rPr><w:color w:val="000000"/></w:rPr><w:t>[1]</w:t></w:r></w:sdtContent></w:sdt>'''
    return sdt

def inject_citations(docx_path, bib_path, output_path):
    print("Loading references...")
    bib_entries = parse_bib(bib_path)
    
    with zipfile.ZipFile(docx_path, 'r') as zin:
        with zipfile.ZipFile(output_path, 'w') as zout:
            for item in zin.infolist():
                content = zin.read(item.filename)
                
                if item.filename == 'word/document.xml':
                    doc_xml = content.decode('utf-8')
                    
                    # Find all [mendeley_cite:XXX]
                    # Note: Because of Word/docx generator, the text might be exactly in <w:t>[mendeley_cite:id]</w:t>
                    # We regex replace the entire <w:r>...<w:t>[mendeley_cite:id]</w:t>...</w:r> if possible, or just replace the text.
                    # It's safer to just replace the placeholder text with the SDT string if the text is continuous.
                    
                    def replace_cite(match):
                        cite_id = match.group(1)
                        if cite_id in bib_entries:
                            csl = bib_to_csl(bib_entries[cite_id])
                            sdt = generate_mendeley_sdt(csl)
                            # Close out the current text run, insert SDT, restart a new run is complex. 
                            # If we just replace `[mendeley_cite:...]` inside the `<w:t>` with `</w:t></w:r>` + sdt + `<w:r><w:t>` it might be safest.
                            return f"</w:t></w:r>{sdt}<w:r><w:t xml:space=\"preserve\">"
                        else:
                            print(f"Warning: Citation ID {cite_id} not found in bib file!")
                            return match.group(0) # don't replace if not found
                    
                    new_xml = re.sub(r'\[mendeley_cite:([^\]]+)\]', replace_cite, doc_xml)
                    
                    # Fix empty text runs introduced by our split
                    new_xml = new_xml.replace('<w:r><w:t xml:space="preserve"></w:t></w:r>', '')
                    
                    zout.writestr(item, new_xml.encode('utf-8'))
                else:
                    zout.writestr(item, content)
    print(f"Success! Output saved to {output_path}")

if __name__ == '__main__':
    doc_in = "e:/ProjectNodeJs/temp_doc_build/proposal_tesis_ragil.docx"
    doc_out = "e:/ProjectNodeJs/temp_doc_build/proposal_tesis_ragil_mendeley.docx"
    bib_in = "e:/ProjectNodeJs/temp_doc_build/references.bib"
    
    if os.path.exists(doc_in):
        inject_citations(doc_in, bib_in, doc_out)
    else:
        print(f"Error: {doc_in} not found. Please run generate_proposal.js first.")
