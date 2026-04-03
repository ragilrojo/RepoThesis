import zipfile
import os
import re

def clean_undefined_media(docx_path):
    temp_path = docx_path + ".clean"
    print(f"Brute-force cleaning media from {docx_path}...")
    
    with zipfile.ZipFile(docx_path, 'r') as zin:
        with zipfile.ZipFile(temp_path, 'w') as zout:
            for item in zin.infolist():
                if item.filename.endswith('.undefined'):
                    continue # Skip the corrupted part
                
                content = zin.read(item.filename)
                
                # If it's [Content_Types].xml, remove any reference to the .undefined file
                if item.filename == '[Content_Types].xml':
                    content_str = content.decode('utf-8')
                    # Regex to remove the part with .undefined extension
                    content_str = re.sub(r'<Override[^>]+PartName="/word/media/[^"]+\.undefined"[^>]+/>', '', content_str)
                    content = content_str.encode('utf-8')
                
                # If it's the document.xml.rels, remove the relationship to the corrupted file
                if item.filename == 'word/_rels/document.xml.rels':
                    content_str = content.decode('utf-8')
                    content_str = re.sub(r'<Relationship[^>]+Target="media/[^"]+\.undefined"[^>]+/>', '', content_str)
                    content = content_str.encode('utf-8')

                zout.writestr(item, content)

    # Replace original with clean one
    os.remove(docx_path)
    os.rename(temp_path, docx_path)
    print("Success! Corrupted media references removed.")

if __name__ == "__main__":
    target = "e:/ProjectNodeJs/temp_doc_build/proposal_tesis_ragil_mendeley.docx"
    clean_undefined_media(target)
