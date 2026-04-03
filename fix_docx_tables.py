from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import os

def fix_docx_tables(docx_path):
    print(f"Cleaning table XML in {docx_path}...")
    if not os.path.exists(docx_path):
        print("Error: File not found.")
        return
        
    doc = Document(docx_path)
    
    # Iterate through all tables in the document
    for table in doc.tables:
        # 1. Clear potentially corrupt table-level properties (tblPr)
        # Word often repairs tables when they have complex grid measurements or conflicting border styles.
        tbl = table._tbl
        tblPr = tbl.find(qn('w:tblPr'))
        
        if tblPr is not None:
            # Remove Look settings which can sometimes cause repair warnings in older Word
            look = tblPr.find(qn('w:tblLook'))
            if look is not None:
                tblPr.remove(look)
        
        # 2. Fix every cell border
        for row in table.rows:
            for cell in row.cells:
                tc = cell._tc
                tcPr = tc.get_or_add_tcPr()
                
                # Ensure no blank tcW or other weird properties
                tcW = tcPr.find(qn('w:tcW'))
                if tcW is not None:
                    # Let Word handle auto-width if it's currently 0 or corrupt
                    if tcW.get(qn('w:w')) == "0":
                         tcPr.remove(tcW)

    doc.save(docx_path)
    print("Success! Table XML structure cleaned.")

if __name__ == "__main__":
    target = "e:/ProjectNodeJs/temp_doc_build/proposal_tesis_ragil_mendeley.docx"
    fix_docx_tables(target)
