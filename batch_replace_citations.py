import re
import os

# Mapping for common citations based on references.bib order
mapping = {
    "[1]": "[mendeley_cite:giudici2020network]",
    "[2]": "[mendeley_cite:momeni2021portfolio]",
    "[3]": "[mendeley_cite:marchenko1967distribution]",
    "[4]": "[mendeley_cite:markowitz1952portfolio]",
    "[5]": "[mendeley_cite:mantegna1999hierarchical]",
    "[6]": "[mendeley_cite:jing2023network]",
    "[7]": "[mendeley_cite:kitanovski2024network]",
    "[8]": "[mendeley_cite:kitanovski2022cryptocurrency]",
    "[9]": "[mendeley_cite:giudici2021network]",
    "[10]": "[mendeley_cite:laloux1999noise]",
    "[11]": "[mendeley_cite:corbet2019cryptocurrencies]",
    "[12]": "[mendeley_cite:peralta2016network]",
    "[13]": "[mendeley_cite:lopezdeprado2018advances]",
    "[14]": "[mendeley_cite:michaud1989markowitz]",
    "[15]": "[mendeley_cite:eom2009topological]",
    "[16]": "[mendeley_cite:jorion2000value]",
    "[17]": "[mendeley_cite:rachev2008advanced]",
    "[18]": "[mendeley_cite:arratia2014computational]",
    "[19]": "[mendeley_cite:artzner1999coherent]",
    "[20]": "[mendeley_cite:lo2004adaptive]",
    "[21]": "[mendeley_cite:jing2025optimising]",
}

def replace_citations(file_path):
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()

    # Pre-processed: check for existing mendeley_cite patterns to avoid double replacement if they coexist with numerals
    # But numerals like [1] are distinct enough.
    
    for old, new in mapping.items():
        # Using escape to handle brackets since they are regex meta-characters
        escaped_old = re.escape(old)
        content = re.sub(escaped_old, new, content)

    # Backup then overwrite
    os.rename(file_path, file_path + ".bak")
    with open(file_path, 'w', encoding='utf-8', newline='\n') as f:
        f.write(content)
    print(f"Replacement complete for {file_path}. Backup created at {file_path}.bak")

if __name__ == "__main__":
    replace_citations("e:/ProjectNodeJs/temp_doc_build/generate_proposal.js")
