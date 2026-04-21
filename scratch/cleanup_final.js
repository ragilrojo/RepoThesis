const fs = require('fs');
const path = 'e:\\ProjectNodeJs\\temp_doc_build\\generate_ppt.js';
let content = fs.readFileSync(path, 'utf8');
const lines = content.split('\n');

// Find the duplicate block of TOC2
// It starts with slideTOC2.addImage... at line 143 (index 142)
// and ends with ⬅ Kembali ke TOC 1 at line 186 (index 185)

// Actually, I'll look for the second occurrence of TOC2 construction.
let firstTOC2Line = -1;
let secondTOC2Line = -1;

for (let i = 0; i < lines.length; i++) {
    if (lines[i].includes('let slideTOC2 = pres.addSlide();')) {
        if (firstTOC2Line === -1) {
            firstTOC2Line = i;
        } else {
            secondTOC2Line = i;
            break;
        }
    }
}

// Wait, the duplicate doesn't have "let slideTOC2" inside it, it just re-adds elements to slideTOC2.
// From view_file:
// Line 143: slideTOC2.addImage...
// ...
// Line 186: slideTOC2.addText("⬅ Kembali ke TOC 1", ...

let targetLineStart = -1;
let count = 0;
for (let i = 0; i < lines.length; i++) {
    if (lines[i].includes('slideTOC2.addImage({ path: "logo_unm.png"')) {
        count++;
        if (count === 2) { // The first one is from the first slide, wait.
            // Actually, slideTOC1 has one, slideTOC2 has one.
            // Let's be more specific.
        }
    }
}

// I'll just use the line numbers from view_file (1-indexed)
// Line 143 to 187 seems to be the redundant block.
lines.splice(142, 187 - 142 + 1); 

fs.writeFileSync(path, lines.join('\n'));
console.log('Cleaned up duplicated TOC2 code.');
