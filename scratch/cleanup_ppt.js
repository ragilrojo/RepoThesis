const fs = require('fs');
const path = 'e:\\ProjectNodeJs\\temp_doc_build\\generate_ppt.js';
let content = fs.readFileSync(path, 'utf8');
const lines = content.split('\n');

// Lines are 1-indexed in my mind, but 0-indexed in array.
// From view_file:
// 1243 is empty/comma ?
// 1244 is "1. Siapkan Kandidat"
// ...
// 1260 is "Kesimpulan Adaptif"

// I want to remove the redundant old block.
// I'll look for the first occurrence of "1. Siapkan Kandidat" after line 1230.
let startIdx = -1;
for (let i = 1240; i < lines.length; i++) {
    if (lines[i].includes('1. Siapkan Kandidat')) {
        startIdx = i;
        break;
    }
}

if (startIdx !== -1) {
    // Look for the end of that block (the ] before the footnote)
    let endIdx = -1;
    for (let i = startIdx; i < lines.length; i++) {
        if (lines[i].includes('], { x: 0.5, y: 1.1')) {
            endIdx = i;
            break;
        }
    }

    if (endIdx !== -1) {
        // We want to keep the closing bracket and the footnote logic, 
        // but we already have one from the previous edit if I didn't mess up.
        // Wait, I replaced line 1230 with a long block.
        // Let's just remove lines from startIdx to just BEFORE the closing bracket.
        lines.splice(startIdx - 1, endIdx - startIdx + 1); 
        fs.writeFileSync(path, lines.join('\n'));
        console.log('Cleanup successful');
    }
}
