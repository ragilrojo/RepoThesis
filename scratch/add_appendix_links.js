const fs = require('fs');
const path = 'e:\\ProjectNodeJs\\temp_doc_build\\generate_ppt.js';
let content = fs.readFileSync(path, 'utf8');

// The pattern to look for (main TOC link)
const pattern = /slide[a-zA-Z0-9]*\.addText\("🏠 Daftar Isi", \{ x: 8\.5, y: 5\.3, w: 1\.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: \{ slide: '2' \}, align: "right" \}\);/g;

// The link to TOC 2 (Lampiran)
const appendixLink = '\n    $&.replace(\'x: 8.5\', \'x: 7.3\').replace(\'🏠 Daftar Isi\', \'📂 Lampiran\').replace("\'2\'", "\'3\'")';
// Wait, that matched approach is a bit complex with regex backreferences in JS replace if using string replacement.
// I'll use a functional replacement.

const newContent = content.replace(pattern, (match) => {
    // Generate the second link by mimicking the first one
    let secondLink = match
        .replace('x: 8.5', 'x: 7.3')
        .replace('🏠 Daftar Isi', '📂 Lampiran')
        .replace("'2'", "'3'");
    return secondLink + '\n    ' + match;
});

fs.writeFileSync(path, newContent);
console.log('Added Lampiran link to all slides with a TOC link.');
