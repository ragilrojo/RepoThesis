const fs = require('fs');
const path = 'e:\\ProjectNodeJs\\temp_doc_build\\generate_ppt.js';
let content = fs.readFileSync(path, 'utf8');

// Replace in TOC and everywhere else
let newContent = content.replace(/Landasan Teori & Literatur/g, 'Landasan Teori');

fs.writeFileSync(path, newContent);
console.log('Removed Literatur from Landasan Teori section.');
