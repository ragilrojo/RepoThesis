const fs = require('fs');
const path = 'e:\\ProjectNodeJs\\temp_doc_build\\generate_ppt.js';
let content = fs.readFileSync(path, 'utf8');

// 1. Remove TOC entry
// Looks like: { text: "   • ", options: {} }, \n { text: "Masalah \"Noise\" di Kripto", ... }, \n { text: "", options: { breakLine: true } },
const tocPattern = /\{ text: "   • ", options: \{\} \},\n\s+\{ text: "Masalah \\"Noise\\" di Kripto", options: \{ hyperlink: \{ slide: '5' \}, fontSize: 16 \} \},\n\s+\{ text: "", options: \{ breakLine: true \} \},/g;
content = content.replace(tocPattern, '');

// 2. Remove Slide 5 (slide3 variable)
// Starts around // --- Slide 3: Konsep "Noise" ...
// Ends before slide4 (Landasan Teori)
const slidePattern = /\/\/ --- Slide 3: Konsep "Noise" dalam Cryptocurrency ---[\s\S]*?(?=\/\/ --- Slide 4: Landasan Teori)/g;
content = content.replace(slidePattern, '');

// 3. Decrement Hyperlinks >= 6
// Pattern: hyperlink: { slide: '(\d+)' }
content = content.replace(/hyperlink: \{ slide: '(\d+)' \}/g, (match, slideNum) => {
    let num = parseInt(slideNum);
    if (num >= 6) {
        return `hyperlink: { slide: '${num - 1}' }`;
    }
    return match;
});

fs.writeFileSync(path, content);
console.log('Removed Noise slide and entry, updated all subsequent slide links.');
