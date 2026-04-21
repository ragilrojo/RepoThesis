const fs = require('fs');
const path = 'e:\\ProjectNodeJs\\temp_doc_build\\generate_ppt.js';
let content = fs.readFileSync(path, 'utf8');
const lines = content.split('\n');

// Find the corrupted part (around line 37-38 in the view_file, which is indices 36-37)
// We need to inject the missing lines between slideTOC1.addImage of logo and Landasan Teori.

const slideTOCSectionStart = '    slideTOC1.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });';
const landasanTeoriLine = '        { text: "Landasan Teori & Literatur", options: { hyperlink: { slide: \'6\' }, fontSize: 16 } },';

let startIdx = -1;
for (let i = 0; i < lines.length; i++) {
    if (lines[i].includes('slideTOC1.addImage({ path: "logo_unm.png"') && lines[i].includes('slideTOC1')) {
        startIdx = i;
        break;
    }
}

if (startIdx !== -1) {
    const missingSnippet = [
        '    slideTOC1.addText("Daftar Isi (Main Sections)", { x: 0.5, y: 0.5, w: "90%", fontSize: 28, bold: true, color: "003366" });',
        '    ',
        '    // Kolom Kiri: Pendahuluan & Strategi',
        '    slideTOC1.addText([',
        '        { text: "I. PENDAHULUAN", options: { bold: true, color: "003366", breakLine: true } },',
        '        { text: "   • ", options: {} },',
        '        { text: "Latar Belakang", options: { hyperlink: { slide: \'4\' }, fontSize: 16 } },',
        '        { text: "", options: { breakLine: true } },',
        '        { text: "   • ", options: {} },',
        '        { text: "Masalah \\"Noise\\" di Kripto", options: { hyperlink: { slide: \'5\' }, fontSize: 16 } },',
        '        { text: "", options: { breakLine: true } },',
        '        { text: "   • ", options: {} },'
    ];
    
    // Check if Landasan Teori is actually the next line
    if (lines[startIdx + 1].includes('Landasan Teori')) {
        lines.splice(startIdx + 1, 0, ...missingSnippet);
        fs.writeFileSync(path, lines.join('\n'));
        console.log('Restored TOC 1 slide with correct content.');
    } else {
        console.log('Mismatch in restoration point.');
    }
}
