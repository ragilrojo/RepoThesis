const fs = require('fs');
const path = 'e:\\ProjectNodeJs\\temp_doc_build\\generate_ppt.js';
let content = fs.readFileSync(path, 'utf8');

const target = '], { x: 5.2, y: 1.1, w: "45%", h: 5, fontSize: 16, color: "333333", valign: "top" });';
const replacement = target + '\n    slideTOC2.addText("🏠 Kembali ke Daftar Isi Utama", { x: 7.0, y: 5.3, w: 2.7, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: ' + "'2'" + ' }, align: "right" });';

// We need to find the instance after slideTOC2.addText([ (Bagian 2)
// slideTOC2 variable is used twice.
const slideTOC2Instance = content.indexOf('slideTOC2.addText([');
const secondInstance = content.indexOf('slideTOC2.addText([', slideTOC2Instance + 1);
const position = content.indexOf(target, secondInstance);

if (position !== -1) {
    let newContent = content.substring(0, position) + replacement + content.substring(position + target.length);
    fs.writeFileSync(path, newContent);
    console.log('Added Return to TOC 1 link to Slide 3 (Lampiran Teknis).');
} else {
    console.log('Target not found.');
}
