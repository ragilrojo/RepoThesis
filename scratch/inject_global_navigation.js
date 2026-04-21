const fs = require('fs');
const path = 'e:\\ProjectNodeJs\\temp_doc_build\\generate_ppt.js';
let content = fs.readFileSync(path, 'utf8');

// Match any slide variable except slide1 (judu)
const pattern = /let ([a-zA-Z0-9]+) = pres\.addSlide\(\);/g;

const newContent = content.replace(pattern, (match, slideVar) => {
    if (slideVar === 'slide1') return match; // Skip title slide

    // Check if the slide already has navigation code to avoid double injection
    if (content.includes(`${slideVar}.addText("🏠 Daftar Isi"`) || content.includes(`${slideVar}.addText("📂 Lampiran"`)) {
        // If it already has one of them, we might want to ensure it has both or skip
        // But for safety, I will skip slides that already have navigation logic
        return match;
    }

    // Inject navigation links at the end of the slide creation (or just after addSlide)
    // Actually, adding them right after addSlide is fine in terms of layer order.
    return `${match}\n    ${slideVar}.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });\n    ${slideVar}.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });`;
});

fs.writeFileSync(path, newContent);
console.log('Applied global navigation links to all slides except the title slide.');
