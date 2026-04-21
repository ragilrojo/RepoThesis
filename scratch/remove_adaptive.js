const fs = require('fs');
const path = 'e:\\ProjectNodeJs\\temp_doc_build\\generate_ppt.js';
let content = fs.readFileSync(path, 'utf8');

// Replacements
let newContent = content
    .replace(/Optimalisasi Portofolio Adaptif/g, 'Optimalisasi Portofolio')
    .replace(/Statis vs Adaptif/g, 'Statis vs Tuned')
    .replace(/Adaptif secara otomatis/g, 'melalui Tuning Parameter')
    .replace(/model Adaptif \(Dynamic Gamma\)/g, 'model dengan Tuning Parameter')
    .replace(/Kesimpulan Adaptif:/g, 'Kesimpulan Tuning:')
    .replace(/parameter adaptif/g, 'parameter hasil tuning')
    .replace(/Sifat adaptif/g, 'Hasil tuning')
    .replace(/Model Adaptif/g, 'Model Tuned');

fs.writeFileSync(path, newContent);
console.log('Removed Adaptive references and replaced with Tuning terms.');
