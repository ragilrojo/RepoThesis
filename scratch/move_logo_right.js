const fs = require('fs');
const path = 'e:\\ProjectNodeJs\\temp_doc_build\\generate_ppt.js';
let content = fs.readFileSync(path, 'utf8');

// Replace the top-left logo coordinates with top-right (x: 9.1)
const pattern = /addImage\(\{ path: "logo_unm\.png", x: 0\.1, y: 0\.1, w: 0\.6, h: 0\.6 \}\)/g;

const newContent = content.replace(pattern, 'addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 })');

fs.writeFileSync(path, newContent);
console.log('Moved logo_unm.png to top-right corner.');
