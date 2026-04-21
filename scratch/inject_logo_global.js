const fs = require('fs');
const path = 'e:\\ProjectNodeJs\\temp_doc_build\\generate_ppt.js';
let content = fs.readFileSync(path, 'utf8');

// Match any slide creation variable: let varName = pres.addSlide();
// and inject the logo line immediately after it.
const pattern = /let ([a-zA-Z0-9]+) = pres\.addSlide\(\);/g;

const newContent = content.replace(pattern, (match, slideVar) => {
    // Inject addImage after slide creation
    return `${match}\n    ${slideVar}.addImage({ path: "logo_unm.png", x: 0.1, y: 0.1, w: 0.6, h: 0.6 });`;
});

fs.writeFileSync(path, newContent);
console.log('Injected logo_unm.png to all slides.');
