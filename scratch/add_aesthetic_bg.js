const fs = require('fs');
const path = 'e:\\ProjectNodeJs\\temp_doc_build\\generate_ppt.js';
let content = fs.readFileSync(path, 'utf8');

// Pattern to find slide variable and logo injection
// Current: let slideVar = pres.addSlide();
//          slideVar.addImage({ path: "logo_unm.png", ... });

// New: let slideVar = pres.addSlide();
//      slideVar.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
//      slideVar.addImage({ path: "logo_unm.png", ... });

const pattern = /let ([a-zA-Z0-9]+) = pres\.addSlide\(\);\n    ([a-zA-Z0-9]+)\.addImage\(\{ path: "logo_unm\.png"/g;

const newContent = content.replace(pattern, (match, slideVar, slideVarAgain) => {
    return `let ${slideVar} = pres.addSlide();\n    ${slideVar}.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });\n    ${slideVar}.addImage({ path: "logo_unm.png"`;
});

fs.writeFileSync(path, newContent);
console.log('Added aesthetic silhouette background to all slides.');
