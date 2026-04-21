const fs = require('fs'); 
const lines = fs.readFileSync('generate_ppt.js', 'utf8').split('\n'); 
let slideCount = 0; 
lines.forEach((line, i) => { 
    if (line.includes('pres.addSlide()')) { 
        slideCount++; 
        console.log('Slide ' + slideCount + ' at line ' + (i+1)); 
        let titleLine = lines.slice(i, i+15).find(l => l.includes('addText(') && !l.includes('Lampiran') && !l.includes('Daftar Isi')); 
        if (titleLine) console.log('  -> ' + titleLine.trim()); 
    } 
});
