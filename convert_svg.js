const svg2img = require('svg2img');
const fs = require('fs');
const path = require('path');

const svgPath = path.join(__dirname, 'financial_network_comparison.svg');
const pngPath = path.join(__dirname, 'financial_network_comparison.png');

console.log('Converting SVG to PNG...');

svg2img(svgPath, { width: 1000, height: 500 }, function(error, buffer) {
    if (error) {
        console.error('Error converting SVG:', error);
        process.exit(1);
    }
    fs.writeFileSync(pngPath, buffer);
    console.log('Successfully created:', pngPath);
    process.exit(0);
});
