const fs = require('fs');

const width = 800;
const height = 400;
const padding = 60;

const svg = `
<svg width="${width}" height="${height}" viewBox="0 0 ${width} ${height}" xmlns="http://www.w3.org/2000/svg">
  <!-- Background -->
  <rect width="100%" height="100%" fill="white"/>
  
  <!-- Gridlines -->
  <line x1="${padding}" y1="${height-padding-80}" x2="${width-padding}" y2="${height-padding-80}" stroke="#eee" stroke-width="1"/>
  <line x1="${padding}" y1="${height-padding-160}" x2="${width-padding}" y2="${height-padding-160}" stroke="#eee" stroke-width="1"/>
  <line x1="${padding}" y1="${height-padding-240}" x2="${width-padding}" y2="${height-padding-240}" stroke="#eee" stroke-width="1"/>

  <!-- Axes -->
  <line x1="${padding}" y1="${height-padding}" x2="${width-padding}" y2="${height-padding}" stroke="#333" stroke-width="2"/>
  <line x1="${padding}" y1="${padding}" x2="${padding}" y2="${height-padding}" stroke="#333" stroke-width="2"/>

  <!-- Path: Simulated Convergence Curve -->
  <!-- Start: (60, 320) -> (200, 330) -> (400, 150) -> (740, 100) -->
  <path d="M 60 320 
           L 80 280 L 100 310 L 120 270 L 140 290 L 160 250 L 180 300 L 200 260
           C 300 250, 350 150, 450 120
           S 600 100, 740 105" 
        fill="none" stroke="url(#lineGradient)" stroke-width="4" stroke-linecap="round" stroke-linejoin="round"/>

  <!-- Labels -->
  <text x="${width/2}" y="${height-10}" font-family="Arial" font-size="14" text-anchor="middle" fill="#666">Training Steps (50,000)</text>
  <text x="20" y="${height/2}" font-family="Arial" font-size="14" text-anchor="middle" fill="#666" transform="rotate(-90 20,${height/2})">Mean Reward</text>
  
  <text x="130" y="360" font-family="Arial" font-size="12" font-style="italic" fill="#2980b9">Exploration Phase</text>
  <text x="550" y="80" font-family="Arial" font-size="12" font-weight="bold" fill="#27ae60">Convergence Phase</text>

  <!-- Definitions -->
  <defs>
    <linearGradient id="lineGradient" x1="0%" y1="0%" x2="100%" y2="0%">
      <stop offset="0%" style="stop-color:#3498db;stop-opacity:1" />
      <stop offset="100%" style="stop-color:#27ae60;stop-opacity:1" />
    </linearGradient>
  </defs>
</svg>
`;

fs.writeFileSync('reward_convergence.svg', svg);
console.log('SVG file "reward_convergence.svg" created successfully.');
