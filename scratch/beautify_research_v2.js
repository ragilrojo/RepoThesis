const fs = require('fs');
const path = 'e:\\ProjectNodeJs\\temp_doc_build\\generate_ppt.js';
let content = fs.readFileSync(path, 'utf8');

const newTableCode = `    slidePrev.addTable(
        [
            [
                { text: "Peneliti", options: { bold: true, fill: "003366", color: "ffffff" } },
                { text: "Tahun", options: { bold: true, fill: "003366", color: "ffffff" } },
                { text: "Kontribusi Utama", options: { bold: true, fill: "003366", color: "ffffff" } }
            ],
            ["Giudici et al.", "2020", "Pelopor Network Markowitz (RMT & MST) di Kripto"],
            ["Kitanovski et al.", "2022", "Diversifikasi berbasis Deteksi Komunitas Jaringan"],
            ["Jing & Rocha", "2023", "Topologi MST mengalahkan semua benchmark"],
            ["Kitanovski et al.", "2024", "Stabilitas Penalti Graf pada Eksposur Ekstrem"],
            ["Jing et al.", "2025", "Prediksi Stabil fase terbaru (Network-MPT)"]
        ],
        { x: 0.5, y: 1.2, w: 9.0, rowH: 0.6, fontSize: 13, border: { pt: 1, color: "dddddd" }, align: "center", valign: "middle" }
    );

    slidePrev.addText("Fokus Penelitian Kami: Optimalisasi Parameter secara Sistematis", { 
        x: 0.5, y: 5.0, w: "90%", fontSize: 14, bold: true, italic: true, color: "27ae60", align: "center" 
    });`;

// Regex for slidePrev content
const blockPattern = /slidePrev\.addText\(\[\n\s+\{ text: "1\. Giudici et al\. \(2020\):"[\s\S]*?di fase terbaru\.", options: \{ bold: true \} \}\n\s+\][\s\S]*?\);/;

let newContent = content.replace(blockPattern, newTableCode);

fs.writeFileSync(path, newContent);
console.log('Restructured slidePrev (Previous Research) into a professional table.');
