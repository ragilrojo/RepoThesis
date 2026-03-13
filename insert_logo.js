const fs = require('fs');
const path = require('path');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const ImageModule = require('docxtemplater-image-module-free');

/**
 * Script untuk mengisi placeholder gambar [ %PLACEHOLDER: ... ] di file docx
 */

const inputPath = path.join(__dirname, 'proposal_tesis_ragil.docx');
if (!fs.existsSync(inputPath)) {
    console.error("File proposal_tesis_ragil.docx tidak ditemukan!");
    process.exit(1);
}

const content = fs.readFileSync(inputPath, 'binary');
const zip = new PizZip(content);

// Konfigurasi Modul Gambar
const imageOpts = {
    centered: true,
    getImage: function(tagValue) {
        return fs.readFileSync(tagValue);
    },
    getSize: function(img, tagValue, tagName) {
        // Tentukan ukuran berdasarkan nama tag
        if (tagName.includes("LOGO")) {
            return [200, 200];
        }
        if (tagName.includes("KERANGKA")) {
            return [550, 350]; // Sesuai ukuran framwrok.jpg sebelumnya
        }
        return [300, 300]; // Default
    }
};

const doc = new Docxtemplater(zip, {
    modules: [new ImageModule(imageOpts)],
    delimiters: {
        start: '[ ',
        end: ' ]'
    }
});

try {
    const logoPath = path.join(__dirname, 'logo_unm.png');
    const frameworkPath = path.join(__dirname, 'framwrok.jpg');

    const renderData = {};

    if (fs.existsSync(logoPath)) {
        renderData["PLACEHOLDER: LOGO UNIVERSITAS NUSA MANDIRI"] = logoPath;
    }
    
    if (fs.existsSync(frameworkPath)) {
        renderData["PLACEHOLDER: GAMBAR KERANGKA KERJA PENELITIAN"] = frameworkPath;
    }

    doc.render(renderData);

    const buf = doc.getZip().generate({ type: 'nodebuffer' });
    fs.writeFileSync(inputPath, buf);

    console.log('--------------------------------------------------');
    console.log('Sukses! Semua gambar placeholder telah diproses.');
    console.log('File diperbarui: proposal_tesis_ragil.docx');
    console.log('--------------------------------------------------');

} catch (error) {
    console.error("Terjadi kesalahan saat memproses dokumen:");
    console.error(error);
}
