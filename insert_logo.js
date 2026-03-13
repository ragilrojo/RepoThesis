const fs = require('fs');
const path = require('path');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const ImageModule = require('docxtemplater-image-module-free');

async function processDocument() {
    console.log("Memulai proses penyuntikan gambar logo dan kerangka kerja...");
    const inputPath = path.resolve(__dirname, 'proposal_tesis_ragil.docx');
    
    if (!fs.existsSync(inputPath)) {
        console.error("File docx tidak ditemukan!");
        return;
    }

    const content = fs.readFileSync(inputPath, 'binary');
    const zip = new PizZip(content);

    const imageOpts = {
        centered: true,
        getImage: function(tagValue) {
            return fs.readFileSync(tagValue);
        },
        getSize: function(img, tagValue, tagName) {
            if (tagName.includes("LOGO")) return [200, 200];
            if (tagName.includes("FRAMEWORK")) return [550, 350];
            return [300, 300];
        }
    };

    const doc = new Docxtemplater(zip, {
        modules: [new ImageModule(imageOpts)],
        delimiters: { start: '[[', end: ']]' }
    });

    try {
        const renderData = {};

        // Hanya menangani gambar fisik
        const logoPath = path.resolve(__dirname, 'logo_unm.png');
        const frameworkPath = path.resolve(__dirname, 'framwrok.jpg');

        if (fs.existsSync(logoPath)) {
            renderData["LOGO_UNM"] = logoPath;
            console.log("✓ Logo UNM siap.");
        }
        if (fs.existsSync(frameworkPath)) {
            renderData["IMAGE_FRAMEWORK"] = frameworkPath;
            console.log("✓ Gambar Kerangka Kerja siap.");
        }

        console.log("Menyuntikkan gambar ke dokumen...");
        doc.render(renderData);

        const buf = doc.getZip().generate({ type: 'nodebuffer' });
        fs.writeFileSync(inputPath, buf);

        console.log("-----------------------------------------");
        console.log("SUKSES! Logo dan gambar kerangka kerja telah terpasang.");
        console.log("Rumus kini menggunakan format native Word (bukan gambar).");
        console.log("-----------------------------------------");

    } catch (error) {
        console.error("Gagal memproses dokumen:", error.message);
    }
}

processDocument();
