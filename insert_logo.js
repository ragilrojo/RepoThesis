const fs = require('fs');
const path = require('path');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const ImageModule = require('docxtemplater-image-module-free');

/**
 * Script untuk mengisi logo pada placeholder [ PLACEHOLDER: LOGO UNIVERSITAS NUSA MANDIRI ]
 */

// 1. Baca file docx sebagai binary
const inputPath = path.join(__dirname, 'proposal_tesis_ragil.docx');
if (!fs.existsSync(inputPath)) {
    console.error("File proposal_tesis_ragil.docx tidak ditemukan!");
    process.exit(1);
}

const content = fs.readFileSync(inputPath, 'binary');
const zip = new PizZip(content);

// 2. Konfigurasi Modul Gambar
const imageOpts = {
    centered: true,
    getImage: function(tagValue) {
        // tagValue akan berisi path gambar yang dikirim dari render data
        return fs.readFileSync(tagValue);
    },
    getSize: function() {
        // Ukuran logo dalam pixel [lebar, tinggi]
        return [200, 200];
    }
};

const doc = new Docxtemplater(zip, {
    modules: [new ImageModule(imageOpts)],
    delimiters: {
        start: '[ ', // Sesuai dengan format placeholder yang kita buat
        end: ' ]'
    }
});

// 3. Proses Penggantian
try {
    const logoPath = path.join(__dirname, 'logo_unm.png');
    if (!fs.existsSync(logoPath)) {
        throw new Error("File logo_unm.png tidak ditemukan!");
    }

    doc.render({
        "PLACEHOLDER: LOGO UNIVERSITAS NUSA MANDIRI": logoPath
    });

    // 4. Timpa file lama dengan yang baru
    const buf = doc.getZip().generate({ type: 'nodebuffer' });
    fs.writeFileSync(inputPath, buf);

    console.log('--------------------------------------------------');
    console.log('Sukses! Logo telah dimasukkan dan file diperbarui.');
    console.log('File diperbarui: proposal_tesis_ragil.docx');
    console.log('--------------------------------------------------');

} catch (error) {
    console.error("Terjadi kesalahan saat memproses dokumen:");
    console.error(error);
}
