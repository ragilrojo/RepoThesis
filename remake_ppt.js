const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

async function createPresentation() {
    console.log("Memulai pembuatan presentasi (Remake Slide Cover & Slide 2)...");

    // Inisialisasi presentasi baru
    let pres = new pptxgen();

    // Set layout (16x9)
    pres.layout = "LAYOUT_16x9";

    // --- Slide 1: Judul (Cover) ---
    let slide1 = pres.addSlide();
    // Pastikan file gambar ini tersedia di direktori yang sama
    if (fs.existsSync("bg_watermark.png")) {
        slide1.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    }
    if (fs.existsSync("logo_unm.png")) {
        slide1.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    }

    slide1.addText("PROPOSAL TESIS", { x: 0.5, y: 0.7, w: "90%", fontSize: 26, bold: true, color: "003366", align: "center" });
    slide1.addText([
        { text: "Optimasi Portofolio Aset " },
        { text: "Cryptocurrency", options: { italic: true } },
        { text: " Menggunakan " },
        { text: "Network Markowitz", options: { italic: true } },
        { text: " Berbasis SAC (" },
        { text: "Soft Actor-Critic", options: { italic: true } },
        { text: ")" }
    ], {
        x: 0.5, y: 1.5, w: "90%", h: 2.0, fontSize: 30, bold: true, color: "003366", align: "center"
    });

    slide1.addText([
        { text: "Nama: Ragil Yulianto\n" },
        { text: "NIM: 14240007\n" },
        { text: "Program Studi: Informatika (S2)\n" },
        { text: "Pembimbing: Dr. Muhammad Haris, M. Eng." }
    ], {
        x: 0.5, y: 4.2, w: "90%", fontSize: 18, align: "center", color: "7f8c8d"
    });

    slide1.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Slide 2: Daftar Isi ---
    let slideTOC = pres.addSlide();
    if (fs.existsSync("bg_watermark.png")) {
        slideTOC.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    }
    if (fs.existsSync("logo_unm.png")) {
        slideTOC.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    }
    slideTOC.addText("Daftar Isi", { x: 0.5, y: 0.5, w: "90%", fontSize: 28, bold: true, color: "003366" });

    slideTOC.addText([
        { text: "I. PENDAHULUAN & MASALAH", options: { bold: true, color: "003366", breakLine: true } },
        { text: "   • ", options: { fontSize: 14 } },
        { text: "Latar Belakang", options: { fontSize: 14, color: "0563C1", underline: true, hyperlink: { slide: '3' } } },
        { text: "\n", options: { breakLine: true } },
        { text: "   • Rumusan Masalah\n", options: { fontSize: 14 } },
        { text: "   • Tujuan Penelitian\n\n", options: { fontSize: 14 } },

        { text: "II. LANDASAN TEORI & PENELITIAN TERDAHULU", options: { bold: true, color: "003366", breakLine: true } },
        { text: "   • Landasan Teori\n", options: { fontSize: 14 } },
        { text: "   • Kerangka Pemikiran\n\n", options: { fontSize: 14 } },

        { text: "III. METODOLOGI & SAC-NET", options: { bold: true, color: "003366", breakLine: true } },
        { text: "   • Mekanisme SAC Agent\n", options: { fontSize: 14 } },
        { text: "   • Parameter Eksperimen\n\n", options: { fontSize: 14 } },

        { text: "IV. EVALUASI & HASIL", options: { bold: true, color: "003366", breakLine: true } },
        { text: "   • Analisis Performa\n", options: { fontSize: 14 } }
    ], { x: 0.5, y: 1.2, w: "90%", h: 4.0, color: "333333", valign: "top" });

    // --- Slide 3: Latar Belakang (Remake) ---
    let slide2 = pres.addSlide();
    if (fs.existsSync("bg_watermark.png")) {
        slide2.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    }
    if (fs.existsSync("logo_unm.png")) {
        slide2.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    }

    slide2.addText("Latar Belakang", { x: 0.5, y: 0.5, w: "90%", fontSize: 28, bold: true, color: "003366" });

    // Area Kiri: Masalah
    slide2.addShape(pres.ShapeType.rect, { x: 0.5, y: 1.5, w: 4.2, h: 3.2, fill: { color: "ffffff" }, line: { color: "003366", width: 2 } });
    slide2.addText("🔍 IDENTIFIKASI MASALAH", { x: 0.5, y: 1.8, w: 4.2, fontSize: 18, bold: true, color: "003366", align: "center" });
    slide2.addText([
        { text: "• Volatilitas Kripto & Noise:", options: { bold: true, fontSize: 13, breakLine: true } },
        { text: "   Gangguan data yang mengaburkan sinyal asli.", options: { fontSize: 12, breakLine: true } },
        { text: "   Ex: Price spikes akibat sentimen sesaat vs korelasi fundamental.", options: { fontSize: 11, italic: true, color: "666666", breakLine: true } },
        { text: "• Optimalitas Window:", options: { bold: true, fontSize: 13, breakLine: true } },
        { text: "   Belum ada standar panjang jendela observasi.", options: { fontSize: 12, breakLine: true } },
        { text: "   Ex: Window 30 hari (sensitif) vs 90 hari (lambat) memberikan alokasi berlawanan.", options: { fontSize: 11, italic: true, color: "666666" } }
    ], { x: 0.7, y: 2.5, w: 3.8, color: "333333", valign: "top" });

    // Area Kanan: Gap & Solusi
    slide2.addShape(pres.ShapeType.rect, { x: 5.3, y: 1.5, w: 4.2, h: 3.2, fill: { color: "ffffff" }, line: { color: "27ae60", width: 2 } });
    slide2.addText("💡 RESEARCH GAP & SOLUSI", { x: 5.3, y: 1.8, w: 4.2, fontSize: 18, bold: true, color: "27ae60", align: "center" });
    slide2.addText([
        { text: "• Keterbatasan Penalti (\u03b3) Statis:", options: { bold: true, fontSize: 13, breakLine: true } },
        { text: "   Model kaku (Giudici, 2020) menghambat profit saat tren kuat.", options: { fontSize: 12, breakLine: true } },
        { text: "   Ex: \u03b3 tinggi terus-menerus memangkas alfa di pasar bullish.\n", options: { fontSize: 11, italic: true, color: "666666", breakLine: true } },
        { text: "• Solusi: SAC-Based Gamma Controller:", options: { bold: true, fontSize: 13, color: "27ae60", breakLine: true } },
        { text: "   Deep RL Agent untuk penyesuaian \u03b3 dinamis secara real-time.", options: { italic: true, fontSize: 11 } }
    ], { x: 5.5, y: 2.5, w: 3.8, color: "333333", valign: "top" });

    // Footer Navigasi Sederhana
    slide2.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Simpan File ---
    const fileName = "Remake_Proposal_Tesis.pptx";
    pres.writeFile({ fileName: fileName })
        .then(fileName => {
            console.log(`Presentasi berhasil dibuat: ${fileName}`);
        })
        .catch(err => {
            console.error("Gagal membuat presentasi:", err);
        });
}

createPresentation();
