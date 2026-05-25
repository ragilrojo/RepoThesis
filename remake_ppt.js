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

    // --- Slide 2: Daftar Isi (PREMIUM REDESIGN) ---
    let slideTOC = pres.addSlide();
    if (fs.existsSync("bg_watermark.png")) {
        slideTOC.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    }
    if (fs.existsSync("logo_unm.png")) {
        slideTOC.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    }
    
    // Judul Slide dengan Garis Bawah Aksen
    slideTOC.addText("Daftar Isi", { x: 0.5, y: 0.4, w: 3.0, fontSize: 32, bold: true, color: "003366" });
    slideTOC.addShape(pres.ShapeType.line, { x: 0.5, y: 0.9, w: 1.5, h: 0, line: { color: "e67e22", width: 3 } });

    // Definisi Grid dan Dimensi Card
    const cardW = 4.4;
    const col1X = 0.5;
    const col2X = 5.1;
    const row1Y = 1.2;
    const row2Y = 3.3;
    const headerH = 0.4;

    // --- CARD I: PENDAHULUAN ---
    slideTOC.addShape(pres.ShapeType.rect, { x: col1X, y: row1Y, w: cardW, h: 1.8, fill: { color: "ffffff" }, line: { color: "003366", width: 1.5 }, rectRadius: 0.05 });
    slideTOC.addShape(pres.ShapeType.rect, { x: col1X, y: row1Y, w: cardW, h: headerH, fill: { color: "003366" } });
    slideTOC.addText("I. PENDAHULUAN & MASALAH", { x: col1X, y: row1Y, w: cardW, h: headerH, fontSize: 14, bold: true, color: "ffffff", align: "center", valign: "middle" });
    slideTOC.addText([
        { text: "Latar Belakang", options: { bullet: { code: '25CF' }, color: "0563C1", underline: true, hyperlink: { slide: '3' } } },
        { text: "Rumusan Masalah", options: { bullet: { code: '25CF' }, color: "0563C1", underline: true, hyperlink: { slide: '4' } } },
        { text: "Tujuan Penelitian", options: { bullet: { code: '25CF' }, color: "0563C1", underline: true, hyperlink: { slide: '5' } } }
    ], { x: col1X + 0.3, y: row1Y + 0.5, w: cardW - 0.6, h: 1.2, fontSize: 11, lineSpacing: 22, valign: "top" });

    // --- CARD II: LANDASAN ---
    slideTOC.addShape(pres.ShapeType.rect, { x: col1X, y: row2Y, w: cardW, h: 1.8, fill: { color: "ffffff" }, line: { color: "27ae60", width: 1.5 }, rectRadius: 0.05 });
    slideTOC.addShape(pres.ShapeType.rect, { x: col1X, y: row2Y, w: cardW, h: headerH, fill: { color: "27ae60" } });
    slideTOC.addText("II. LANDASAN & SIMULASI", { x: col1X, y: row2Y, w: cardW, h: headerH, fontSize: 14, bold: true, color: "ffffff", align: "center", valign: "middle" });
    slideTOC.addText([
        { text: "Z-Score Normalisasi", options: { bullet: { code: '25CF' }, color: "0563C1", underline: true, hyperlink: { slide: '6' } } },
        { text: "Landasan Teori (Markowitz)", options: { bullet: { code: '25CF' }, color: "0563C1", underline: true, hyperlink: { slide: '7' } } },
        { text: "Kerangka Pemikiran", options: { bullet: { code: '25CF' }, color: "333333" } }
    ], { x: col1X + 0.3, y: row2Y + 0.45, w: cardW - 0.6, h: 1.3, fontSize: 10.5, lineSpacing: 18, valign: "top" });

    // --- CARD III: METODOLOGI ---
    slideTOC.addShape(pres.ShapeType.rect, { x: col2X, y: row1Y, w: cardW, h: 1.8, fill: { color: "ffffff" }, line: { color: "2980b9", width: 1.5 }, rectRadius: 0.05 });
    slideTOC.addShape(pres.ShapeType.rect, { x: col2X, y: row1Y, w: cardW, h: headerH, fill: { color: "2980b9" } });
    slideTOC.addText("III. METODOLOGI & SAC-NET", { x: col2X, y: row1Y, w: cardW, h: headerH, fontSize: 14, bold: true, color: "ffffff", align: "center", valign: "middle" });
    slideTOC.addText([
        { text: "Network Markowitz", options: { bullet: { code: '25CF' }, color: "0563C1", underline: true, hyperlink: { slide: '8' } } },
        { text: "Fitur Observasi SAC", options: { bullet: { code: '25CF' }, color: "0563C1", underline: true, hyperlink: { slide: '9' } } },
        { text: "Simulasi Hitung Fitur", options: { bullet: { code: '25CF' }, color: "0563C1", underline: true, hyperlink: { slide: '18' } } },
        { text: "Ilustrasi Feature Scaling", options: { bullet: { code: '25CF' }, color: "0563C1", underline: true, hyperlink: { slide: '19' } } }
    ], { x: col2X + 0.3, y: row1Y + 0.45, w: cardW - 0.6, h: 1.3, fontSize: 10.5, lineSpacing: 18, valign: "top" });

    // --- CARD IV: EVALUASI ---
    slideTOC.addShape(pres.ShapeType.rect, { x: col2X, y: row2Y, w: cardW, h: 1.8, fill: { color: "ffffff" }, line: { color: "8e44ad", width: 1.5 }, rectRadius: 0.05 });
    slideTOC.addShape(pres.ShapeType.rect, { x: col2X, y: row2Y, w: cardW, h: headerH, fill: { color: "8e44ad" } });
    slideTOC.addText("IV. EVALUASI & HASIL", { x: col2X, y: row2Y, w: cardW, h: headerH, fontSize: 14, bold: true, color: "ffffff", align: "center", valign: "middle" });
    slideTOC.addText([
        { text: "Evaluasi Portofolio (Ratio)", options: { bullet: { code: '25CF' }, color: "0563C1", underline: true, hyperlink: { slide: '20' } } },
        { text: "Analisis Performa & Visualisasi", options: { bullet: { code: '25CF' }, color: "333333" } },
        { text: "Kesimpulan & Saran", options: { bullet: { code: '25CF' }, color: "333333" } }
    ], { x: col2X + 0.3, y: row2Y + 0.5, w: cardW - 0.6, h: 1.2, fontSize: 11, lineSpacing: 22, valign: "top" });




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

    // --- Slide 4: Rumusan Masalah Penelitian ---
    let slideProb = pres.addSlide();
    if (fs.existsSync("bg_watermark.png")) {
        slideProb.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    }
    if (fs.existsSync("logo_unm.png")) {
        slideProb.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    }

    slideProb.addText("Rumusan Masalah Penelitian", { x: 0.5, y: 0.4, w: "90%", fontSize: 32, bold: true, color: "003366" });
    slideProb.addShape(pres.ShapeType.line, { x: 0.5, y: 0.9, w: 1.5, h: 0, line: { color: "e67e22", width: 3 } });

    const cardStartX = 0.5;
    const cardStartY = 1.3;
    const cardGap = 1.2;
    const sideW = 0.8;
    const bodyW = 8.5;
    const cardH = 1.0;

    // Q1
    slideProb.addShape(pres.ShapeType.rect, { x: cardStartX, y: cardStartY, w: sideW + bodyW, h: cardH, fill: { color: "ffffff" }, line: { color: "3498db", width: 1.5 } });
    slideProb.addShape(pres.ShapeType.rect, { x: cardStartX, y: cardStartY, w: sideW, h: cardH, fill: { color: "3498db" } });
    slideProb.addText("Q1", { x: cardStartX, y: cardStartY, w: sideW, h: cardH, fontSize: 22, bold: true, color: "ffffff", align: "center", valign: "middle" });
    slideProb.addText("Bagaimana efektivitas algoritma Soft Actor-Critic (SAC) dalam mengendalikan parameter penalti sentralitas secara dinamis pada model Network-Markowitz?", { x: cardStartX + sideW + 0.2, y: cardStartY, w: bodyW - 0.4, h: cardH, fontSize: 13, bold: true, color: "2c3e50", valign: "middle" });

    // Q2
    slideProb.addShape(pres.ShapeType.rect, { x: cardStartX, y: cardStartY + cardGap, w: sideW + bodyW, h: cardH, fill: { color: "ffffff" }, line: { color: "2c3e50", width: 1.5 } });
    slideProb.addShape(pres.ShapeType.rect, { x: cardStartX, y: cardStartY + cardGap, w: sideW, h: cardH, fill: { color: "2c3e50" } });
    slideProb.addText("Q2", { x: cardStartX, y: cardStartY + cardGap, w: sideW, h: cardH, fontSize: 22, bold: true, color: "ffffff", align: "center", valign: "middle" });
    slideProb.addText("Bagaimana metode Explainable AI (SHAP) dapat menjelaskan pengaruh fitur Network dan Market terhadap keputusan agen dalam alokasi aset portofolio?", { x: cardStartX + sideW + 0.2, y: cardStartY + cardGap, w: bodyW - 0.4, h: cardH, fontSize: 13, bold: true, color: "2c3e50", valign: "middle" });

    // Q3
    slideProb.addShape(pres.ShapeType.rect, { x: cardStartX, y: cardStartY + (cardGap * 2), w: sideW + bodyW, h: cardH, fill: { color: "ffffff" }, line: { color: "e67e22", width: 1.5 } });
    slideProb.addShape(pres.ShapeType.rect, { x: cardStartX, y: cardStartY + (cardGap * 2), w: sideW, h: cardH, fill: { color: "e67e22" } });
    slideProb.addText("Q3", { x: cardStartX, y: cardStartY + (cardGap * 2), w: sideW, h: cardH, fontSize: 22, bold: true, color: "ffffff", align: "center", valign: "middle" });
    slideProb.addText("Apakah model SAC-Net Markowitz menunjukkan performa yang unggul secara signifikan dibandingkan benchmark berdasarkan metrik Sharpe, Sortino, Calmar, dan Ulcer Index?", { x: cardStartX + sideW + 0.2, y: cardStartY + (cardGap * 2), w: bodyW - 0.4, h: cardH, fontSize: 13, bold: true, color: "2c3e50", valign: "middle" });

    slideProb.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });


    // --- Slide 5: Tujuan Penelitian ---
    let slideTujuan = pres.addSlide();
    if (fs.existsSync("bg_watermark.png")) {
        slideTujuan.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    }
    if (fs.existsSync("logo_unm.png")) {
        slideTujuan.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    }

    slideTujuan.addText("Tujuan Penelitian", { x: 0.5, y: 0.4, w: "90%", fontSize: 32, bold: true, color: "003366" });
    slideTujuan.addShape(pres.ShapeType.line, { x: 0.5, y: 0.9, w: 1.5, h: 0, line: { color: "e67e22", width: 3 } });

    const tCardStartX = 0.5;
    const tCardStartY = 1.3;
    const tCardGap = 1.2;
    const tSideW = 0.8;
    const tBodyW = 8.5;
    const tCardH = 1.0;

    // Obj 1
    slideTujuan.addShape(pres.ShapeType.rect, { x: tCardStartX, y: tCardStartY, w: tSideW + tBodyW, h: tCardH, fill: { color: "ffffff" }, line: { color: "27ae60", width: 1.5 } });
    slideTujuan.addShape(pres.ShapeType.rect, { x: tCardStartX, y: tCardStartY, w: tSideW, h: tCardH, fill: { color: "27ae60" } });
    slideTujuan.addText("Obj 1", { x: tCardStartX, y: tCardStartY, w: tSideW, h: tCardH, fontSize: 22, bold: true, color: "ffffff", align: "center", valign: "middle" });
    slideTujuan.addText("Merancang dan melatih agen Soft Actor-Critic (SAC) sebagai Gamma Controller dinamis untuk mengoptimasi parameter penalti sentralitas dalam kerangka Network-Markowitz.", { x: tCardStartX + tSideW + 0.2, y: tCardStartY, w: tBodyW - 0.4, h: tCardH, fontSize: 13, bold: true, color: "2c3e50", valign: "middle" });

    // Obj 2
    slideTujuan.addShape(pres.ShapeType.rect, { x: tCardStartX, y: tCardStartY + tCardGap, w: tSideW + tBodyW, h: tCardH, fill: { color: "ffffff" }, line: { color: "2c3e50", width: 1.5 } });
    slideTujuan.addShape(pres.ShapeType.rect, { x: tCardStartX, y: tCardStartY + tCardGap, w: tSideW, h: tCardH, fill: { color: "2c3e50" } });
    slideTujuan.addText("Obj 2", { x: tCardStartX, y: tCardStartY + tCardGap, w: tSideW, h: tCardH, fontSize: 22, bold: true, color: "ffffff", align: "center", valign: "middle" });
    slideTujuan.addText("Menganalisis kontribusi fitur gabungan (Network & Market Indicators) menggunakan metode Explainable AI (SHAP) untuk transparansi keputusan model.", { x: tCardStartX + tSideW + 0.2, y: tCardStartY + tCardGap, w: tBodyW - 0.4, h: tCardH, fontSize: 13, bold: true, color: "2c3e50", valign: "middle" });

    // Obj 3
    slideTujuan.addShape(pres.ShapeType.rect, { x: tCardStartX, y: tCardStartY + (tCardGap * 2), w: tSideW + tBodyW, h: tCardH, fill: { color: "ffffff" }, line: { color: "f39c12", width: 1.5 } });
    slideTujuan.addShape(pres.ShapeType.rect, { x: tCardStartX, y: tCardStartY + (tCardGap * 2), w: tSideW, h: tCardH, fill: { color: "f39c12" } });
    slideTujuan.addText("Obj 3", { x: tCardStartX, y: tCardStartY + (tCardGap * 2), w: tSideW, h: tCardH, fontSize: 22, bold: true, color: "ffffff", align: "center", valign: "middle" });
    slideTujuan.addText("Menguji performa model secara statistik (Wilcoxon Test) dibandingkan benchmark melalui metrik evaluasi Sharpe, Sortino, Calmar, dan Ulcer Index.", { x: tCardStartX + tSideW + 0.2, y: tCardStartY + (tCardGap * 2), w: tBodyW - 0.4, h: tCardH, fontSize: 13, bold: true, color: "2c3e50", valign: "middle" });

    slideTujuan.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });
    // --- Slide X: Z-Score Normalisasi (Redesigned with Example) ---
    let slideZ = pres.addSlide();
    if (fs.existsSync("bg_watermark.png")) {
        slideZ.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    }
    if (fs.existsSync("logo_unm.png")) {
        slideZ.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    }

    // Header
    slideZ.addText("Z-Score Normalisasi", { x: 0.5, y: 0.3, w: "85%", fontSize: 26, bold: true, color: "003366" });
    slideZ.addShape(pres.ShapeType.line, { x: 0.5, y: 0.72, w: 1.5, h: 0, line: { color: "e67e22", width: 3 } });
    slideZ.addText("Teknik standarisasi fitur agar neural network belajar secara stabil:", { x: 0.5, y: 0.82, w: "85%", fontSize: 12, color: "555555" });

    // === CARD KIRI: Rumus & Konsep ===
    slideZ.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.15, w: 4.5, h: 3.55, fill: { color: "ffffff" }, line: { color: "003366", width: 1.5 }, rectRadius: 0.05 });
    slideZ.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.15, w: 4.5, h: 0.42, fill: { color: "003366" } });
    slideZ.addText("\ud83d\udcd0 Rumus & Konsep", { x: 0.4, y: 1.15, w: 4.5, h: 0.42, fontSize: 13, bold: true, color: "ffffff", align: "center", valign: "middle" });

    slideZ.addText([
        { text: "Rumus Z-Score:\n", options: { bold: true, fontSize: 12, color: "003366", breakLine: true } },
        { text: "   Z = (x - \u03bc) / \u03c3\n\n", options: { bold: true, fontSize: 16, color: "c0392b", italic: true } },
        { text: "Keterangan:\n", options: { bold: true, fontSize: 11, color: "003366" } },
        { text: "  \u2022 x  = nilai fitur asli (setelah feature scaling)\n", options: { fontSize: 10 } },
        { text: "  \u2022 \u03bc  = rata-rata (mean) seluruh data dalam window\n", options: { fontSize: 10 } },
        { text: "  \u2022 \u03c3  = standar deviasi seluruh data dalam window\n\n", options: { fontSize: 10 } },
        { text: "Tujuan Normalisasi:\n", options: { bold: true, fontSize: 11, color: "27ae60" } },
        { text: "  \u2022 Mengubah fitur ke mean = 0, std = 1\n", options: { fontSize: 10 } },
        { text: "  \u2022 Mencegah fitur bernilai besar mendominasi\n", options: { fontSize: 10 } },
        { text: "  \u2022 Mempercepat konvergensi neural network\n", options: { fontSize: 10 } },
        { text: "  \u2022 Menstabilkan proses backpropagation\n", options: { fontSize: 10 } },
    ], { x: 0.6, y: 1.65, w: 4.1, h: 2.95, color: "333333", valign: "top", lineSpacing: 8 });

    // === CARD KANAN: Contoh Numerik ===
    slideZ.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.15, w: 4.5, h: 3.55, fill: { color: "ffffff" }, line: { color: "2980b9", width: 1.5 }, rectRadius: 0.05 });
    slideZ.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.15, w: 4.5, h: 0.42, fill: { color: "2980b9" } });
    slideZ.addText("\ud83d\udd22 Contoh Numerik (VolShort \u00d7 100)", { x: 5.1, y: 1.15, w: 4.5, h: 0.42, fontSize: 12, bold: true, color: "ffffff", align: "center", valign: "middle" });

    // Tabel data (3 baris data + 2 baris summary = compact)
    slideZ.addTable([
        [
            { text: "Hari", options: { bold: true, fill: "2980b9", color: "ffffff", align: "center", fontSize: 9 } },
            { text: "VolShort (x)", options: { bold: true, fill: "2980b9", color: "ffffff", align: "center", fontSize: 9 } },
            { text: "Z-Score", options: { bold: true, fill: "27ae60", color: "ffffff", align: "center", fontSize: 9 } }
        ],
        [
            { text: "Day 1", options: { align: "center", fontSize: 9 } },
            { text: "1.20", options: { align: "center", fontSize: 9 } },
            { text: "-1.07", options: { align: "center", fontSize: 9, bold: true, color: "c0392b" } }
        ],
        [
            { text: "Day 2", options: { align: "center", fontSize: 9 } },
            { text: "1.80", options: { align: "center", fontSize: 9 } },
            { text: "0.00", options: { align: "center", fontSize: 9 } }
        ],
        [
            { text: "Day 3", options: { align: "center", fontSize: 9 } },
            { text: "2.40", options: { align: "center", fontSize: 9 } },
            { text: "+1.07", options: { align: "center", fontSize: 9, bold: true, color: "27ae60" } }
        ],
        [
            { text: "Mean (\u03bc)", options: { bold: true, fill: "ebf5fb", align: "center", fontSize: 9 } },
            { text: "1.80", options: { bold: true, fill: "ebf5fb", align: "center", fontSize: 9 } },
            { text: "\u2013", options: { fill: "ebf5fb", align: "center", fontSize: 9 } }
        ],
        [
            { text: "Std Dev (\u03c3)", options: { bold: true, fill: "ebf5fb", align: "center", fontSize: 9 } },
            { text: "0.56", options: { bold: true, fill: "ebf5fb", align: "center", fontSize: 9 } },
            { text: "\u2013", options: { fill: "ebf5fb", align: "center", fontSize: 9 } }
        ],
    ], { x: 5.25, y: 1.7, w: 4.2, fontSize: 9, border: { pt: 1, color: "cccccc" }, align: "center", valign: "middle", rowH: 0.27 });

    // Penjabaran contoh — positioned below table (6 rows × 0.27 = 1.62, start 1.7 → ends ~3.32)
    slideZ.addText([
        { text: "Langkah Perhitungan:\n", options: { bold: true, fontSize: 10, color: "003366", breakLine: true } },
        { text: "\u2022 Day 3 (x = 2.40):\n", options: { bold: true, fontSize: 9.5, color: "2c3e50" } },
        { text: "  Z = (2.40 \u2212 1.80) / 0.56 = ", options: { fontSize: 9.5 } },
        { text: "+1.07", options: { bold: true, fontSize: 10, color: "27ae60" } },
        { text: " \u2192 di atas rata-rata\n", options: { fontSize: 8.5, italic: true, color: "666666" } },
        { text: "\u2022 Day 1 (x = 1.20):\n", options: { bold: true, fontSize: 9.5, color: "2c3e50" } },
        { text: "  Z = (1.20 \u2212 1.80) / 0.56 = ", options: { fontSize: 9.5 } },
        { text: "\u22121.07", options: { bold: true, fontSize: 10, color: "c0392b" } },
        { text: " \u2192 di bawah rata-rata", options: { fontSize: 8.5, italic: true, color: "666666" } },
    ], { x: 5.25, y: 3.42, w: 4.2, h: 1.2, color: "333333", valign: "top", lineSpacing: 8 });

    // Summary Box
    slideZ.addShape(pres.ShapeType.rect, { x: 0.4, y: 4.82, w: 9.2, h: 0.42, fill: { color: "eaf4fb" }, line: { color: "2980b9", width: 1.0 }, rectRadius: 0.05 });
    slideZ.addText("Kesimpulan: Z-Score mengubah semua fitur (VolShort, Mom5d, dll.) ke skala seragam (mean=0, \u03c3=1), mencegah dominasi fitur berskala besar dan mempercepat konvergensi SAC Agent.", {
        x: 0.4, y: 4.82, w: 9.2, h: 0.42, fontSize: 10, bold: true, color: "003366", align: "center", valign: "middle"
    });

    slideZ.addText("\ud83c\udfe0 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });


    // --- Slide 4: Simulasi Markowitz Classic ---
    let slideSim = pres.addSlide();
    if (fs.existsSync("bg_watermark.png")) {
        slideSim.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    }
    if (fs.existsSync("logo_unm.png")) {
        slideSim.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    }

    slideSim.addText("Simulasi Markowitz Classic (MVO)", { x: 0.5, y: 0.4, w: "90%", fontSize: 26, bold: true, color: "003366" });
    slideSim.addText("Contoh Alokasi Portofolio Berbasis Mean-Variance Optimization:", { x: 0.5, y: 1.0, w: "90%", fontSize: 14, color: "333333" });

    // Tabel Simulasi (Lebih Ramping)
    slideSim.addTable([
        [
            { text: "Aset", options: { bold: true, fill: "003366", color: "ffffff", align: "center" } },
            { text: "Return", options: { bold: true, fill: "003366", color: "ffffff", align: "center" } },
            { text: "Risk", options: { bold: true, fill: "003366", color: "ffffff", align: "center" } },
            { text: "Weight", options: { bold: true, fill: "003366", color: "ffffff", align: "center" } }
        ],
        ["Bitcoin (BTC)", "20.5%", "45.2%", "40.0%"],
        ["Ethereum (ETH)", "25.8%", "52.1%", "35.0%"],
        ["Binance Coin (BNB)", "18.2%", "38.5%", "25.0%"],
        [
            { text: "Total Portfolio", options: { bold: true, fill: "f4f6f7" } },
            { text: "21.8%", options: { bold: true, fill: "f4f6f7" } },
            { text: "35.4%", options: { bold: true, fill: "f4f6f7" } },
            { text: "100.0%", options: { bold: true, fill: "f4f6f7" } }
        ]
    ], { x: 0.5, y: 1.5, w: 5.0, fontSize: 12, border: { pt: 1, color: "dddddd" }, align: "center", valign: "middle" });

    // Box Penjelasan di Kanan
    slideSim.addShape(pres.ShapeType.rect, { x: 5.8, y: 1.5, w: 3.7, h: 3.5, fill: { color: "fdfef9" }, line: { color: "27ae60", width: 1.5 } });
    slideSim.addText([
        { text: "Langkah Penjabaran Skor:\n", options: { bold: true, color: "003366", fontSize: 13, breakLine: true } },
        { text: "1. Skor BTC: ", options: { fontSize: 11 } }, { text: "20.5/45.2\u00b2 \u2248 1.00\n", options: { bold: true } },
        { text: "2. Skor ETH: ", options: { fontSize: 11 } }, { text: "25.8/52.1\u00b2 \u2248 0.88\n", options: { bold: true } },
        { text: "3. Skor BNB: ", options: { fontSize: 11 } }, { text: "18.2/38.5\u00b2 \u2248 0.62\n", options: { bold: true } },
        { text: "------------------------------------------\n", options: {} },
        { text: "Total Skor: ", options: { fontSize: 11 } }, { text: "1.00 + 0.88 + 0.62 = 2.50\n", options: { bold: true, color: "c0392b" } },
        { text: "------------------------------------------\n", options: {} },
        { text: "Alokasi BTC: ", options: { fontSize: 11 } }, { text: "1.00/2.50 = 40%\n", options: { bold: true, color: "27ae60" } },
        { text: "Alokasi ETH: ", options: { fontSize: 11 } }, { text: "0.88/2.50 = 35%\n", options: { bold: true, color: "27ae60" } },
        { text: "Alokasi BNB: ", options: { fontSize: 11 } }, { text: "0.62/2.50 = 25%", options: { bold: true, color: "27ae60" } }
    ], { x: 6.0, y: 1.7, w: 3.3, color: "333333", valign: "top" });

    // Footer Navigasi
    slideSim.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Slide 5: Network Markowitz ---
    let slideNet = pres.addSlide();
    if (fs.existsSync("bg_watermark.png")) {
        slideNet.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    }
    if (fs.existsSync("logo_unm.png")) {
        slideNet.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    }

    slideNet.addText("Implementasi Network Markowitz", { x: 0.5, y: 0.4, w: "90%", fontSize: 26, bold: true, color: "003366" });
    slideNet.addText("Penyesuaian Bobot Berdasarkan Skor Sentralitas (Risiko Sistemik):", { x: 0.5, y: 1.0, w: "90%", fontSize: 14, color: "333333" });

    // Tabel Perbandingan
    slideNet.addTable([
        [
            { text: "Aset", options: { bold: true, fill: "003366", color: "ffffff", align: "center" } },
            { text: "Centrality (C)", options: { bold: true, fill: "003366", color: "ffffff", align: "center" } },
            { text: "Classic Weight", options: { bold: true, fill: "003366", color: "ffffff", align: "center" } },
            { text: "Network Weight", options: { bold: true, fill: "27ae60", color: "ffffff", align: "center" } }
        ],
        ["Bitcoin (BTC)", "0.85 (High)", "40.0%", "35.6%"],
        ["Ethereum (ETH)", "0.60 (Med)", "35.0%", "36.3%"],
        ["Binance Coin (BNB)", "0.35 (Low)", "25.0%", "28.1%"]
    ], { x: 0.5, y: 1.6, w: 5.5, fontSize: 12, border: { pt: 1, color: "dddddd" }, align: "center", valign: "middle" });

    // Box Penjelasan Logika Network (Lebih Detail)
    slideNet.addShape(pres.ShapeType.rect, { x: 6.2, y: 1.5, w: 3.5, h: 3.5, fill: { color: "f4fcf4" }, line: { color: "27ae60", width: 1.5 } });
    slideNet.addText([
        { text: "Langkah Penjabaran (\u03b3 = 0.5):\n", options: { bold: true, color: "003366", fontSize: 13, breakLine: true } },
        { text: "Rumus: Skor\u2099\u2091\u209c = Skor\u2092\u2091\u2098 - (\u03b3 \u00d7 C)\n\n", options: { italic: true, fontSize: 10, color: "c0392b" } },
        { text: "1. BTC: ", options: { fontSize: 11 } }, { text: "1.00 - (0.5 \u00d7 0.85) = 0.57\n", options: { bold: true } },
        { text: "2. ETH: ", options: { fontSize: 11 } }, { text: "0.88 - (0.5 \u00d7 0.60) = 0.58\n", options: { bold: true } },
        { text: "3. BNB: ", options: { fontSize: 11 } }, { text: "0.62 - (0.5 \u00d7 0.35) = 0.45\n", options: { bold: true } },
        { text: "------------------------------------------\n", options: {} },
        { text: "Total Skor Net: ", options: { fontSize: 11 } }, { text: "0.57 + 0.58 + 0.45 = 1.60\n", options: { bold: true } },
        { text: "------------------------------------------\n", options: {} },
        { text: "Bobot BTC: ", options: { fontSize: 11 } }, { text: "0.57/1.60 \u2248 35.6%\n", options: { bold: true, color: "27ae60" } },
        { text: "Bobot ETH: ", options: { fontSize: 11 } }, { text: "0.58/1.60 \u2248 36.3%\n", options: { bold: true, color: "27ae60" } },
        { text: "Bobot BNB: ", options: { fontSize: 11 } }, { text: "0.45/1.60 \u2248 28.1%", options: { bold: true, color: "27ae60" } }
    ], { x: 6.35, y: 1.7, w: 3.2, color: "333333", valign: "top" });

    // --- Slide 8 sampai 16: Detail 9 Fitur Observasi SAC ---
    const features = [
        {
            num: 1,
            name: "MST.Dist",
            fullName: "Minimum Spanning Tree Distance",
            theme: "27ae60",
            lightTheme: "f4fcf4",
            leftTitle: "🌐 Formulasi & Definisi",
            leftTexts: [
                { text: "Jarak Minimum Spanning Tree (MST)\n", options: { bold: true, fontSize: 13, color: "003366" } },
                { text: "Mengukur total panjang (sum of edge weights) dari pohon rentang minimum yang dibentuk dari matriks jarak korelasi aset.\n\n", options: { fontSize: 10 } },
                { text: "Rumus Jarak Korelasi:\n", options: { bold: true, fontSize: 11, color: "003366" } },
                { text: "d_ij = sqrt(2 * (1 - ρ_ij)), di mana ρ_ij adalah koefisien korelasi Pearson antar log-return aset.\n\n", options: { fontSize: 10, italic: true } },
                { text: "⚠️ Penskalaan (Feature Scaling):\n", options: { bold: true, fontSize: 11, color: "c0392b" } },
                { text: "Fitur dikalikan dengan 0.1 (MST × 0.1) untuk menyeimbangkan rentang nilai agar Neural Network belajar secara stabil.", options: { fontSize: 10, italic: true } }
            ],
            rightTitle: "💡 Peran & Intuisi Agen SAC",
            rightTexts: [
                { text: "Deteksi Kerapatan & Risiko Sistemik:\n", options: { bold: true, fontSize: 12, color: "003366" } },
                { text: "Mengidentifikasi tingkat integrasi pasar kripto. Jarak MST yang semakin pendek menunjukkan korelasi antar aset sedang menguat secara masif (pasar menyatu).\n\n", options: { fontSize: 10 } },
                { text: "Respons Kontroler SAC:\n", options: { bold: true, fontSize: 11, color: "27ae60" } },
                { text: "Agen SAC belajar menaikkan parameter penalti sentralitas (γ) saat MST menyusut, guna melindungi portofolio dari penularan kegagalan sistemik.\n\n", options: { fontSize: 10, italic: true } },
                { text: "📚 Referensi Akademis:\n", options: { bold: true, fontSize: 9, color: "7f8c8d" } },
                { text: "• Mantegna (1999) - Hierarchical structure in financial markets.\n• Giudici (2020) - Network-based risk in crypto.", options: { fontSize: 8.5, italic: true } }
            ],
            summary: "Kesimpulan: Jarak MST yang mengecil menandakan pasar berada dalam kondisi rentan sistemik tinggi, memicu agen SAC untuk bertindak defensif dengan menaikkan parameter penalti sentralitas (γ)."
        },
        {
            num: 2,
            name: "Spectral.Gap",
            fullName: "Spectral Gap (Laplacian Eigenvalue)",
            theme: "1abc9c",
            lightTheme: "f0fbf9",
            leftTitle: "🌐 Formulasi & Definisi",
            leftTexts: [
                { text: "Spectral Gap (Algebraic Connectivity)\n", options: { bold: true, fontSize: 13, color: "003366" } },
                { text: "Selisih antara dua eigenvalue terkecil dari matriks Laplacian Laplacian ternormalisasi (L_norm) graf MST: λ_2 - λ_1. Karena λ_1 selalu bernilai 0, maka Spectral Gap sama dengan λ_2.\n\n", options: { fontSize: 10 } },
                { text: "Konektivitas Aljabar:\n", options: { bold: true, fontSize: 11, color: "003366" } },
                { text: "Mengukur tingkat kemudahan graf untuk terbagi menjadi subnet independen.\n\n", options: { fontSize: 10 } },
                { text: "⚠️ Penskalaan (Feature Scaling):\n", options: { bold: true, fontSize: 11, color: "27ae60" } },
                { text: "Menggunakan skala asli (x1) karena nilainya sudah berada dalam rentang terikat 0 s/d 1.", options: { fontSize: 10, italic: true } }
            ],
            rightTitle: "💡 Peran & Intuisi Agen SAC",
            rightTexts: [
                { text: "Ketahanan Topologi Jaringan:\n", options: { bold: true, fontSize: 12, color: "003366" } },
                { text: "Spectral Gap yang sangat kecil menunjukkan struktur jaringan yang sangat kompak dan homogen. Guncangan pada satu aset kripto akan merambat sangat cepat ke aset lainnya.\n\n", options: { fontSize: 10 } },
                { text: "Respons Kontroler SAC:\n", options: { bold: true, fontSize: 11, color: "1abc9c" } },
                { text: "Agen SAC mendeteksi kerentanan struktural ini. Ketika Spectral Gap mengecil, agen secara preventif meningkatkan penalti sentralitas untuk menyebar bobot keluar dari pusat jaringan.\n\n", options: { fontSize: 10, italic: true } },
                { text: "📚 Referensi Akademis:\n", options: { bold: true, fontSize: 9, color: "7f8c8d" } },
                { text: "• Giudici & Spelta (2016) - Graphical network models for systemic risk.", options: { fontSize: 8.5, italic: true } }
            ],
            summary: "Kesimpulan: Spectral Gap yang rendah memperingatkan agen tentang struktur korelasi jaringan yang sangat ringkih, memicu tindakan diversifikasi defensif dari SAC Controller."
        },
        {
            num: 3,
            name: "VolShort",
            fullName: "Short-Term Volatility (V5)",
            theme: "2980b9",
            lightTheme: "f4f9fc",
            leftTitle: "🌐 Formulasi & Definisi",
            leftTexts: [
                { text: "Volatility Jendela 5 Hari (VolShort)\n", options: { bold: true, fontSize: 13, color: "003366" } },
                { text: "Standar deviasi harian dari rata-rata tertimbang log-return koin dalam portfolio selama 5 hari bursa terakhir.\n\n", options: { fontSize: 10 } },
                { text: "Fungsi Pengukuran:\n", options: { bold: true, fontSize: 11, color: "003366" } },
                { text: "Menangkap lonjakan fluktuasi jangka pendek, mendeteksi ketidakpastian mendadak akibat rumor, berita, atau sentimen pasar sesaat.\n\n", options: { fontSize: 10 } },
                { text: "⚠️ Penskalaan (Feature Scaling):\n", options: { bold: true, fontSize: 11, color: "c0392b" } },
                { text: "Dikalikan 100 (VolShort × 100) karena standar deviasi harian sangat kecil (sekitar 0.01 - 0.03), diubah menjadi rentang 1.0 - 3.0 agar sensitif bagi model.", options: { fontSize: 10, italic: true } }
            ],
            rightTitle: "💡 Peran & Intuisi Agen SAC",
            rightTexts: [
                { text: "Deteksi Guncangan & Noise Pasar:\n", options: { bold: true, fontSize: 12, color: "003366" } },
                { text: "Volatilitas jangka pendek yang melonjak memberi tahu agen adanya kepanikan atau spekulasi yang sedang terjadi di pasar.\n\n", options: { fontSize: 10 } },
                { text: "Respons Kontroler SAC:\n", options: { bold: true, fontSize: 11, color: "2980b9" } },
                { text: "SAC belajar menurunkan eksposur risiko dengan menggeser portofolio ke aset dengan volatilitas rendah atau menyesuaikan γ secara ketat agar portofolio tidak over-exposed terhadap noise trading.\n\n", options: { fontSize: 10, italic: true } },
                { text: "📚 Referensi Akademis:\n", options: { bold: true, fontSize: 9, color: "7f8c8d" } },
                { text: "• Jiang et al. (2017) - A Deep Reinforcement Learning Framework for Portfolio Management.\n• Markowitz (1952) - Portfolio Selection.", options: { fontSize: 8.5, italic: true } }
            ],
            summary: "Kesimpulan: Volatilitas 5 hari menangkap guncangan harga tak terduga secara real-time, memberi sinyal kepada SAC untuk langsung membatasi alokasi pada aset berisiko tinggi."
        },
        {
            num: 4,
            name: "VolLong",
            fullName: "Long-Term Volatility (V20)",
            theme: "34495e",
            lightTheme: "f6f8fa",
            leftTitle: "🌐 Formulasi & Definisi",
            leftTexts: [
                { text: "Volatility Jendela 20 Hari (VolLong)\n", options: { bold: true, fontSize: 13, color: "003366" } },
                { text: "Standar deviasi harian dari rata-rata log-return koin dalam portfolio selama jendela observasi yang lebih panjang (20 hari bursa).\n\n", options: { fontSize: 10 } },
                { text: "Fungsi Pengukuran:\n", options: { bold: true, fontSize: 11, color: "003366" } },
                { text: "Mengukur risiko historis jangka menengah, memberikan baseline risiko yang lebih stabil dibanding fluktuasi harian.\n\n", options: { fontSize: 10 } },
                { text: "⚠️ Penskalaan (Feature Scaling):\n", options: { bold: true, fontSize: 11, color: "c0392b" } },
                { text: "Dikalikan 100 (VolLong × 100) demi menyamakan skala fitur volatilitas jangka pendek.", options: { fontSize: 10, italic: true } }
            ],
            rightTitle: "💡 Peran & Intuisi Agen SAC",
            rightTexts: [
                { text: "Identifikasi Rezim Risiko Pasar:\n", options: { bold: true, fontSize: 12, color: "003366" } },
                { text: "Berperan sebagai jangkar stabil bagi agen. Membantu membedakan apakah lonjakan volatilitas di VolShort adalah anomali singkat atau pergeseran ke rezim pasar berisiko tinggi (bearish market).\n\n", options: { fontSize: 10 } },
                { text: "Respons Kontroler SAC:\n", options: { bold: true, fontSize: 11, color: "34495e" } },
                { text: "Menggunakan VolLong untuk merumuskan kebijakan alokasi jangka panjang yang konsisten dan stabil (tidak terlalu reaktif terhadap kejutan harian).\n\n", options: { fontSize: 10, italic: true } },
                { text: "📚 Referensi Akademis:\n", options: { bold: true, fontSize: 9, color: "7f8c8d" } },
                { text: "• Jiang et al. (2017) - A Deep Reinforcement Learning Framework for Portfolio Management.", options: { fontSize: 8.5, italic: true } }
            ],
            summary: "Kesimpulan: Volatilitas 20 hari bertindak sebagai jangkar risiko jangka panjang, membantu agen SAC membedakan antara guncangan harga sesaat dengan tren risiko makro."
        },
        {
            num: 5,
            name: "Vol.Ratio",
            fullName: "Volatility Ratio (V5 / V20)",
            theme: "f39c12",
            lightTheme: "fefbf4",
            leftTitle: "🌐 Formulasi & Definisi",
            leftTexts: [
                { text: "Volatility Ratio (Rasio Volatilitas)\n", options: { bold: true, fontSize: 13, color: "003366" } },
                { text: "Rasio matematis antara volatilitas jangka pendek (5 hari) terhadap volatilitas jangka panjang (20 hari): Vol.Ratio = V5 / V20.\n\n", options: { fontSize: 10 } },
                { text: "Fungsi Pengukuran:\n", options: { bold: true, fontSize: 11, color: "003366" } },
                { text: "Mengukur rasio perubahan volatilitas. Jika rasio > 1, pasar sedang mengalami eskalasi risiko baru. Jika < 1, pasar sedang mendingin.\n\n", options: { fontSize: 10 } },
                { text: "⚠️ Penskalaan (Feature Scaling):\n", options: { bold: true, fontSize: 11, color: "27ae60" } },
                { text: "Skala asli (x1) karena merupakan pembagian tanpa dimensi yang secara alami bernilai di sekitar 0.5 s/d 2.5.", options: { fontSize: 10, italic: true } }
            ],
            rightTitle: "💡 Peran & Intuisi Agen SAC",
            rightTexts: [
                { text: "Deteksi Transisi Rezim Pasar:\n", options: { bold: true, fontSize: 12, color: "003366" } },
                { text: "Rasio > 1 mengisyaratkan ketidakstabilan mendadak (gejolak baru). Rasio < 1 menunjukkan pasar mulai mendingin dan kembali tenang.\n\n", options: { fontSize: 10 } },
                { text: "Respons Kontroler SAC:\n", options: { bold: true, fontSize: 11, color: "f39c12" } },
                { text: "Ketika rasio menunjukkan pasar mendingin (<1), agen SAC akan memicu pelonggaran penalti sentralitas (γ) untuk memberi ruang bagi alokasi agresif demi menangkap alpha (return maksimal).\n\n", options: { fontSize: 10, italic: true } },
                { text: "📚 Referensi Akademis:\n", options: { bold: true, fontSize: 9, color: "7f8c8d" } },
                { text: "• Jiang et al. (2017) - DRL in Portfolio Management.", options: { fontSize: 8.5, italic: true } }
            ],
            summary: "Kesimpulan: Volatility Ratio secara presisi mendeteksi momen transisi kecepatan pasar, memandu transisi agen SAC dari gaya investasi defensif ke agresif."
        },
        {
            num: 6,
            name: "Mom5d",
            fullName: "Short-Term Momentum (M5)",
            theme: "8e44ad",
            lightTheme: "faf5fc",
            leftTitle: "🌐 Formulasi & Definisi",
            leftTexts: [
                { text: "Momentum Jendela 5 Hari (Mom5d)\n", options: { bold: true, fontSize: 13, color: "003366" } },
                { text: "Akumulasi total log-return dari indeks rata-rata aset portofolio selama 5 hari bursa terakhir.\n\n", options: { fontSize: 10 } },
                { text: "Fungsi Pengukuran:\n", options: { bold: true, fontSize: 11, color: "003366" } },
                { text: "Mengukur tren harga jangka pendek guna mendeteksi kecenderungan kelanjutan tren (momentum) akibat perilaku pelaku pasar.\n\n", options: { fontSize: 10 } },
                { text: "⚠️ Penskalaan (Feature Scaling):\n", options: { bold: true, fontSize: 11, color: "c0392b" } },
                { text: "Dikalikan 100 (Mom5d × 100) karena log-return harian bernilai sangat kecil (misal 0.003), diubah menjadi skala 0.3 agar gradien neural network sensitif.", options: { fontSize: 10, italic: true } }
            ],
            rightTitle: "💡 Peran & Intuisi Agen SAC",
            rightTexts: [
                { text: "Eksploitasi Tren Jangka Pendek:\n", options: { bold: true, fontSize: 12, color: "003366" } },
                { text: "Menangkap arah sentimen jangka pendek. Tren positif kuat menandakan pasar sedang dalam fase bullish jangka pendek.\n\n", options: { fontSize: 10 } },
                { text: "Respons Kontroler SAC:\n", options: { bold: true, fontSize: 11, color: "8e44ad" } },
                { text: "Jika momentum jangka pendek sangat positif, agen SAC mengadopsi taktik agresif (trend-following) untuk memperbesar return. Sebaliknya, momentum negatif memicu agen membatasi alokasi pada koin rentan.\n\n", options: { fontSize: 10, italic: true } },
                { text: "📚 Referensi Akademis:\n", options: { bold: true, fontSize: 9, color: "7f8c8d" } },
                { text: "• Jegadeesh & Titman (1993) - Returns to Buying Winners and Selling Losers.", options: { fontSize: 8.5, italic: true } }
            ],
            summary: "Kesimpulan: Momentum 5 hari membantu agen SAC mendeteksi dorongan tren harga jangka pendek guna menangkap profit maksimal di pasar yang sedang menguat."
        },
        {
            num: 7,
            name: "Mom20d",
            fullName: "Long-Term Momentum (M20)",
            theme: "9b59b6",
            lightTheme: "faf6fd",
            leftTitle: "🌐 Formulasi & Definisi",
            leftTexts: [
                { text: "Momentum Jendela 20 Hari (Mom20d)\n", options: { bold: true, fontSize: 13, color: "003366" } },
                { text: "Akumulasi total log-return dari indeks rata-rata aset portofolio selama 20 hari bursa terakhir.\n\n", options: { fontSize: 10 } },
                { text: "Fungsi Pengukuran:\n", options: { bold: true, fontSize: 11, color: "003366" } },
                { text: "Mengukur kekuatan tren harga jangka menengah/panjang, menyaring gangguan noise harga dari pergerakan harian.\n\n", options: { fontSize: 10 } },
                { text: "⚠️ Penskalaan (Feature Scaling):\n", options: { bold: true, fontSize: 11, color: "c0392b" } },
                { text: "Dikalikan 100 (Mom20d × 100) agar setara dengan skala fitur momentum jangka pendek.", options: { fontSize: 10, italic: true } }
            ],
            rightTitle: "💡 Peran & Intuisi Agen SAC",
            rightTexts: [
                { text: "Konfirmasi Tren Makro:\n", options: { bold: true, fontSize: 12, color: "003366" } },
                { text: "Menjadi indikator utama apakah pasar kripto sedang berada di fase bullish atau bearish jangka panjang secara struktural.\n\n", options: { fontSize: 10 } },
                { text: "Respons Kontroler SAC:\n", options: { bold: true, fontSize: 11, color: "9b59b6" } },
                { text: "Agen SAC menyandingkan fitur network dengan Mom20d untuk memastikan keputusan alokasi didasari oleh tren pasar yang solid dan matang, bukan manipulasi tren sesaat.\n\n", options: { fontSize: 10, italic: true } },
                { text: "📚 Referensi Akademis:\n", options: { bold: true, fontSize: 9, color: "7f8c8d" } },
                { text: "• Jegadeesh & Titman (1993) - Returns to Buying Winners and Selling Losers.", options: { fontSize: 8.5, italic: true } }
            ],
            summary: "Kesimpulan: Momentum 20 hari memberikan indikasi arah tren utama pasar, memastikan bahwa keputusan alokasi agen SAC selaras dengan rezim makro pasar."
        },
        {
            num: 8,
            name: "MomCross",
            fullName: "Momentum Crossover (M5 - M20)",
            theme: "d35400",
            lightTheme: "fdf7f4",
            leftTitle: "🌐 Formulasi & Definisi",
            leftTexts: [
                { text: "Momentum Crossover (MomCross)\n", options: { bold: true, fontSize: 13, color: "003366" } },
                { text: "Selisih matematis antara momentum jangka pendek (5 hari) dan momentum jangka panjang (20 hari): MomCross = M5 - M20.\n\n", options: { fontSize: 10 } },
                { text: "Fungsi Pengukuran:\n", options: { bold: true, fontSize: 11, color: "003366" } },
                { text: "Mendeteksi akselerasi atau deselerasi tren harga. Berguna untuk memprediksi perubahan arah tren.\n\n", options: { fontSize: 10 } },
                { text: "⚠️ Penskalaan (Feature Scaling):\n", options: { bold: true, fontSize: 11, color: "c0392b" } },
                { text: "Dikalikan 100 (MomCross × 100) agar skalanya selaras dengan fitur momentum pembentuknya.", options: { fontSize: 10, italic: true } }
            ],
            rightTitle: "💡 Peran & Intuisi Agen SAC",
            rightTexts: [
                { text: "Deteksi Titik Reversal Arah Pasar:\n", options: { bold: true, fontSize: 12, color: "003366" } },
                { text: "MomCross > 0 mengindikasikan Golden Cross (akselerasi naik). MomCross < 0 mengindikasikan Death Cross (perlambatan/tren turun).\n\n", options: { fontSize: 10 } },
                { text: "Respons Kontroler SAC:\n", options: { bold: true, fontSize: 11, color: "d35400" } },
                { text: "Saat terjadi Death Cross (< 0), agen SAC secara dinamis meningkatkan parameter penalti sentralitas (γ) untuk memicu diversifikasi defensif sebelum pasar jatuh lebih dalam.\n\n", options: { fontSize: 10, italic: true } },
                { text: "📚 Referensi Akademis:\n", options: { bold: true, fontSize: 9, color: "7f8c8d" } },
                { text: "• Jegadeesh & Titman (1993) - Returns to Buying Winners and Selling Losers.", options: { fontSize: 8.5, italic: true } }
            ],
            summary: "Kesimpulan: Momentum Crossover secara dinamis menangkap titik jenuh tren harga, memicu rebalancing defensif yang cepat sebelum terjadi kejatuhan pasar."
        },
        {
            num: 9,
            name: "Pct.Uptrend",
            fullName: "Percentage of Uptrend Coins (PU)",
            theme: "c0392b",
            lightTheme: "fdf5f5",
            leftTitle: "🌐 Formulasi & Definisi",
            leftTexts: [
                { text: "Persentase Aset Menguat (Pct.Uptrend)\n", options: { bold: true, fontSize: 13, color: "003366" } },
                { text: "Persentase jumlah koin dalam universe investasi yang mencatatkan return harian positif pada hari observasi.\n\n", options: { fontSize: 10 } },
                { text: "Fungsi Pengukuran:\n", options: { bold: true, fontSize: 11, color: "003366" } },
                { text: "Mengukur tingkat kerataan kenaikan di pasar (Market Breadth Indicator) untuk mengetahui seberapa meluas penguatan pasar.\n\n", options: { fontSize: 10 } },
                { text: "⚠️ Penskalaan (Feature Scaling):\n", options: { bold: true, fontSize: 11, color: "27ae60" } },
                { text: "Skala asli (x1) karena nilainya sudah pasti berada dalam rentang terikat 0 s/d 1.", options: { fontSize: 10, italic: true } }
            ],
            rightTitle: "💡 Peran & Intuisi Agen SAC",
            rightTexts: [
                { text: "Deteksi Kesehatan Pasar Menyeluruh:\n", options: { bold: true, fontSize: 12, color: "003366" } },
                { text: "Nilai tinggi (>80%) menunjukkan kenaikan pasar yang sehat dan didorong merata. Nilai rendah (<20%) menunjukkan kepanikan pasar di mana hampir seluruh koin rontok.\n\n", options: { fontSize: 10 } },
                { text: "Respons Kontroler SAC:\n", options: { bold: true, fontSize: 11, color: "c0392b" } },
                { text: "Pada pasar bullish merata, agen melonggarkan γ demi alokasi terpusat pada koin alpha. Pada kejatuhan massal (<20%), agen memperketat γ secara maksimal untuk memaksa diversifikasi ekstrim.\n\n", options: { fontSize: 10, italic: true } },
                { text: "📚 Referensi Akademis:\n", options: { bold: true, fontSize: 9, color: "7f8c8d" } },
                { text: "• Jiang et al. (2017) - DRL in Portfolio Management.", options: { fontSize: 8.5, italic: true } }
            ],
            summary: "Kesimpulan: Percentage Uptrend mendeteksi tingkat optimisme kolektif pasar kripto, memandu SAC Controller menentukan derajat diversifikasi optimal portofolio."
        }
    ];

    features.forEach(feat => {
        let slideFeat = pres.addSlide();
        if (fs.existsSync("bg_watermark.png")) {
            slideFeat.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
        }
        if (fs.existsSync("logo_unm.png")) {
            slideFeat.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
        }

        // Title
        slideFeat.addText(`Fitur Observasi SAC (${feat.num}/9): ${feat.name}`, { x: 0.5, y: 0.4, w: "90%", fontSize: 26, bold: true, color: "003366" });
        slideFeat.addShape(pres.ShapeType.line, { x: 0.5, y: 0.9, w: 1.5, h: 0, line: { color: "e67e22", width: 3 } });
        slideFeat.addText(`Analisis mendalam fitur ${feat.fullName} sebagai input state SAC Agent:`, { x: 0.5, y: 1.05, w: "90%", fontSize: 13, color: "333333" });

        // Left Card
        slideFeat.addShape(pres.ShapeType.rect, { x: 0.5, y: 1.4, w: 4.35, h: 3.2, fill: { color: "ffffff" }, line: { color: "003366", width: 1.5 }, rectRadius: 0.05 });
        slideFeat.addShape(pres.ShapeType.rect, { x: 0.5, y: 1.4, w: 4.35, h: 0.45, fill: { color: "003366" } });
        slideFeat.addText(feat.leftTitle, { x: 0.5, y: 1.4, w: 4.35, h: 0.45, fontSize: 14, bold: true, color: "ffffff", align: "center", valign: "middle" });
        slideFeat.addText(feat.leftTexts, { x: 0.7, y: 2.0, w: 3.95, h: 2.5, color: "333333", valign: "top", lineSpacing: 10 });

        // Right Card
        slideFeat.addShape(pres.ShapeType.rect, { x: 5.15, y: 1.4, w: 4.35, h: 3.2, fill: { color: "ffffff" }, line: { color: feat.theme, width: 1.5 }, rectRadius: 0.05 });
        slideFeat.addShape(pres.ShapeType.rect, { x: 5.15, y: 1.4, w: 4.35, h: 0.45, fill: { color: feat.theme } });
        slideFeat.addText(feat.rightTitle, { x: 5.15, y: 1.4, w: 4.35, h: 0.45, fontSize: 14, bold: true, color: "ffffff", align: "center", valign: "middle" });
        slideFeat.addText(feat.rightTexts, { x: 5.35, y: 2.0, w: 3.95, h: 2.5, color: "333333", valign: "top", lineSpacing: 10 });

        // Summary Box
        slideFeat.addShape(pres.ShapeType.rect, { x: 0.5, y: 4.75, w: 9.0, h: 0.45, fill: { color: feat.lightTheme }, line: { color: feat.theme, width: 1.0 }, rectRadius: 0.05 });
        slideFeat.addText(feat.summary, { x: 0.5, y: 4.75, w: 9.0, h: 0.45, fontSize: 10, bold: true, color: "003366", align: "center", valign: "middle" });

        // Footer Navigasi
        slideFeat.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });
    });

    // --- Slide 18: Simulasi Penghitungan Fitur Observasi SAC (Toy Dataset) ---
    let slideSimFitur = pres.addSlide();
    if (fs.existsSync("bg_watermark.png")) {
        slideSimFitur.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    }
    if (fs.existsSync("logo_unm.png")) {
        slideSimFitur.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    }

    // Header
    slideSimFitur.addText("Simulasi Penghitungan Fitur Observasi SAC", { x: 0.5, y: 0.3, w: "85%", fontSize: 26, bold: true, color: "003366" });
    slideSimFitur.addShape(pres.ShapeType.line, { x: 0.5, y: 0.72, w: 1.5, h: 0, line: { color: "e67e22", width: 3 } });
    slideSimFitur.addText("Contoh konkret perhitungan matematis beberapa fitur state SAC dengan Toy Dataset (3 koin, 3 hari):", { x: 0.5, y: 0.82, w: "85%", fontSize: 12, color: "555555" });

    // === CARD KIRI: Toy Dataset & Vol/Mom Fitur ===
    slideSimFitur.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.15, w: 4.5, h: 3.55, fill: { color: "ffffff" }, line: { color: "003366", width: 1.5 }, rectRadius: 0.05 });
    slideSimFitur.addShape(pres.ShapeType.rect, { x: 0.4, y: 1.15, w: 4.5, h: 0.42, fill: { color: "003366" } });
    slideSimFitur.addText("📊 Toy Dataset & Vol/Mom Fitur", { x: 0.4, y: 1.15, w: 4.5, h: 0.42, fontSize: 13, bold: true, color: "ffffff", align: "center", valign: "middle" });

    // Tabel Toy Return
    slideSimFitur.addTable([
        [
            { text: "Hari", options: { bold: true, fill: "003366", color: "ffffff", align: "center", fontSize: 8.5 } },
            { text: "BTC Return", options: { bold: true, fill: "003366", color: "ffffff", align: "center", fontSize: 8.5 } },
            { text: "ETH Return", options: { bold: true, fill: "003366", color: "ffffff", align: "center", fontSize: 8.5 } },
            { text: "BNB Return", options: { bold: true, fill: "003366", color: "ffffff", align: "center", fontSize: 8.5 } },
            { text: "Portofolio", options: { bold: true, fill: "e67e22", color: "ffffff", align: "center", fontSize: 8.5 } }
        ],
        [
            { text: "Day 1", options: { align: "center", fontSize: 8 } },
            { text: "+1.0%", options: { align: "center", fontSize: 8 } },
            { text: "+2.0%", options: { align: "center", fontSize: 8 } },
            { text: "-1.0%", options: { align: "center", fontSize: 8 } },
            { text: "+0.7%", options: { align: "center", fontSize: 8, bold: true } }
        ],
        [
            { text: "Day 2", options: { align: "center", fontSize: 8 } },
            { text: "+2.0%", options: { align: "center", fontSize: 8 } },
            { text: "-1.0%", options: { align: "center", fontSize: 8 } },
            { text: "+2.0%", options: { align: "center", fontSize: 8 } },
            { text: "+1.1%", options: { align: "center", fontSize: 8, bold: true } }
        ],
        [
            { text: "Day 3", options: { align: "center", fontSize: 8 } },
            { text: "-1.0%", options: { align: "center", fontSize: 8 } },
            { text: "+1.0%", options: { align: "center", fontSize: 8 } },
            { text: "+1.0%", options: { align: "center", fontSize: 8 } },
            { text: "+0.2%", options: { align: "center", fontSize: 8, bold: true } }
        ],
    ], { x: 0.5, y: 1.62, w: 4.3, fontSize: 8.5, border: { pt: 1, color: "cccccc" }, align: "center", valign: "middle", rowH: 0.27 });

    // Penjelasan Vol/Mom calculations
    slideSimFitur.addText([
        { text: "Kalkulasi Fitur Momentum & Volatilitas:", options: { bold: true, fontSize: 10, color: "003366", breakLine: true } },
        { text: "\u2022 Mom5d (M3 - Toy): \u03a3 Ret(D1:D3) = 0.7%+1.1%+0.2% = ", options: { bold: true, fontSize: 8.5, color: "2c3e50" } },
        { text: "+2.0%", options: { bold: true, fontSize: 8.5, color: "27ae60" } },
        { text: " | Scaled (x100): ", options: { fontSize: 8.5 } },
        { text: "2.0", options: { bold: true, fontSize: 8.5, color: "27ae60", breakLine: true } },
        { text: "", options: { breakLine: true, fontSize: 4 } },

        { text: "\u2022 MomCross: Ret(D3) - Mom(M3) = 0.2% - 2.0% = ", options: { bold: true, fontSize: 8.5, color: "2c3e50" } },
        { text: "-1.8%", options: { bold: true, fontSize: 8.5, color: "c0392b" } },
        { text: " | Scaled (x100): ", options: { fontSize: 8.5 } },
        { text: "-1.8", options: { bold: true, fontSize: 8.5, color: "c0392b", breakLine: true } },
        { text: "", options: { breakLine: true, fontSize: 4 } },

        { text: "\u2022 Vol.Ratio: Vol_Short / Vol_Long = 0.25% / 0.57% = ", options: { bold: true, fontSize: 8.5, color: "2c3e50" } },
        { text: "0.44", options: { bold: true, fontSize: 8.5, color: "f39c12" } },
        { text: " | Scaled (x1): ", options: { fontSize: 8.5 } },
        { text: "0.44", options: { bold: true, fontSize: 8.5, color: "f39c12", breakLine: true } }
    ], { x: 0.5, y: 2.85, w: 4.3, h: 1.8, color: "333333", valign: "top", lineSpacing: 16 });

    // === CARD KANAN: Pct.Uptrend & MST Jarak ===
    slideSimFitur.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.15, w: 4.5, h: 3.55, fill: { color: "ffffff" }, line: { color: "27ae60", width: 1.5 }, rectRadius: 0.05 });
    slideSimFitur.addShape(pres.ShapeType.rect, { x: 5.1, y: 1.15, w: 4.5, h: 0.42, fill: { color: "27ae60" } });
    slideSimFitur.addText("🌐 Pct.Uptrend & MST Jarak", { x: 5.1, y: 1.15, w: 4.5, h: 0.42, fontSize: 13, bold: true, color: "ffffff", align: "center", valign: "middle" });

    slideSimFitur.addText([
        { text: "1. Persentase Koin Menguat (Pct.Uptrend / PU):", options: { bold: true, fontSize: 10, color: "003366", breakLine: true } },
        { text: "\u2022 Day 3 Return: ", options: { fontSize: 8.5 } },
        { text: "BTC (-1.0% ❌), ETH (+1.0% ✅), BNB (+1.0% ✅)", options: { fontSize: 8.5, italic: true, breakLine: true } },
        { text: "\u2022 PU = 2 koin naik / 3 koin = ", options: { fontSize: 8.5 } },
        { text: "66.67% (0.67)", options: { bold: true, color: "27ae60", fontSize: 9 } },
        { text: " | Scaled (x1): ", options: { fontSize: 8.5 } },
        { text: "0.67", options: { bold: true, color: "27ae60", fontSize: 9, breakLine: true } },
        { text: "", options: { breakLine: true, fontSize: 6 } }, // Spacing

        { text: "2. Jarak Korelasi & MST (MST.Dist):", options: { bold: true, fontSize: 10, color: "003366", breakLine: true } },
        { text: "\u2022 Korelasi Pearson: ", options: { fontSize: 8.5 } },
        { text: "\u03c1_AB = 0.5, \u03c1_BC = -0.2, \u03c1_AC = 0.1", options: { fontSize: 8.5, italic: true, breakLine: true } },
        { text: "\u2022 Jarak d = \u221a(2\u00d7(1-\u03c1)): ", options: { fontSize: 8.5 } },
        { text: "d_AB = 1.00, d_BC = 1.55, d_AC = 1.34", options: { fontSize: 8.5, italic: true, breakLine: true } },
        { text: "\u2022 MST Jarak (\u03a3 edge A-B + A-C) = 1.00+1.34 = ", options: { fontSize: 8.5 } },
        { text: "2.34", options: { bold: true, color: "c0392b", fontSize: 9 } },
        { text: " | Scaled (x0.1): ", options: { fontSize: 8.5 } },
        { text: "0.234", options: { bold: true, color: "c0392b", fontSize: 9, breakLine: true } }
    ], { x: 5.2, y: 1.65, w: 4.3, h: 3.0, color: "333333", valign: "top", lineSpacing: 16 });

    // Summary Box
    slideSimFitur.addShape(pres.ShapeType.rect, { x: 0.4, y: 4.82, w: 9.2, h: 0.42, fill: { color: "eaf4fb" }, line: { color: "2980b9", width: 1.0 }, rectRadius: 0.05 });
    slideSimFitur.addText("Aplikasi State: Fitur-fitur ini (MST.Dist = 0.234, PU = 0.67, MomCross = -1.8, dll.) digabung menjadi Vektor State s_t berdimensi 9. Vektor ini di-umpankan ke Aktor & Kritik SAC untuk memprediksi nilai penalti sentralitas (\u03b3) optimal.", {
        x: 0.4, y: 4.82, w: 9.2, h: 0.42, fontSize: 9.5, bold: true, color: "003366", align: "center", valign: "middle"
    });

    slideSimFitur.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Slide 10: Ilustrasi Feature Scaling ---
    let slideScale = pres.addSlide();
    if (fs.existsSync("bg_watermark.png")) {
        slideScale.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    }
    if (fs.existsSync("logo_unm.png")) {
        slideScale.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    }

    slideScale.addText("Ilustrasi Pentingnya Feature Scaling (x100)", { x: 0.5, y: 0.4, w: "90%", fontSize: 28, bold: true, color: "003366" });
    slideScale.addShape(pres.ShapeType.line, { x: 0.5, y: 0.9, w: 1.5, h: 0, line: { color: "e67e22", width: 3 } });
    slideScale.addText("Simulasi membandingkan rambatan sinyal (forward pass) pada layer Neural Network:", { x: 0.5, y: 1.05, w: "90%", fontSize: 13, color: "333333" });

    // --- CARD KIRI: TANPA SCALING ---
    slideScale.addShape(pres.ShapeType.rect, { x: 0.5, y: 1.5, w: 4.3, h: 2.8, fill: { color: "fffafa" }, line: { color: "c0392b", width: 1.5 }, rectRadius: 0.05 });
    slideScale.addShape(pres.ShapeType.rect, { x: 0.5, y: 1.5, w: 4.3, h: 0.45, fill: { color: "c0392b" } });
    slideScale.addText("❌ Tanpa Scaling (Raw Data)", { x: 0.5, y: 1.5, w: 4.3, h: 0.45, fontSize: 14, bold: true, color: "ffffff", align: "center", valign: "middle" });
    
    slideScale.addText([
        { text: "1. Input Asli (Contoh Return 0.5%): ", options: { bold: true, fontSize: 11, color: "003366" } }, { text: "X = 0.005\n", options: { fontSize: 11 } },
        { text: "2. Bobot Awal Neural Network: ", options: { bold: true, fontSize: 11, color: "003366" } }, { text: "W = 0.1\n", options: { fontSize: 11 } },
        { text: "3. Aktivasi Layer 1 (X \u00d7 W): ", options: { bold: true, fontSize: 11, color: "003366" } }, { text: "0.0005\n", options: { fontSize: 11 } },
        { text: "4. Aktivasi Layer 2 (L1 \u00d7 W): ", options: { bold: true, fontSize: 11, color: "003366" } }, { text: "0.00005\n\n", options: { fontSize: 11 } },
        { text: "💥 Dampak (Vanishing Gradient):\n", options: { bold: true, fontSize: 12, color: "c0392b", breakLine: true } },
        { text: "Angka menjadi sangat kecil (mendekati nol). Saat proses penyesuaian (Backpropagation), nilai gradien error nyaris hilang (\u22480). Agen SAC kesulitan membedakan sinyal penting, sehingga proses belajar melambat atau berhenti total.", options: { fontSize: 10, italic: true } }
    ], { x: 0.7, y: 2.05, w: 3.9, h: 2.1, color: "333333", valign: "top", lineSpacing: 10 });

    // --- CARD KANAN: DENGAN SCALING ---
    slideScale.addShape(pres.ShapeType.rect, { x: 5.0, y: 1.5, w: 4.3, h: 2.8, fill: { color: "f4fcf4" }, line: { color: "27ae60", width: 1.5 }, rectRadius: 0.05 });
    slideScale.addShape(pres.ShapeType.rect, { x: 5.0, y: 1.5, w: 4.3, h: 0.45, fill: { color: "27ae60" } });
    slideScale.addText("✅ Dengan Scaling (x100)", { x: 5.0, y: 1.5, w: 4.3, h: 0.45, fontSize: 14, bold: true, color: "ffffff", align: "center", valign: "middle" });
    
    slideScale.addText([
        { text: "1. Input Skala (Return 0.5% \u00d7 100): ", options: { bold: true, fontSize: 11, color: "003366" } }, { text: "X = 0.5\n", options: { fontSize: 11 } },
        { text: "2. Bobot Awal Neural Network: ", options: { bold: true, fontSize: 11, color: "003366" } }, { text: "W = 0.1\n", options: { fontSize: 11 } },
        { text: "3. Aktivasi Layer 1 (X \u00d7 W): ", options: { bold: true, fontSize: 11, color: "003366" } }, { text: "0.05\n", options: { fontSize: 11 } },
        { text: "4. Aktivasi Layer 2 (L1 \u00d7 W): ", options: { bold: true, fontSize: 11, color: "003366" } }, { text: "0.005\n\n", options: { fontSize: 11 } },
        { text: "🚀 Dampak (Pembelajaran Stabil):\n", options: { bold: true, fontSize: 12, color: "27ae60", breakLine: true } },
        { text: "Sinyal bertahan pada rentang skala yang proporsional. Fungsi aktivasi (ReLU/Tanh) merespons dengan baik, dan gradien tetap terjaga signifikansinya. Agen SAC dapat dengan cepat mengidentifikasi pola pasar untuk aksi optimal.", options: { fontSize: 10, italic: true } }
    ], { x: 5.2, y: 2.05, w: 3.9, h: 2.1, color: "333333", valign: "top", lineSpacing: 10 });

    // Summary Box Bottom
    slideScale.addShape(pres.ShapeType.rect, { x: 0.5, y: 4.5, w: 8.8, h: 0.6, fill: { color: "ebf5fb" }, line: { color: "2980b9", width: 1.0 }, rectRadius: 0.05 });
    slideScale.addText("Kesimpulan: Scaling x100 menerjemahkan fraksi desimal super kecil menjadi nilai persentase yang terbaca jelas, mencegah fenomena gradien hilang (Vanishing Gradient) pada arsitektur Deep RL.", { x: 0.5, y: 4.5, w: 8.8, h: 0.6, fontSize: 11, bold: true, color: "003366", align: "center", valign: "middle" });

    slideScale.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Slide 11: Simulasi Perhitungan Sharpe Ratio ---
    let slideSharpe = pres.addSlide();
    if (fs.existsSync("bg_watermark.png")) {
        slideSharpe.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    }
    if (fs.existsSync("logo_unm.png")) {
        slideSharpe.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    }

    slideSharpe.addText("Simulasi Perhitungan Sharpe Ratio", { x: 0.5, y: 0.4, w: "90%", fontSize: 26, bold: true, color: "003366" });
    slideSharpe.addText("Contoh Data Historis (3 Aset, 4 Hari) - Bobot: 40% BTC, 30% ETH, 30% BNB", { x: 0.5, y: 1.0, w: "90%", fontSize: 13, color: "333333" });

    // Tabel Return Harian
    slideSharpe.addTable([
        [
            { text: "Hari", options: { bold: true, fill: "003366", color: "ffffff", align: "center" } },
            { text: "BTC (40%)", options: { bold: true, fill: "003366", color: "ffffff", align: "center" } },
            { text: "ETH (30%)", options: { bold: true, fill: "003366", color: "ffffff", align: "center" } },
            { text: "BNB (30%)", options: { bold: true, fill: "003366", color: "ffffff", align: "center" } },
            { text: "Portofolio", options: { bold: true, fill: "e67e22", color: "ffffff", align: "center" } }
        ],
        ["Day 1", "1.0%", "2.0%", "-1.0%", "0.7%"],
        ["Day 2", "2.0%", "-1.0%", "2.0%", "1.1%"],
        ["Day 3", "-1.0%", "1.0%", "1.0%", "0.2%"],
        ["Day 4", "0.0%", "2.0%", "0.0%", "0.6%"]
    ], { x: 0.5, y: 1.5, w: 5.5, fontSize: 11, border: { pt: 1, color: "dddddd" }, align: "center", valign: "middle" });

    // Box Penjabaran Kalkulasi (Detail Risk)
    slideSharpe.addShape(pres.ShapeType.rect, { x: 6.2, y: 1.0, w: 3.5, h: 4.4, fill: { color: "fff8ef" }, line: { color: "e67e22", width: 1.5 } });
    slideSharpe.addText([
        { text: "Langkah Kalkulasi:\n", options: { bold: true, color: "d35400", fontSize: 12, breakLine: true } },
        { text: "1. Mean Return (Rp): ", options: { bold: true, fontSize: 9 } }, { text: "0.65%\n", options: { fontSize: 9 } },
        
        { text: "2. Kalkulasi Risk (\u03c3p):\n", options: { bold: true, fontSize: 9 } },
        { text: "   \u2022 Dev\u00b2 Day 1: (0.7-0.65)\u00b2 = 0.0025\n", options: { fontSize: 8 } },
        { text: "   \u2022 Dev\u00b2 Day 2: (1.1-0.65)\u00b2 = 0.2025\n", options: { fontSize: 8 } },
        { text: "   \u2022 Dev\u00b2 Day 3: (0.2-0.65)\u00b2 = 0.2025\n", options: { fontSize: 8 } },
        { text: "   \u2022 Dev\u00b2 Day 4: (0.6-0.65)\u00b2 = 0.0025\n", options: { fontSize: 8 } },
        { text: "   \u2022 Sum Dev\u00b2 = 0.41 | Var = 0.41 / (4-1) = 0.137\n", options: { fontSize: 8 } },
        { text: "   \u2022 \u03c3p = \u221a0.137 = ", options: { fontSize: 8 } }, { text: "0.37%\n\n", options: { bold: true, fontSize: 8 } },
        
        { text: "3. Sharpe Ratio (Rf = 5% Ann):\n", options: { bold: true, fontSize: 9 } },
        { text: "   (0.65 - 0.014) / 0.37 = ", options: { fontSize: 9 } }, { text: "1.72\n", options: { bold: true, color: "27ae60", fontSize: 9 } },
        
        { text: "------------------------------------------\n", options: {} },
        { text: "Keterangan Istilah & Singkatan:\n", options: { bold: true, fontSize: 8.5, color: "d35400" } },
        { text: "\u2022 Rp: ", options: { bold: true, fontSize: 7.5 } }, { text: "Return Portofolio | ", options: { fontSize: 7.5 } },
        { text: "\u03c3p: ", options: { bold: true, fontSize: 7.5 } }, { text: "Risiko Portofolio\n", options: { fontSize: 7.5 } },
        { text: "\u2022 Rf: ", options: { bold: true, fontSize: 7.5 } }, { text: "Risk-Free Rate (Suku Bunga Bebas Risiko)\n", options: { fontSize: 7.5 } },
        { text: "\u2022 Ann: ", options: { bold: true, fontSize: 7.5 } }, { text: "Annualized (Tahunan). Rf harian = 5%/365 \u2248 0.014%\n", options: { fontSize: 7.5 } },
        { text: "\u2022 Dev & Var: ", options: { bold: true, fontSize: 7.5 } }, { text: "Deviasi (Selisih) & Varians (Rata-rata Dev\u00b2)\n", options: { fontSize: 7.5 } },
        
        { text: "------------------------------------------\n", options: {} },
        { text: "Apa yang digambarkan Sharpe?\n", options: { bold: true, fontSize: 9, color: "c0392b" } },
        { text: "Mengukur efisiensi: Keuntungan ekstra per unit risiko. ", options: { fontSize: 7.5 } },
        { text: "Semakin besar semakin bagus ", options: { fontSize: 7.5, bold: true, color: "27ae60" } },
        { text: "(return > risiko).", options: { fontSize: 7.5 } }
    ], { x: 6.4, y: 1.0, w: 3.2, color: "333333", valign: "top" });

    // Box Simulasi Rebalancing (Visual Alur)
    slideSharpe.addShape(pres.ShapeType.rect, { x: 0.5, y: 3.5, w: 5.5, h: 1.5, fill: { color: "f4f6f7" }, line: { color: "003366", width: 1.0 } });
    slideSharpe.addText([
        { text: "Simulasi Mekanisme Rebalancing (Day 1 \u2192 Day 2):\n", options: { bold: true, color: "003366", fontSize: 11, breakLine: true } },
        { text: "\u2022 Akhir Day 1 (Drift): ", options: { bold: true, fontSize: 10 } },
        { text: "Harga BTC & ETH naik, BNB turun. Bobot bergeser menjadi BTC 40.1%, ETH 30.4%, BNB 29.5%.\n", options: { fontSize: 9 } },
        { text: "\u2022 Aksi Rebalancing: ", options: { bold: true, fontSize: 10, color: "c0392b" } },
        { text: "Jual BTC & ETH, Beli BNB (Kembali ke 40/30/30).\n", options: { fontSize: 9 } },
        { text: "\u2022 Awal Day 2: ", options: { bold: true, fontSize: 10, color: "27ae60" } },
        { text: "Bobot sudah konsisten (40/30/30) sebelum menghitung return 1.1% di hari kedua.", options: { fontSize: 9 } }
    ], { x: 0.7, y: 3.6, w: 5.2, color: "333333", valign: "top" });

    slideSharpe.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Slide 7: Simulasi Sortino Ratio ---
    let slideSortino = pres.addSlide();
    if (fs.existsSync("bg_watermark.png")) {
        slideSortino.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    }
    if (fs.existsSync("logo_unm.png")) {
        slideSortino.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    }

    slideSortino.addText("Simulasi Performa: Sortino Ratio", { x: 0.5, y: 0.4, w: "90%", fontSize: 26, bold: true, color: "003366" });
    slideSortino.addText("Fokus pada Downside Risk (Hanya Menghitung Kerugian):", { x: 0.5, y: 1.0, w: "90%", fontSize: 14, color: "333333" });

    // Tabel Return (dengan Hari Negatif)
    slideSortino.addTable([
        [
            { text: "Hari", options: { bold: true, fill: "003366", color: "ffffff", align: "center" } },
            { text: "Return Portofolio", options: { bold: true, fill: "003366", color: "ffffff", align: "center" } },
            { text: "Status", options: { bold: true, fill: "003366", color: "ffffff", align: "center" } }
        ],
        ["Day 1", "2.0%", "Profit (Diabaikan di Downside)"],
        ["Day 2", "-1.5%", { text: "Loss (Dihitung)", options: { color: "c0392b", bold: true } }],
        ["Day 3", "0.5%", "Profit (Diabaikan di Downside)"],
        ["Day 4", "-0.5%", { text: "Loss (Dihitung)", options: { color: "c0392b", bold: true } }]
    ], { x: 0.5, y: 1.6, w: 5.5, fontSize: 11, border: { pt: 1, color: "dddddd" }, align: "center", valign: "middle" });

    // Box Penjelasan Sortino
    slideSortino.addShape(pres.ShapeType.rect, { x: 6.2, y: 1.0, w: 3.5, h: 4.4, fill: { color: "f4f0f7" }, line: { color: "8e44ad", width: 1.5 } });
    slideSortino.addText([
        { text: "Langkah Kalkulasi:\n", options: { bold: true, color: "8e44ad", fontSize: 12, breakLine: true } },
        { text: "1. Mean Return (Rp): ", options: { bold: true, fontSize: 9.5 } }, { text: "0.125%\n", options: { fontSize: 9.5 } },
        
        { text: "2. Downside Deviation (\u03c3d):\n", options: { bold: true, fontSize: 9.5 } },
        { text: "   Hanya hitung hari negatif terhadap Rf (0%):\n", options: { fontSize: 7.5, italic: true } },
        { text: "   \u2022 Loss 1: (-1.5-0)\u00b2 = 2.25 | Loss 2: (-0.5-0)\u00b2 = 0.25\n", options: { fontSize: 8 } },
        { text: "   \u2022 \u03c3d = \u221a[(2.25+0.25)/4] = ", options: { fontSize: 8 } }, { text: "0.79%\n\n", options: { bold: true, fontSize: 8 } },
        
        { text: "3. Sortino Ratio:\n", options: { bold: true, fontSize: 9.5 } },
        { text: "   (0.125 - 0) / 0.79 = ", options: { fontSize: 9.5 } }, { text: "0.16\n", options: { bold: true, color: "8e44ad", fontSize: 9.5 } },
        
        { text: "------------------------------------------\n", options: {} },
        { text: "Keterangan Istilah & Singkatan:\n", options: { bold: true, fontSize: 8.5, color: "8e44ad" } },
        { text: "\u2022 Rp: ", options: { bold: true, fontSize: 7.5 } }, { text: "Return Portofolio (Mean Return)\n", options: { fontSize: 7.5 } },
        { text: "\u2022 \u03c3d: ", options: { bold: true, fontSize: 7.5 } }, { text: "Downside Deviation (Hanya mengukur risiko kerugian, mengabaikan volatilitas positif)\n", options: { fontSize: 7.5 } },
        { text: "\u2022 Rf: ", options: { bold: true, fontSize: 7.5 } }, { text: "Risk-Free Rate (Suku Bunga Bebas Risiko)\n", options: { fontSize: 7.5 } },
        
        { text: "------------------------------------------\n", options: {} },
        { text: "Apa yang digambarkan Sortino?\n", options: { bold: true, fontSize: 9, color: "8e44ad" } },
        { text: "Mengukur efisiensi terhadap KERUGIAN. ", options: { fontSize: 7.5 } },
        { text: "Semakin besar semakin bagus ", options: { fontSize: 7.5, bold: true, color: "27ae60" } },
        { text: "(profit lebih besar dibanding risiko drawdown).", options: { fontSize: 7.5 } }
    ], { x: 6.4, y: 1.1, w: 3.2, color: "333333", valign: "top" });

    slideSortino.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Slide 8: Simulasi Calmar Ratio ---
    let slideCalmar = pres.addSlide();
    if (fs.existsSync("bg_watermark.png")) {
        slideCalmar.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    }
    if (fs.existsSync("logo_unm.png")) {
        slideCalmar.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    }

    slideCalmar.addText("Simulasi Performa: Calmar Ratio", { x: 0.5, y: 0.4, w: "90%", fontSize: 26, bold: true, color: "003366" });
    slideCalmar.addText("Fokus pada Maximum Drawdown (Penurunan Terparah):", { x: 0.5, y: 1.0, w: "90%", fontSize: 14, color: "333333" });

    // Tabel Nilai Portofolio (Simulasi 12 Bulan)
    slideCalmar.addTable([
        [
            { text: "Bulan", options: { bold: true, fill: "003366", color: "ffffff", align: "center" } },
            { text: "Nilai Portofolio", options: { bold: true, fill: "003366", color: "ffffff", align: "center" } },
            { text: "Keterangan", options: { bold: true, fill: "003366", color: "ffffff", align: "center" } }
        ],
        ["Bulan 1", "Rp 100 Juta", "Mulai Investasi"],
        ["Bulan 4", "Rp 120 Juta", { text: "Puncak Tertinggi (Peak)", options: { color: "27ae60", bold: true } }],
        ["Bulan 8", "Rp 90 Juta", { text: "Titik Terendah (Trough)", options: { color: "c0392b", bold: true } }],
        ["Bulan 12", "Rp 135 Juta", "Hasil Akhir Tahun"]
    ], { x: 0.5, y: 1.6, w: 5.5, fontSize: 11, border: { pt: 1, color: "dddddd" }, align: "center", valign: "middle" });

    // Box Penjelasan Calmar
    slideCalmar.addShape(pres.ShapeType.rect, { x: 6.2, y: 1.0, w: 3.5, h: 4.4, fill: { color: "ebf5fb" }, line: { color: "2980b9", width: 1.5 } });
    slideCalmar.addText([
        { text: "Kalkulasi (Data 12 Bulan):\n", options: { bold: true, color: "2980b9", fontSize: 12, breakLine: true } },
        { text: "1. Annual Return (Rp):\n", options: { bold: true, fontSize: 9.5 } },
        { text: "(135 - 100) / 100 = ", options: { fontSize: 9 } }, { text: "35%\n\n", options: { bold: true, fontSize: 9 } },
        
        { text: "2. Maximum Drawdown (MDD):\n", options: { bold: true, fontSize: 9.5 } },
        { text: "Penurunan terburuk dalam 12 bln:\n", options: { fontSize: 7.5, italic: true } },
        { text: "(120 - 90) / 120 = ", options: { fontSize: 9 } }, { text: "25%\n\n", options: { bold: true, color: "c0392b", fontSize: 9 } },
        
        { text: "3. Calmar Ratio:\n", options: { bold: true, fontSize: 9.5 } },
        { text: "Annual Return / MDD = 35% / 25% = ", options: { fontSize: 9 } }, { text: "1.40\n", options: { bold: true, color: "27ae60", fontSize: 9 } },
        
        { text: "------------------------------------------\n", options: {} },
        { text: "Keterangan Istilah & Singkatan:\n", options: { bold: true, fontSize: 8.5, color: "2980b9" } },
        { text: "\u2022 Rp: ", options: { bold: true, fontSize: 7.5 } }, { text: "Return Portofolio (Annual Return / Imbal Hasil Tahunan)\n", options: { fontSize: 7.5 } },
        { text: "\u2022 MDD: ", options: { bold: true, fontSize: 7.5 } }, { text: "Maximum Drawdown (Penurunan terbesar dari titik puncak ke titik terendah sebelum puncak baru)\n", options: { fontSize: 7.5 } },
        
        { text: "------------------------------------------\n", options: {} },
        { text: "Apa yang digambarkan Calmar?\n", options: { bold: true, fontSize: 9, color: "003366" } },
        { text: "Mengukur profit tahunan dibanding risiko jatuh terdalam. ", options: { fontSize: 7.5 } },
        { text: "Semakin besar semakin bagus ", options: { fontSize: 7.5, bold: true, color: "27ae60" } },
        { text: "(imbal hasil melampaui kerugian historis terparah).", options: { fontSize: 7.5 } }
    ], { x: 6.4, y: 1.1, w: 3.2, color: "333333", valign: "top" });

    slideCalmar.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Slide 9: Simulasi Ulcer Index ---
    let slideUlcer = pres.addSlide();
    if (fs.existsSync("bg_watermark.png")) {
        slideUlcer.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    }
    if (fs.existsSync("logo_unm.png")) {
        slideUlcer.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    }

    slideUlcer.addText("Simulasi Performa: Ulcer Index (UI)", { x: 0.5, y: 0.4, w: "90%", fontSize: 26, bold: true, color: "003366" });
    slideUlcer.addText("Mengukur 'Stres' Investasi (Kedalaman & Durasi Penurunan):", { x: 0.5, y: 1.0, w: "90%", fontSize: 14, color: "333333" });

    // Tabel Drawdown Harian
    slideUlcer.addTable([
        [
            { text: "Hari", options: { bold: true, fill: "003366", color: "ffffff", align: "center" } },
            { text: "Nilai", options: { bold: true, fill: "003366", color: "ffffff", align: "center" } },
            { text: "Drawdown (DD)", options: { bold: true, fill: "003366", color: "ffffff", align: "center" } },
            { text: "DD Kuadrat (DD\u00b2)", options: { bold: true, fill: "003366", color: "ffffff", align: "center" } }
        ],
        ["Day 1", "100", "0%", "0"],
        ["Day 2", "95", "5%", "25"],
        ["Day 3", "92", "8%", "64"],
        ["Day 4", "98", "2%", "4"],
        ["Day 5", "105", "0%", "0"]
    ], { x: 0.5, y: 1.6, w: 5.5, fontSize: 10, border: { pt: 1, color: "dddddd" }, align: "center", valign: "middle" });

    // Box Penjelasan Ulcer
    slideUlcer.addShape(pres.ShapeType.rect, { x: 6.2, y: 1.0, w: 3.5, h: 4.4, fill: { color: "fdf2e9" }, line: { color: "e67e22", width: 1.5 } });
    slideUlcer.addText([
        { text: "Langkah Kalkulasi:\n", options: { bold: true, color: "e67e22", fontSize: 12, breakLine: true } },
        { text: "1. Kumpulkan Data Drawdown:\n", options: { bold: true, fontSize: 9.5 } },
        { text: "Persentase penurunan dari puncak terakhir di setiap titik waktu.\n\n", options: { fontSize: 7.5, italic: true } },
        
        { text: "2. Rata-rata Kuadrat (Mean Sq):\n", options: { bold: true, fontSize: 9.5 } },
        { text: "(0 + 25 + 64 + 4 + 0) / 5 = ", options: { fontSize: 9 } }, { text: "18.6\n\n", options: { bold: true, fontSize: 9 } },
        
        { text: "3. Ulcer Index (UI):\n", options: { bold: true, fontSize: 9.5 } },
        { text: "\u221a18.6 = ", options: { fontSize: 9 } }, { text: "4.31%\n\n", options: { bold: true, color: "c0392b", fontSize: 9 } },
        
        { text: "------------------------------------------\n", options: {} },
        { text: "Keterangan Istilah & Singkatan:\n", options: { bold: true, fontSize: 8.5, color: "e67e22" } },
        { text: "\u2022 Drawdown (DD): ", options: { bold: true, fontSize: 7.5 } }, { text: "Penurunan dari puncak sebelumnya\n", options: { fontSize: 7.5 } },
        { text: "\u2022 Mean Sq: ", options: { bold: true, fontSize: 7.5 } }, { text: "Rata-rata kuadrat dari seluruh nilai Drawdown\n", options: { fontSize: 7.5 } },
        { text: "\u2022 UI: ", options: { bold: true, fontSize: 7.5 } }, { text: "Ulcer Index (Mengukur kedalaman & durasi stres/penurunan)\n", options: { fontSize: 7.5 } },
        
        { text: "------------------------------------------\n", options: {} },
        { text: "Interpretasi Khusus:\n", options: { bold: true, fontSize: 9, color: "003366" } },
        { text: "Untuk Ulcer Index: ", options: { fontSize: 7.5 } },
        { text: "Semakin KECIL semakin bagus. ", options: { fontSize: 7.5, bold: true, color: "27ae60" } },
        { text: "UI rendah berarti portofolio jarang mengalami penurunan dalam/lama (minim stres).", options: { fontSize: 7.5 } }
    ], { x: 6.4, y: 1.1, w: 3.2, color: "333333", valign: "top" });

    slideUlcer.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

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
