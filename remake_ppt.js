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
        { text: "Landasan Teori (Markowitz)", options: { bullet: { code: '25CF' }, color: "0563C1", underline: true, hyperlink: { slide: '6' } } },
        { text: "Kerangka Pemikiran", options: { bullet: { code: '25CF' }, color: "333333" } }
    ], { x: col1X + 0.3, y: row2Y + 0.5, w: cardW - 0.6, h: 1.2, fontSize: 11, lineSpacing: 22, valign: "top" });

    // --- CARD III: METODOLOGI ---
    slideTOC.addShape(pres.ShapeType.rect, { x: col2X, y: row1Y, w: cardW, h: 1.8, fill: { color: "ffffff" }, line: { color: "2980b9", width: 1.5 }, rectRadius: 0.05 });
    slideTOC.addShape(pres.ShapeType.rect, { x: col2X, y: row1Y, w: cardW, h: headerH, fill: { color: "2980b9" } });
    slideTOC.addText("III. METODOLOGI & SAC-NET", { x: col2X, y: row1Y, w: cardW, h: headerH, fontSize: 14, bold: true, color: "ffffff", align: "center", valign: "middle" });
    slideTOC.addText([
        { text: "Network Markowitz", options: { bullet: { code: '25CF' }, color: "0563C1", underline: true, hyperlink: { slide: '7' } } },
        { text: "Fitur Observasi SAC", options: { bullet: { code: '25CF' }, color: "0563C1", underline: true, hyperlink: { slide: '8' } } },
        { text: "Parameter Eksperimen", options: { bullet: { code: '25CF' }, color: "333333" } }
    ], { x: col2X + 0.3, y: row1Y + 0.5, w: cardW - 0.6, h: 1.2, fontSize: 11, lineSpacing: 22, valign: "top" });

    // --- CARD IV: EVALUASI ---
    slideTOC.addShape(pres.ShapeType.rect, { x: col2X, y: row2Y, w: cardW, h: 1.8, fill: { color: "ffffff" }, line: { color: "8e44ad", width: 1.5 }, rectRadius: 0.05 });
    slideTOC.addShape(pres.ShapeType.rect, { x: col2X, y: row2Y, w: cardW, h: headerH, fill: { color: "8e44ad" } });
    slideTOC.addText("IV. EVALUASI & HASIL", { x: col2X, y: row2Y, w: cardW, h: headerH, fontSize: 14, bold: true, color: "ffffff", align: "center", valign: "middle" });
    slideTOC.addText([
        { text: "Evaluasi Portofolio (Ratio)", options: { bullet: { code: '25CF' }, color: "0563C1", underline: true, hyperlink: { slide: '10' } } },
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
    slideProb.addText("Bagaimana agen SAC mampu mengoptimasi parameter penalti sentralitas secara dinamis berdasarkan kondisi jaringan dan momentum pasar?", { x: cardStartX + sideW + 0.2, y: cardStartY, w: bodyW - 0.4, h: cardH, fontSize: 14, bold: true, color: "2c3e50", valign: "middle" });

    // Q2
    slideProb.addShape(pres.ShapeType.rect, { x: cardStartX, y: cardStartY + cardGap, w: sideW + bodyW, h: cardH, fill: { color: "ffffff" }, line: { color: "2c3e50", width: 1.5 } });
    slideProb.addShape(pres.ShapeType.rect, { x: cardStartX, y: cardStartY + cardGap, w: sideW, h: cardH, fill: { color: "2c3e50" } });
    slideProb.addText("Q2", { x: cardStartX, y: cardStartY + cardGap, w: sideW, h: cardH, fontSize: 22, bold: true, color: "ffffff", align: "center", valign: "middle" });
    slideProb.addText("Sejauh mana penggunaan fitur gabungan (network & market metrics) meningkatkan ketahanan portofolio terhadap volatilitas ekstrem?", { x: cardStartX + sideW + 0.2, y: cardStartY + cardGap, w: bodyW - 0.4, h: cardH, fontSize: 14, bold: true, color: "2c3e50", valign: "middle" });

    // Q3
    slideProb.addShape(pres.ShapeType.rect, { x: cardStartX, y: cardStartY + (cardGap * 2), w: sideW + bodyW, h: cardH, fill: { color: "ffffff" }, line: { color: "e67e22", width: 1.5 } });
    slideProb.addShape(pres.ShapeType.rect, { x: cardStartX, y: cardStartY + (cardGap * 2), w: sideW, h: cardH, fill: { color: "e67e22" } });
    slideProb.addText("Q3", { x: cardStartX, y: cardStartY + (cardGap * 2), w: sideW, h: cardH, fontSize: 22, bold: true, color: "ffffff", align: "center", valign: "middle" });
    slideProb.addText("Apakah model SAC-Net Markowitz mampu melampaui performa benchmark pada metrik Sharpe dan Calmar Ratio?", { x: cardStartX + sideW + 0.2, y: cardStartY + (cardGap * 2), w: bodyW - 0.4, h: cardH, fontSize: 14, bold: true, color: "2c3e50", valign: "middle" });

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
    slideTujuan.addText("Merancang dan melatih agen SAC yang bertindak sebagai Gamma Controller untuk mengoptimasi parameter penalti sentralitas secara dinamis.", { x: tCardStartX + tSideW + 0.2, y: tCardStartY, w: tBodyW - 0.4, h: tCardH, fontSize: 14, bold: true, color: "2c3e50", valign: "middle" });

    // Obj 2
    slideTujuan.addShape(pres.ShapeType.rect, { x: tCardStartX, y: tCardStartY + tCardGap, w: tSideW + tBodyW, h: tCardH, fill: { color: "ffffff" }, line: { color: "2c3e50", width: 1.5 } });
    slideTujuan.addShape(pres.ShapeType.rect, { x: tCardStartX, y: tCardStartY + tCardGap, w: tSideW, h: tCardH, fill: { color: "2c3e50" } });
    slideTujuan.addText("Obj 2", { x: tCardStartX, y: tCardStartY + tCardGap, w: tSideW, h: tCardH, fontSize: 22, bold: true, color: "ffffff", align: "center", valign: "middle" });
    slideTujuan.addText("Mengevaluasi efektivitas penggunaan fitur gabungan (network metrics dan market indicators) dalam adaptasi portofolio.", { x: tCardStartX + tSideW + 0.2, y: tCardStartY + tCardGap, w: tBodyW - 0.4, h: tCardH, fontSize: 14, bold: true, color: "2c3e50", valign: "middle" });

    // Obj 3
    slideTujuan.addShape(pres.ShapeType.rect, { x: tCardStartX, y: tCardStartY + (tCardGap * 2), w: tSideW + tBodyW, h: tCardH, fill: { color: "ffffff" }, line: { color: "f39c12", width: 1.5 } });
    slideTujuan.addShape(pres.ShapeType.rect, { x: tCardStartX, y: tCardStartY + (tCardGap * 2), w: tSideW, h: tCardH, fill: { color: "f39c12" } });
    slideTujuan.addText("Obj 3", { x: tCardStartX, y: tCardStartY + (tCardGap * 2), w: tSideW, h: tCardH, fontSize: 22, bold: true, color: "ffffff", align: "center", valign: "middle" });
    slideTujuan.addText("Membandingkan performa model SAC-Net Markowitz terhadap strategi benchmark menggunakan metrik Sharpe dan Calmar Ratio.", { x: tCardStartX + tSideW + 0.2, y: tCardStartY + (tCardGap * 2), w: tBodyW - 0.4, h: tCardH, fontSize: 14, bold: true, color: "2c3e50", valign: "middle" });

    slideTujuan.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });


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

    // --- Slide 7: Fitur Observasi SAC (Part 1/2) ---
    let slideFeat1 = pres.addSlide();
    if (fs.existsSync("bg_watermark.png")) {
        slideFeat1.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    }
    if (fs.existsSync("logo_unm.png")) {
        slideFeat1.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    }

    slideFeat1.addText("Fitur Observasi SAC (1/2)", { x: 0.5, y: 0.4, w: "90%", fontSize: 30, bold: true, color: "003366" });
    slideFeat1.addShape(pres.ShapeType.line, { x: 0.5, y: 0.9, w: 1.5, h: 0, line: { color: "e67e22", width: 3 } });
    slideFeat1.addText("Network Features dan Market Volatility sebagai input kondisi sistemik:", { x: 0.5, y: 1.05, w: "90%", fontSize: 14, color: "333333" });

    const fCardY1 = 1.4;
    const fCardH1 = 3.5;
    const fHalfW1 = 4.35;
    const fColGap1 = 0.3;

    // --- CARD 1: NETWORK ---
    slideFeat1.addShape(pres.ShapeType.rect, { x: 0.5, y: fCardY1, w: fHalfW1, h: fCardH1, fill: { color: "ffffff" }, line: { color: "27ae60", width: 1.5 }, rectRadius: 0.05 });
    slideFeat1.addShape(pres.ShapeType.rect, { x: 0.5, y: fCardY1, w: fHalfW1, h: 0.45, fill: { color: "27ae60" } });
    slideFeat1.addText("🌐 Network Features", { x: 0.5, y: fCardY1, w: fHalfW1, h: 0.45, fontSize: 16, bold: true, color: "ffffff", align: "center", valign: "middle" });
    slideFeat1.addText([
        { text: "1. MST.Dist (x0.1):", options: { bold: true, fontSize: 13, color: "003366", breakLine: true } },
        { text: "Total jarak Minimum Spanning Tree. Mengukur tingkat integrasi atau kerapatan korelasi antar aset kripto.\n", options: { fontSize: 11, breakLine: true } },
        { text: "2. Spectral.Gap:", options: { bold: true, fontSize: 13, color: "003366", breakLine: true } },
        { text: "Algebraic connectivity dari graf MST. Nilai kecil menandakan korelasi yang sangat kuat (fragile) dan rentan guncangan sistemik.\n", options: { fontSize: 11, breakLine: true } },
        { text: "📚 Referensi Dasar:", options: { bold: true, fontSize: 10, color: "27ae60", breakLine: true } },
        { text: "Mantegna (1999) - Hierarchical structure in financial markets.\nGiudici (2020) - Network-based risk in crypto.", options: { fontSize: 9, italic: true } }
    ], { x: 0.7, y: fCardY1 + 0.6, w: fHalfW1 - 0.4, h: 2.8, color: "333333", valign: "top", lineSpacing: 14 });

    // --- CARD 2: VOLATILITY ---
    slideFeat1.addShape(pres.ShapeType.rect, { x: 0.5 + fHalfW1 + fColGap1, y: fCardY1, w: fHalfW1, h: fCardH1, fill: { color: "ffffff" }, line: { color: "2980b9", width: 1.5 }, rectRadius: 0.05 });
    slideFeat1.addShape(pres.ShapeType.rect, { x: 0.5 + fHalfW1 + fColGap1, y: fCardY1, w: fHalfW1, h: 0.45, fill: { color: "2980b9" } });
    slideFeat1.addText("📈 Market Volatility", { x: 0.5 + fHalfW1 + fColGap1, y: fCardY1, w: fHalfW1, h: 0.45, fontSize: 16, bold: true, color: "ffffff", align: "center", valign: "middle" });
    slideFeat1.addText([
        { text: "1. VolShort (x100):", options: { bold: true, fontSize: 13, color: "003366", breakLine: true } },
        { text: "Rata-rata deviasi standar harga jendela 5 hari terakhir.\n", options: { fontSize: 10, breakLine: true } },
        { text: "2. VolLong (x100):", options: { bold: true, fontSize: 13, color: "003366", breakLine: true } },
        { text: "Rata-rata deviasi standar harga jendela 20 hari terakhir.\n", options: { fontSize: 10, breakLine: true } },
        { text: "3. Vol.Ratio:", options: { bold: true, fontSize: 13, color: "003366", breakLine: true } },
        { text: "Rasio volatilitas pendek terhadap panjang. Digunakan untuk membedakan tren stabil vs lonjakan noise.\n", options: { fontSize: 10, breakLine: true } },
        { text: "📚 Referensi Dasar:", options: { bold: true, fontSize: 10, color: "2980b9", breakLine: true } },
        { text: "Jiang et al. (2017) - Deep RL Framework for Portfolio Management.", options: { fontSize: 9, italic: true } }
    ], { x: 0.5 + fHalfW1 + fColGap1 + 0.2, y: fCardY1 + 0.6, w: fHalfW1 - 0.4, h: 2.8, color: "333333", valign: "top", lineSpacing: 11 });

    slideFeat1.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });


    // --- Slide 8: Fitur Observasi SAC (Part 2/2) ---
    let slideFeat2 = pres.addSlide();
    if (fs.existsSync("bg_watermark.png")) {
        slideFeat2.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    }
    if (fs.existsSync("logo_unm.png")) {
        slideFeat2.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    }

    slideFeat2.addText("Fitur Observasi SAC (2/2)", { x: 0.5, y: 0.4, w: "90%", fontSize: 30, bold: true, color: "003366" });
    slideFeat2.addShape(pres.ShapeType.line, { x: 0.5, y: 0.9, w: 1.5, h: 0, line: { color: "e67e22", width: 3 } });
    slideFeat2.addText("Market Momentum dan Ringkasan Dimensi State SAC Agent:", { x: 0.5, y: 1.05, w: "90%", fontSize: 14, color: "333333" });

    // --- CARD 3: MOMENTUM ---
    slideFeat2.addShape(pres.ShapeType.rect, { x: 0.5, y: 1.4, w: 9.0, h: 2.7, fill: { color: "ffffff" }, line: { color: "8e44ad", width: 1.5 }, rectRadius: 0.05 });
    slideFeat2.addShape(pres.ShapeType.rect, { x: 0.5, y: 1.4, w: 9.0, h: 0.45, fill: { color: "8e44ad" } });
    slideFeat2.addText("🚀 Market Momentum", { x: 0.5, y: 1.4, w: 9.0, h: 0.45, fontSize: 16, bold: true, color: "ffffff", align: "center", valign: "middle" });
    
    slideFeat2.addText([
        { text: "• Mom5d (x100):", options: { bold: true, fontSize: 12, color: "003366", breakLine: true } },
        { text: "Log-return aset jendela 5 hari.\n", options: { fontSize: 10, breakLine: true } },
        { text: "• Mom20d (x100):", options: { bold: true, fontSize: 12, color: "003366", breakLine: true } },
        { text: "Log-return aset jendela 20 hari.\n", options: { fontSize: 10, breakLine: true } },
        { text: "• MomCross (x100):", options: { bold: true, fontSize: 12, color: "003366", breakLine: true } },
        { text: "Selisih momentum pendek vs panjang.\n", options: { fontSize: 10, breakLine: true } },
        { text: "📚 Referensi Dasar:", options: { bold: true, fontSize: 10, color: "8e44ad", breakLine: true } },
        { text: "Jegadeesh & Titman (1993) - Returns to Buying Winners & Selling Losers.\nJiang et al. (2017) - DRL in Portfolio.", options: { fontSize: 9, italic: true } }
    ], { x: 0.8, y: 1.9, w: 4.0, h: 2.1, color: "333333", valign: "top", lineSpacing: 10 });

    slideFeat2.addText([
        { text: "• Pct.Uptrend:", options: { bold: true, fontSize: 12, color: "003366", breakLine: true } },
        { text: "Persentase koin dengan return positif.\n", options: { fontSize: 10, breakLine: true } },
        { text: "• State Dimension:", options: { bold: true, fontSize: 12, color: "003366", breakLine: true } },
        { text: "Total 9 fitur numerik vektor observasi.\n", options: { fontSize: 10, breakLine: true } },
        { text: "• State Array: ", options: { bold: true, fontSize: 12, color: "003366" } },
        { text: "[MST, Spec, V5, V20, VR, M5, M20, MC, PU]\n", options: { fontSize: 9, italic: true, breakLine: true } },
        { text: "💡 Intuisi Agen SAC:", options: { bold: true, fontSize: 10, color: "8e44ad", breakLine: true } },
        { text: "Agen belajar memetakan ke-9 dimensi state ini ke aksi penyesuaian gamma (\u03b3) dinamis.", options: { fontSize: 9, italic: true } }
    ], { x: 5.2, y: 1.9, w: 4.0, h: 2.1, color: "333333", valign: "top", lineSpacing: 10 });

    // Summary Box Bottom
    slideFeat2.addShape(pres.ShapeType.rect, { x: 0.5, y: 4.2, w: 9.0, h: 0.7, fill: { color: "f4f6f7" }, line: { color: "003366", width: 1.0 }, rectRadius: 0.05 });
    slideFeat2.addText("Total State Space: \u211d\u2079 (9-dimensional continuous observation vector)", { x: 0.5, y: 4.2, w: 9.0, h: 0.7, fontSize: 14, bold: true, color: "003366", align: "center", valign: "middle" });

    slideFeat2.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Slide 9: Simulasi Perhitungan Sharpe Ratio ---
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
        { text: "Langkah Kalkulasi:\n", options: { bold: true, color: "d35400", fontSize: 13, breakLine: true } },
        { text: "1. Mean Return (Rp): ", options: { bold: true, fontSize: 10 } }, { text: "0.65%\n", options: { fontSize: 10 } },
        
        { text: "2. Kalkulasi Risk (\u03c3p):\n", options: { bold: true, fontSize: 10 } },
        { text: "   \u2022 Dev\u00b2 Day 1: (0.7-0.65)\u00b2 = 0.0025\n", options: { fontSize: 9 } },
        { text: "   \u2022 Dev\u00b2 Day 2: (1.1-0.65)\u00b2 = 0.2025\n", options: { fontSize: 9 } },
        { text: "   \u2022 Dev\u00b2 Day 3: (0.2-0.65)\u00b2 = 0.2025\n", options: { fontSize: 9 } },
        { text: "   \u2022 Dev\u00b2 Day 4: (0.6-0.65)\u00b2 = 0.0025\n", options: { fontSize: 9 } },
        { text: "   \u2022 Sum Dev\u00b2 = 0.41\n", options: { fontSize: 9 } },
        { text: "   \u2022 Var = 0.41 / (4-1) = 0.137\n", options: { fontSize: 9 } },
        { text: "   \u2022 \u03c3p = \u221a0.137 = ", options: { fontSize: 9 } }, { text: "0.37%\n\n", options: { bold: true } },
        
        { text: "3. Sharpe Ratio (Rf = 5% Ann):\n", options: { bold: true, fontSize: 10 } },
        { text: "   (0.65 - 0.014) / 0.37 = ", options: { fontSize: 10 } }, { text: "1.72\n", options: { bold: true, color: "27ae60" } },
        
        { text: "------------------------------------------\n", options: {} },
        { text: "Apa yang digambarkan Sharpe?\n", options: { bold: true, fontSize: 10, color: "c0392b" } },
        { text: "Sharpe mengukur efisiensi: Seberapa banyak 'keuntungan ekstra' untuk setiap 1 unit risiko. ", options: { fontSize: 8.5 } },
        { text: "Semakin besar nilainya semakin bagus ", options: { fontSize: 8.5, bold: true, color: "27ae60" } },
        { text: "(karena return lebih besar dari risikonya).", options: { fontSize: 8.5 } }
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
        { text: "Langkah Kalkulasi:\n", options: { bold: true, color: "8e44ad", fontSize: 13, breakLine: true } },
        { text: "1. Mean Return (Rp): ", options: { bold: true, fontSize: 10 } }, { text: "0.125%\n", options: { fontSize: 10 } },
        
        { text: "2. Downside Deviation (\u03c3d):\n", options: { bold: true, fontSize: 10 } },
        { text: "   Hanya hitung hari negatif terhadap Rf (0%):\n", options: { fontSize: 8, italic: true } },
        { text: "   \u2022 Loss 1: (-1.5-0)\u00b2 = 2.25\n", options: { fontSize: 9 } },
        { text: "   \u2022 Loss 2: (-0.5-0)\u00b2 = 0.25\n", options: { fontSize: 9 } },
        { text: "   \u2022 \u03c3d = \u221a[(2.25+0.25)/4] = ", options: { fontSize: 9 } }, { text: "0.79%\n\n", options: { bold: true } },
        
        { text: "3. Sortino Ratio:\n", options: { bold: true, fontSize: 10 } },
        { text: "   (0.125 - 0) / 0.79 = ", options: { fontSize: 10 } }, { text: "0.16\n", options: { bold: true, color: "8e44ad" } },
        
        { text: "------------------------------------------\n", options: {} },
        { text: "Apa yang digambarkan Sortino?\n", options: { bold: true, fontSize: 10, color: "8e44ad" } },
        { text: "Sortino mengukur efisiensi terhadap KERUGIAN. Ini menggambarkan kemampuan portofolio memberikan imbal hasil dengan mengabaikan volatilitas positif. ", options: { fontSize: 8.5 } },
        { text: "Semakin besar semakin bagus ", options: { fontSize: 8.5, bold: true, color: "27ae60" } },
        { text: "karena artinya profit yang didapat jauh lebih besar dibanding risiko jatuhnya harga.", options: { fontSize: 8.5 } }
    ], { x: 6.4, y: 1.2, w: 3.2, color: "333333", valign: "top" });

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
        { text: "Kalkulasi (Data 12 Bulan):\n", options: { bold: true, color: "2980b9", fontSize: 13, breakLine: true } },
        { text: "1. Annual Return (Rp):\n", options: { bold: true, fontSize: 10 } },
        { text: "(135 - 100) / 100 = ", options: { fontSize: 10 } }, { text: "35%\n\n", options: { bold: true } },
        
        { text: "2. Maximum Drawdown (MDD):\n", options: { bold: true, fontSize: 10 } },
        { text: "Penurunan terburuk dalam 12 bln:\n", options: { fontSize: 8, italic: true } },
        { text: "(120 - 90) / 120 = ", options: { fontSize: 10 } }, { text: "25%\n\n", options: { bold: true, color: "c0392b" } },
        
        { text: "3. Calmar Ratio:\n", options: { bold: true, fontSize: 10 } },
        { text: "Annual Return / MDD\n", options: { fontSize: 9 } },
        { text: "35% / 25% = ", options: { fontSize: 10 } }, { text: "1.40\n", options: { bold: true, color: "27ae60" } },
        
        { text: "------------------------------------------\n", options: {} },
        { text: "Apa yang digambarkan Calmar?\n", options: { bold: true, fontSize: 10, color: "003366" } },
        { text: "Mengukur perbandingan antara profit tahunan dengan risiko 'jatuh' terdalam. ", options: { fontSize: 8.5 } },
        { text: "Semakin besar semakin bagus ", options: { fontSize: 8.5, bold: true, color: "27ae60" } },
        { text: "karena imbal hasil jauh melampaui riwayat kerugian terparahnya.", options: { fontSize: 8.5 } }
    ], { x: 6.4, y: 1.2, w: 3.2, color: "333333", valign: "top" });

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
        { text: "Langkah Kalkulasi:\n", options: { bold: true, color: "e67e22", fontSize: 13, breakLine: true } },
        { text: "1. Kumpulkan Data Drawdown:\n", options: { bold: true, fontSize: 10 } },
        { text: "Persentase penurunan dari puncak terakhir di setiap titik waktu.\n\n", options: { fontSize: 8, italic: true } },
        
        { text: "2. Rata-rata Kuadrat (Mean Sq):\n", options: { bold: true, fontSize: 10 } },
        { text: "(0 + 25 + 64 + 4 + 0) / 5 = ", options: { fontSize: 10 } }, { text: "18.6\n\n", options: { bold: true } },
        
        { text: "3. Ulcer Index (UI):\n", options: { bold: true, fontSize: 10 } },
        { text: "\u221a18.6 = ", options: { fontSize: 10 } }, { text: "4.31%\n\n", options: { bold: true, color: "c0392b" } },
        
        { text: "------------------------------------------\n", options: {} },
        { text: "Interpretasi Khusus:\n", options: { bold: true, fontSize: 10, color: "003366" } },
        { text: "Berbeda dengan metrik lain, untuk Ulcer Index: ", options: { fontSize: 8.5 } },
        { text: "Semakin KECIL semakin bagus. ", options: { fontSize: 8.5, bold: true, color: "27ae60" } },
        { text: "UI rendah berarti portofolio jarang mengalami penurunan yang dalam atau lama (minim 'stres').", options: { fontSize: 8.5 } }
    ], { x: 6.4, y: 1.2, w: 3.2, color: "333333", valign: "top" });

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
