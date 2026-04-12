const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

async function createPresentation() {
    console.log("Memulai pembuatan presentasi sesuai hasil riset...");

    // Inisialisasi presentasi baru
    let pres = new pptxgen();

    // Set layout (opsional, defaulnya 16x9)
    pres.layout = "LAYOUT_16x9";

    // --- Slide 1: Judul ---
    let slide1 = pres.addSlide();
    slide1.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slide1.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slide1.addText("Proposal Tesis:\nOptimalisasi Portofolio Cryptocurrency", {
        x: 0.5, y: 1.2, w: "90%", fontSize: 36, bold: true, align: "center", color: "003366"
    });
    slide1.addText([
        { text: "Berbasis Pendekatan\n" },
        { text: "Network Markowitz", options: { italic: true } },
        { text: " dengan " },
        { text: "2-Stage Tuning Parameter", options: { italic: true } }
    ], {
        x: 0.5, y: 2.6, w: "90%", fontSize: 24, align: "center", color: "34495e"
    });

    slide1.addText("Oleh: Ragil Yulianto", {
        x: 0.5, y: 4.5, w: "90%", fontSize: 18, align: "center", color: "7f8c8d"
    });

    // --- Slide 2: Daftar Isi (Bagian 1: Utama) ---
    let slideTOC1 = pres.addSlide();
    slideTOC1.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slideTOC1.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });
    slideTOC1.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slideTOC1.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slideTOC1.addText("Daftar Isi (Main Sections)", { x: 0.5, y: 0.5, w: "90%", fontSize: 28, bold: true, color: "003366" });

    // Kolom Kiri: Pendahuluan & Strategi
    slideTOC1.addText([
        { text: "I. PENDAHULUAN & MASALAH", options: { bold: true, color: "003366", breakLine: true } },
        { text: "   • ", options: {} },
        { text: "Latar Belakang", options: { hyperlink: { slide: '4' }, fontSize: 16 } },
        { text: "", options: { breakLine: true } },
        { text: "   • ", options: {} },
        { text: "Identifikasi Masalah (Risk)", options: { hyperlink: { slide: '5' }, fontSize: 16 } },
        { text: "", options: { breakLine: true } },
        { text: "   • ", options: {} },
        { text: "Landasan Teori", options: { hyperlink: { slide: '6' }, fontSize: 16 } },
        { text: "", options: { breakLine: true } },
        { text: "   • ", options: {} },
        { text: "Penelitian Terdahulu", options: { hyperlink: { slide: '7' }, fontSize: 16 } },
        { text: "", options: { breakLine: true } },
        { text: "   • ", options: {} },
        { text: "Kerangka Pemikiran", options: { hyperlink: { slide: '8' }, fontSize: 16 } },
        { text: "", options: { breakLine: true } },
        { text: "   • ", options: {} },
        { text: "Dataset Penelitian", options: { hyperlink: { slide: '9' }, fontSize: 16 } },
        { text: "", options: { breakLine: true, breakLine: true } },

        { text: "II. STRATEGI PORTOFOLIO", options: { bold: true, color: "003366", breakLine: true } },
        { text: "   • ", options: {} },
        { text: "Equally Weighted (EW)", options: { hyperlink: { slide: '11' }, fontSize: 16 } },
        { text: "", options: { breakLine: true } },
        { text: "   • ", options: {} },
        { text: "Classical Markowitz (CM)", options: { hyperlink: { slide: '12' }, fontSize: 16 } },
        { text: "", options: { breakLine: true } },
        { text: "   • ", options: {} },
        { text: "Graphical Lasso (GM)", options: { hyperlink: { slide: '13' }, fontSize: 16 } },
        { text: "", options: { breakLine: true } },
        { text: "   • ", options: {} },
        { text: "Network Markowitz (Statis)", options: { hyperlink: { slide: '14' }, fontSize: 16 } },
        { text: "", options: { breakLine: true } },
        { text: "   • ", options: {} },
        { text: "Network Markowitz (2-Stage GS)", options: { hyperlink: { slide: '15' }, fontSize: 16 } }
    ], { x: 0.5, y: 1.1, w: "45%", h: 5, fontSize: 18, color: "333333", valign: "top" });

    // Kolom Kanan: Evaluasi & Navigasi
    slideTOC1.addText([
        { text: "III. EVALUASI PERFORMA", options: { bold: true, color: "003366", breakLine: true } },
        { text: "   • ", options: {} },
        { text: "P&L, Sharpe, VaR, Rachev", options: { hyperlink: { slide: '16' }, fontSize: 16 } },
        { text: "", options: { breakLine: true } },
        { text: "   • ", options: {} },
        { text: "Analisis per Fase Pasar", options: { hyperlink: { slide: '17' }, fontSize: 16 } },
        { text: "", options: { breakLine: true, breakLine: true } },

        { text: "Lanjutan Daftar Isi:", options: { bold: true, color: "e67e22", breakLine: true } },
        { text: "   ➤ ", options: {} },
        { text: "LAMPIRAN TEKNIS (APPENDIX)", options: { hyperlink: { slide: '3' }, fontSize: 16, color: "d35400", bold: true } },
        { text: "", options: { breakLine: true } },
        { text: "   ➤ ", options: {} },
        { text: "END (TERIMA KASIH)", options: { hyperlink: { slide: '61' }, fontSize: 16, color: "003366", bold: true } },
    ], { x: 5.2, y: 1.1, w: "45%", h: 5, fontSize: 18, color: "333333", valign: "top" });

    // --- Slide 3: Daftar Isi (Bagian 2: Appendix) ---
    let slideTOC2 = pres.addSlide();
    slideTOC2.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slideTOC2.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });
    slideTOC2.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slideTOC2.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slideTOC2.addText("Daftar Isi (IV. Lampiran Teknis)", { x: 0.5, y: 0.5, w: "90%", fontSize: 28, bold: true, color: "003366" });

    // Kolom Kiri
    slideTOC2.addText([
        { text: "IV. LAMPIRAN (BAGIAN 1)", options: { bold: true, color: "003366", breakLine: true } },
        { text: "   • ", options: {} },
        { text: "Simulasi Strategi EW", options: { hyperlink: { slide: '32' }, fontSize: 16, breakLine: true } },
        { text: "   • ", options: {} },
        { text: "Simulasi Strategi CM", options: { hyperlink: { slide: '33' }, fontSize: 16, breakLine: true } },
        { text: "   • ", options: {} },
        { text: "Simulasi Strategi GLasso", options: { hyperlink: { slide: '34' }, fontSize: 16, breakLine: true } },
        { text: "   • ", options: {} },
        { text: "Simulasi Strategi NW", options: { hyperlink: { slide: '36' }, fontSize: 16, breakLine: true } },
        { text: "   • ", options: {} },
        { text: "Random Matrix Theory (RMT)", options: { hyperlink: { slide: '18' }, fontSize: 16, breakLine: true } },
        { text: "   • ", options: {} },
        { text: "Minimum Spanning Tree (MST)", options: { hyperlink: { slide: '25' }, fontSize: 16, breakLine: true } },
        { text: "   • ", options: {} },
        { text: "Penalty & Centrality Logic", options: { hyperlink: { slide: '27' }, fontSize: 16, breakLine: true } },
        { text: "   • ", options: {} },
        { text: "Justifikasi Parameter Gamma", options: { hyperlink: { slide: '28' }, fontSize: 16, breakLine: true } },
    ], { x: 0.5, y: 1.1, w: "45%", h: 5, fontSize: 16, color: "333333", valign: "top" });

    // Kolom Kanan
    slideTOC2.addText([
        { text: "IV. LAMPIRAN (BAGIAN 2)", options: { bold: true, color: "003366", breakLine: true } },
        { text: "   • ", options: {} },
        { text: "Simulasi 60/40 Weight Shift", options: { hyperlink: { slide: '37' }, fontSize: 16, breakLine: true } },
        { text: "   • ", options: {} },
        { text: "Simulasi 2-Stage Grid Search", options: { hyperlink: { slide: '49' }, fontSize: 16, breakLine: true } },
        { text: "   • ", options: {} },
        { text: "Justifikasi Rolling Window", options: { hyperlink: { slide: '50' }, fontSize: 16, breakLine: true } },
        { text: "   • ", options: {} },
        { text: "Simulasi Cumulative P&L", options: { hyperlink: { slide: '57' }, fontSize: 16, breakLine: true } },
        { text: "   • ", options: {} },
        { text: "Simulasi Value at Risk (VaR)", options: { hyperlink: { slide: '58' }, fontSize: 16, breakLine: true } },
        { text: "   • ", options: {} },
        { text: "Simulasi Sharpe Ratio", options: { hyperlink: { slide: '59' }, fontSize: 16, breakLine: true } },
        { text: "   • ", options: {} },
        { text: "Simulasi Rachev Ratio", options: { hyperlink: { slide: '60' }, fontSize: 16, breakLine: true } },
    ], { x: 5.2, y: 1.1, w: "45%", h: 5, fontSize: 16, color: "333333", valign: "top" });
    slideTOC2.addText("🏠 Kembali ke Daftar Isi Utama", { x: 7.0, y: 5.3, w: 2.7, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    let slide2 = pres.addSlide();
    slide2.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slide2.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slide2.addText("Latar Belakang", { x: 0.5, y: 0.5, w: "90%", fontSize: 28, bold: true, color: "003366" });
    // Judul tetap di atas

    // Area Kiri: Masalah (Padding diperbaiki)
    slide2.addShape(pres.ShapeType.rect, { x: 0.5, y: 1.5, w: 4.2, h: 3.2, fill: { color: "ffffff" }, line: { color: "003366", width: 2 } });
    slide2.addText("🔍 IDENTIFIKASI MASALAH", { x: 0.5, y: 1.8, w: 4.2, fontSize: 18, bold: true, color: "003366", align: "center" });
    slide2.addText([
        { text: "• Volatilitas Kripto & Noise:", options: { bold: true, fontSize: 14, breakLine: true } },
        { text: "   Gangguan data yang mengaburkan sinyal asli.", options: { fontSize: 13, breakLine: true } },
        { text: "• Optimalitas Window:", options: { bold: true, fontSize: 14, breakLine: true } },
        { text: "   Belum ada standar panjang jendela observasi.", options: { fontSize: 13 } }
    ], { x: 0.7, y: 2.6, w: 3.8, color: "333333", valign: "top" });

    // Area Kanan: Gap & Solusi (Padding diperbaiki)
    slide2.addShape(pres.ShapeType.rect, { x: 5.3, y: 1.5, w: 4.2, h: 3.2, fill: { color: "ffffff" }, line: { color: "27ae60", width: 2 } });
    slide2.addText("💡 RESEARCH GAP & SOLUSI", { x: 5.3, y: 1.8, w: 4.2, fontSize: 18, bold: true, color: "27ae60", align: "center" });
    slide2.addText([
        { text: "• Ketidakpastian Penalti (γ):", options: { bold: true, fontSize: 14, breakLine: true } },
        { text: "   Belum ditemukan bobot ideal secara universal.", options: { fontSize: 13, breakLine: true } },
        { text: "• Urgensi Tuning Parameter:", options: { bold: true, fontSize: 14, color: "27ae60", breakLine: true } },
        { text: "   Mekanisme kalibrasi sistematis diperlukan.", options: { italic: true, fontSize: 13 } }
    ], { x: 5.5, y: 2.6, w: 3.8, color: "333333", valign: "top" });
    slide2.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slide2.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Slide 4: Landasan Teori (Dua Kolom) ---
    let slide4 = pres.addSlide();
    slide4.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slide4.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slide4.addText("Landasan Teori Utama", { x: 0.5, y: 0.5, w: "90%", fontSize: 28, bold: true, color: "003366" });

    // Kolom Kiri: Portofolio & Risiko
    slide4.addText([
        { text: "Tinjauan Portofolio & Risiko:", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "➤ ", options: {} },
        { text: "Portofolio", options: { bold: true, underline: true } },
        { text: ": Diversifikasi aset untuk optimasi return-risiko.", options: { breakLine: true } },
        { text: "➤ ", options: {} },
        { text: "Volatilitas", options: { bold: true, underline: true } },
        { text: ": Ukuran fluktuasi harga pasar.", options: { breakLine: true } },
        { text: "➤ ", options: {} },
        { text: "Kovarians", options: { bold: true, underline: true } },
        { text: ": Ukuran pergerakan bersama aset.", options: { breakLine: true } }
    ], { x: 0.5, y: 1.2, w: "45%", h: 4, fontSize: 20, color: "333333", valign: "top" });

    // Kolom Kanan: Struktur Jaringan
    slide4.addText([
        { text: "Pendekatan Struktur Jaringan:", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "➤ ", options: {} },
        { text: "RMT (Random Matrix Theory)", options: { bold: true, underline: true } },
        { text: ": Filter noise untuk kestabilan matriks.", options: { breakLine: true } },
        { text: "➤ ", options: {} },
        { text: "MST (Minimum Spanning Tree)", options: { bold: true, underline: true } },
        { text: ": Jaringan korelasi terkuat tanpa loop.", options: { breakLine: true } },
        { text: "➤ ", options: {} },
        { text: "Centrality", options: { bold: true, underline: true } },
        { text: ": Metrik risiko penularan sistemik.", options: { breakLine: true } }
    ], { x: 5.2, y: 1.2, w: "45%", h: 4, fontSize: 20, color: "333333", valign: "top" });
    slide4.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slide4.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Slide 4.1: Penelitian Terdahulu ---
    let slidePrev = pres.addSlide();
    slidePrev.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slidePrev.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slidePrev.addText("Penelitian Terdahulu (State of the Art)", { x: 0.5, y: 0.5, w: "90%", fontSize: 28, bold: true, color: "003366" });
    slidePrev.addTable(
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
    });
    slidePrev.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slidePrev.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Slide 4.5: Kerangka Penelitian ---
    let slideFramework = pres.addSlide();
    slideFramework.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slideFramework.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slideFramework.addText("Kerangka Pemikiran / Penelitian", { x: 0.5, y: 0.5, w: "90%", fontSize: 28, bold: true, color: "003366" });
    slideFramework.addImage({ path: "propose_method_gs.drawio.png", x: 2.25, y: 0.85, w: 5.5, h: 4.4 });
    slideFramework.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slideFramework.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Slide 4.6: Dataset - 10 Aset Kripto Utama ---
    let slideData = pres.addSlide();
    slideData.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slideData.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slideData.addText("Dataset: 10 Aset Kripto Utama", { x: 0.5, y: 0.5, w: "90%", fontSize: 28, bold: true, color: "003366" });

    let tableRows = [
        [
            { text: "Ticker", options: { bold: true, fill: "003366", color: "ffffff", align: "center" } },
            { text: "Nama Aset", options: { bold: true, fill: "003366", color: "ffffff", align: "center" } },
            { text: "Kategori / Use Case", options: { bold: true, fill: "003366", color: "ffffff", align: "center" } }
        ],
        ["BTC", "Bitcoin", "Layer 1 / Store of Value"],
        ["ETH", "Ethereum", "Layer 1 / Smart Contract"],
        ["XRP", "Ripple", "Payment / Bridge Currency"],
        ["USDT", "Tether", "Stablecoin / USD Pegged"],
        ["BCH", "Bitcoin Cash", "Payment / Peer-to-Peer Cash"],
        ["LTC", "Litecoin", "Payment / Digital Silver"],
        ["BNB", "Binance Coin", "Layer 1 / Exchange Token"],
        ["EOS", "EOS", "Layer 1 / Smart Contract"],
        ["XLM", "Stellar", "Payment / Bridge Currency"],
        ["TRX", "Tron", "Layer 1 / Smart Contract"]
    ];

    slideData.addTable(tableRows, {
        x: 0.5, y: 0.9, w: 9.0,
        colWidths: [1.2, 2.5, 5.3],
        border: { type: "solid", color: "cccccc", pt: 1 },
        fontSize: 14,
        color: "333333"
    });

    // Statistik Dataset (Giudici et al. 2020 baseline)
    slideData.addText([
        { text: "Statistik Dataset: ", options: { bold: true, color: "003366" } },
        { text: "14 Sept 2017 - 17 Okt 2019 ", options: {} },
        { text: "| Total: ", options: { bold: true, color: "003366" } },
        { text: "764 observasi harian", options: {} }
    ], { x: 0.5, y: 5.1, w: "80%", fontSize: 12, color: "333333" });

    slideData.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slideData.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Slide 5: Strategi yang Dibandingkan ---
    let slide5 = pres.addSlide();
    slide5.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slide5.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slide5.addText("Strategi Portofolio yang Disimulasikan", { x: 0.5, y: 0.5, w: "90%", fontSize: 28, bold: true, color: "003366" });
    slide5.addText([
        { text: "1. Kelompok Baseline:", options: { bold: true, color: "003366", breakLine: true } },
        { text: "   • ", options: {} },
        { text: "EW (Equally Weighted)", options: { hyperlink: { slide: '11' }, color: "0563C1", underline: true } },
        { text: "", options: { breakLine: true } },
        { text: "   • ", options: {} },
        { text: "CM (Classical Markowitz)", options: { hyperlink: { slide: '12' }, color: "0563C1", underline: true } },
        { text: "", options: { breakLine: true, breakLine: true } },

        { text: "2. Kelompok Regularisasi:", options: { bold: true, color: "003366", breakLine: true } },
        { text: "   • ", options: {} },
        { text: "GM (Glasso Markowitz)", options: { hyperlink: { slide: '13' }, color: "0563C1", underline: true } },
        { text: "", options: { breakLine: true, breakLine: true } },

        { text: "3. Kelompok Network (Statis):", options: { bold: true, color: "003366", breakLine: true } },
        { text: "   • ", options: {} },
        { text: "NW Statis (γ fixed)", options: { hyperlink: { slide: '14' }, color: "0563C1", underline: true } },
        { text: "", options: { breakLine: true, breakLine: true } },

        { text: "4. Kelompok Network (Tuned / 2-Stage GS):", options: { bold: true, color: "003366", breakLine: true } },
        { text: "   • ", options: {} },
        { text: "NW 2-Stage GS (VAR, Sharpe, Rachev)", options: { hyperlink: { slide: '15' }, color: "0563C1", underline: true } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5, fontSize: 22, color: "333333", valign: "top" });
    slide5.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slide5.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Slide 5.1: Equally Weighted (EW) ---
    let slideEW = pres.addSlide();
    slideEW.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slideEW.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slideEW.addText("1.1. Equally Weighted (EW)", { x: 0.5, y: 0.5, w: "90%", fontSize: 28, bold: true, color: "003366" });
    slideEW.addText([
        { text: "Konsep Dasar:", options: { bold: true, breakLine: true } },
        { text: "Strategi alokasi ", options: { bullet: true } },
        { text: "1/N", options: { bold: true } },
        { text: " tanpa mempertimbangkan ", options: {} },
        { text: "parameter risiko", options: { bold: true } },
        { text: ".", options: { breakLine: true } },
        { text: "Keunggulan:", options: { bold: true, breakLine: true } },
        { text: "Berfungsi sebagai ", options: { bullet: true } },
        { text: "benchmark naif", options: { bold: true } },
        { text: " yang tangguh.", options: { breakLine: true } },
        { text: "Tidak memiliki ", options: { bullet: true } },
        { text: "estimation risk", options: { bold: true } },
        { text: " karena minim statistik.", options: { breakLine: true } },
        { text: "", options: { breakLine: true } },
        { text: "[Lihat Detail Simulasi Lampiran]", options: { fontSize: 14, color: "0563C1", underline: true, hyperlink: { slide: '32' } } }
    ], { x: 0.5, y: 1.2, w: "90%", h: 4, fontSize: 20, color: "333333", valign: "top" });
    slideEW.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slideEW.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Slide 5.2: Classical Markowitz (CM) ---
    let slideCM = pres.addSlide();
    slideCM.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slideCM.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slideCM.addText("1.2. Classical Markowitz (CM)", { x: 0.5, y: 0.5, w: "90%", fontSize: 28, bold: true, color: "003366" });
    slideCM.addText([
        { text: "Konsep Dasar:", options: { bold: true, breakLine: true } },
        { text: "Meminimalkan variansi untuk tingkat ", options: { bullet: true } },
        { text: "imbal hasil", options: { bold: true } },
        { text: " tertentu.", options: { breakLine: true } },
        { text: "Kelemahan:", options: { bold: true, breakLine: true } },
        { text: "Menderita ", options: { bullet: true } },
        { text: "ketidakstabilan numerik", options: { bold: true } },
        { text: " pada data yang ", options: {} },
        { text: "sangat berisik", options: { bold: true } },
        { text: " (noisy).", options: { breakLine: true } },
        { text: "Pondasi dasar sebagai ", options: { bullet: true } },
        { text: "teori tradisional", options: { bold: true } },
        { text: " dalam penelitian ini.", options: { breakLine: true } },
        { text: "", options: { breakLine: true } },
        { text: "[Lihat Detail Simulasi Lampiran]", options: { fontSize: 14, color: "0563C1", underline: true, hyperlink: { slide: '33' } } }
    ], { x: 0.5, y: 1.2, w: "90%", h: 4, fontSize: 20, color: "333333", valign: "top" });
    slideCM.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slideCM.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Slide 5.3: Graphical Lasso Markowitz (GM) ---
    let slideGM = pres.addSlide();
    slideGM.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slideGM.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slideGM.addText("2. Graphical Lasso Markowitz (GM)", { x: 0.5, y: 0.5, w: "90%", fontSize: 28, bold: true, color: "003366" });
    slideGM.addText([
        { text: "Konsep Dasar:", options: { bold: true, breakLine: true } },
        { text: "Menggunakan algoritma ", options: { bullet: true } },
        { text: "Lasso", options: { bold: true } },
        { text: " pada matriks presisi (invers kovarians) untuk memaksa korelasi yang tidak signifikan menjadi nol.", options: { breakLine: true } },
        { text: "Tujuan:", options: { bold: true, breakLine: true } },
        { text: "Menciptakan struktur ", options: { bullet: true } },
        { text: "'sparsity'", options: { bold: true } },
        { text: " (kerekatan) pada jaringan.", options: { breakLine: true } },
        { text: "Menangani tantangan data kripto yang sering terkorelasi secara ", options: { bullet: true } },
        { text: "palsu", options: { bold: true } },
        { text: " (spurious correlations).", options: { breakLine: true } },
        { text: "", options: { breakLine: true } },
        { text: "[Lihat Detail Simulasi Lampiran]", options: { fontSize: 14, color: "0563C1", underline: true, hyperlink: { slide: '34' } } }
    ], { x: 0.5, y: 1.2, w: "90%", h: 4, fontSize: 20, color: "333333", valign: "top" });
    slideGM.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slideGM.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Slide 5.4: Network Markowitz (NW) Statis ---
    let slideNWStatic = pres.addSlide();
    slideNWStatic.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slideNWStatic.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slideNWStatic.addText("3. Network Markowitz (NW) Statis", { x: 0.5, y: 0.5, w: "90%", fontSize: 28, bold: true, color: "003366" });
    slideNWStatic.addText([
        { text: "Konsep Dasar:", options: { bold: true, breakLine: true } },
        { text: "Model jaringan original (Giudici et al., 2020) yang menggabungkan ", options: { bullet: true } },
        { text: "filter RMT (Random Matrix Theory)", options: { bold: true } },
        { text: " dan ", options: {} },
        { text: "penalti sentralitas graf", options: { bold: true } },
        { text: ".", options: { breakLine: true } },
        { text: "Karakteristik:", options: { bold: true, breakLine: true } },
        { text: "Menggunakan parameter penalti (gamma) yang bersifat ", options: { bullet: true } },
        { text: "statis/tetap", options: { bold: true } },
        { text: " (hard-coded).", options: { breakLine: true } },
        { text: "Digunakan sebagai ", options: { bullet: true } },
        { text: "pembanding langsung", options: { bold: true } },
        { text: " untuk menguji efisiensi parameter hasil tuning.", options: { breakLine: true } },
        { text: "", options: { breakLine: true } },
        { text: "[Lihat Detail Simulasi Lampiran]", options: { fontSize: 14, color: "0563C1", underline: true, hyperlink: { slide: '36' } } }
    ], { x: 0.5, y: 1.2, w: "90%", h: 4, fontSize: 20, color: "333333", valign: "top" });
    slideNWStatic.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slideNWStatic.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Slide 5.5: Network Markowitz (NW) 2-Stage GS (Dua Kolom) ---
    let slideNWAdaptive = pres.addSlide();
    slideNWAdaptive.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slideNWAdaptive.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slideNWAdaptive.addText("4. Network Markowitz (NW) 2-Stage GS", { x: 0.5, y: 0.5, w: "90%", fontSize: 28, bold: true, color: "003366" });
    // Kolom Kiri: Mekanisme
    slideNWAdaptive.addText([
        { text: "2-Stage Grid Search (Coarse-to-Fine):", options: { bold: true, color: "003366", breakLine: true } },
        { text: "❑ Tahap 1 (Coarse): Mencari range optimal (W, γ) secara makro.", options: { breakLine: true } },
        { text: "❑ Tahap 2 (Fine): Zoom-in pencarian di sekitar titik terbaik Stage 1.", options: { breakLine: true } },
        { text: "❑ Tuning spesifik untuk 3 Target: VAR, SHARPE, dan RACHEV.", options: { breakLine: true } },
        { text: "❑ Menjamin parameter terbaik tanpa beban komputasi brute-force penuh.", options: { breakLine: true, fontSize: 18 } }
    ], { x: 0.5, y: 1.2, w: "45%", h: 4, fontSize: 20, color: "333333", valign: "top" });

    // Kolom Kanan: Keuntungan
    slideNWAdaptive.addText([
        { text: "Objektif & Resiliensi:", options: { bold: true, color: "003366", breakLine: true } },
        { text: "❑ Multi-Metric: Mencakup aspek risiko (VAR), efisiensi (Sharpe), dan ekor distribusi (Rachev).", options: { breakLine: true } },
        { text: "❑ Tuning Parameter: Menyesuaikan Window Size (stabilitas) & Gamma (penalti jaringan).", options: { breakLine: true } },
        { text: "❑ Konsistensi: Validasi melalui Fine-Search memastikan solusi bukan noise lokal.", options: { breakLine: true } },
        { text: "❑ Hasil: Strategi yang lebih tangguh di berbagai fase pasar (Bearish, Recovery, Stable).", options: { breakLine: true, color: "e67e22" } },
        { text: "", options: { breakLine: true } },
        { text: "[Lihat Hasil Comparison & Grid Search Lampiran]", options: { fontSize: 14, color: "0563C1", underline: true, hyperlink: { slide: '49' } } }
    ], { x: 5.2, y: 1.2, w: "45%", h: 4, fontSize: 20, color: "333333", valign: "top" });
    slideNWAdaptive.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slideNWAdaptive.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Slide 6: Matriks Evaluasi Performa ---
    let slide6 = pres.addSlide();
    slide6.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slide6.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slide6.addText("Matriks Evaluasi Performa (Multi-Metric)", { x: 0.5, y: 0.5, w: "90%", fontSize: 28, bold: true, color: "003366" });
    slide6.addText([
        { text: "1. Cumulative Profits & Losses (P&L):", options: { bold: true, breakLine: true, color: "003366", hyperlink: { slide: '56' } } },
        { text: "   Mengukur total keuntungan/kerugian akumulatif vs benchmark.", options: { breakLine: true } },

        { text: "2. Value at Risk (VaR 95%):", options: { bold: true, breakLine: true, color: "c0392b", hyperlink: { slide: '57' } } },
        { text: "   Fokus optimasi pada resiliensi risiko/batas bawah.", options: { breakLine: true } },

        { text: "3. Sharpe Ratio (SR):", options: { bold: true, breakLine: true, color: "27ae60", hyperlink: { slide: '58' } } },
        { text: "   Fokus optimasi pada efisiensi imbal hasil per unit risiko.", options: { breakLine: true } },

        { text: "4. Rachev Ratio (RR):", options: { bold: true, breakLine: true, color: "8e44ad", hyperlink: { slide: '59' } } },
        { text: "   Fokus optimasi pada perbandingan ekor distribusi (Fat-Tails).", options: { breakLine: true } },

        { text: "Justifikasi Multi-Objective:", options: { bold: true, breakLine: true, color: "d35400", fontSize: 14 } },
        { text: "Pasar kripto yang dinamis memerlukan tuning parameter yang berbeda tergantung apakah investor memprioritaskan keamanan (VaR), efisiensi (Sharpe), atau pemulihan (Rachev).", options: { italic: true, fontSize: 14 } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.0, fontSize: 17, color: "333333", valign: "top" });
    slide6.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slide6.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Slide 17: Analisis per Fase Pasar ---
    let slidePhase = pres.addSlide();
    slidePhase.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slidePhase.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slidePhase.addText("Evaluasi Berdasarkan Fase Pasar", { x: 0.5, y: 0.5, w: "90%", fontSize: 28, bold: true, color: "003366" });
    slidePhase.addText([
        { text: "Pengujian pada 3 Kondisi Berbeda:", options: { bold: true, color: "003366", breakLine: true } },
        { text: "1. Fase Bearish (Jan 2018 - Mar 2019):", options: { bold: true, color: "c0392b" } },
        { text: " Menguji ketahanan model saat pasar jatuh tajam.", options: { breakLine: true } },
        { text: "2. Fase Recovery (Apr 2019 - Jun 2019):", options: { bold: true, color: "27ae60" } },
        { text: " Menguji kecepatan adaptasi model saat tren berbalik positif.", options: { breakLine: true } },
        { text: "3. Fase Stable (Jul 2019 - Okt 2019):", options: { bold: true, color: "003366" } },
        { text: " Menguji konsistensi model dalam kondisi pasar mendatar.", options: { breakLine: true, breakLine: true } },

        { text: "Metodologi Panel Performance:", options: { bold: true, color: "d35400", breakLine: true } },
        { text: "✔ Panel A (Sharpe): Membandingkan efisiensi antar fase.", options: { breakLine: true } },
        { text: "✔ Panel B (Rachev): Membandingkan resiliensi ekor distribusi.", options: { breakLine: true } },
        { text: "✔ Kesimpulan: NW (2-Stage GS) menunjukkan stabilitas lebih tinggi dibanding benchmark statis di setiap fase.", options: { italic: true, breakLine: true } }
    ], { x: 0.5, y: 1.2, w: "90%", h: 4.5, fontSize: 18, valign: "top" });
    slidePhase.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slidePhase.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });


    // --- Slide 9: Lampiran - Analogi RMT ---
    let slide9 = pres.addSlide();
    slide9.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slide9.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slide9.addText([
        { text: "Lampiran: Analogi " },
        { text: "RMT (Random Matrix Theory)", options: { hyperlink: { slide: '16' } } },
        { text: " sebagai \"Noise-Canceling\"" }
    ], { x: 0.5, y: 0.5, w: "90%", fontSize: 28, bold: true, color: "003366" });
    slide9.addText([
        { text: "Pasar Kripto = Pesta yang Bising:", options: { bold: true, breakLine: true } },
        { text: "Banyak fluktuasi harga karena ", options: { bullet: true } },
        { text: "sentimen sesaat", options: { bold: true } },
        { text: " / kebetulan (noise).", options: { breakLine: true } },

        { text: "Sinyal Korelasi Asli = Suara yang Ingin Didengar:", options: { bold: true, breakLine: true } },
        { text: "Hubungan nyata antar-aset yang ", options: { bullet: true } },
        { text: "stabil dan berbobot", options: { bold: true } },
        { text: ".", options: { breakLine: true } },
        { text: "Random Matrix Theory (RMT) = Headphone Noise-Canceling:", options: { bold: true, breakLine: true } },
        { text: "Membedakan gelombang statistik acak dari ", options: { bullet: true } },
        { text: "pola suara asli", options: { bold: true } },
        { text: " menggunakan Distribusi MP (Marchenko-Pastur).", options: { breakLine: true } },
        { text: "Meredam spekulasi jangka pendek untuk mencegah ", options: { bullet: true } },
        { text: "estimation error", options: { italic: true, bold: true } },
        { text: ".", options: { breakLine: true } },
        { text: "Matriks tersisa adalah hubungan yang ", options: { bullet: true } },
        { text: "bersih dan terpercaya", options: { bold: true } },
        { text: "." }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5, fontSize: 22, color: "333333", valign: "top" });
    slide9.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slide9.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Slide 10: Lampiran - Signal vs Noise ---
    let slide10 = pres.addSlide();
    slide10.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slide10.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slide10.addText("Lampiran: Membedakan Hubungan Sejati (Signal) vs Noise", { x: 0.5, y: 0.5, w: "90%", fontSize: 26, bold: true, color: "003366" });
    slide10.addText([
        { text: "1. Mencari Nilai Eigen (Eigenvalues):", options: { bold: true, breakLine: true } },
        { text: "Mengekstrak angka yang mewakili ", options: { bullet: true } },
        { text: "kekuatan pola", options: { bold: true } },
        { text: " pergerakan bersama.", options: { breakLine: true } },

        { text: "2. Batas Noise (Marchenko-Pastur):", options: { bold: true, breakLine: true } },
        { text: "RMT (Random Matrix Theory) menghitung ", options: { bullet: true } },
        { text: "batas teoretis", options: { bold: true } },
        { text: " maksimum dari matriks acak.", options: { breakLine: true } },
        { text: "3. Uji Coba Signal vs Noise:", options: { bold: true, breakLine: true } },
        { text: "NOISE JALUR: Jika Eigenvalue < λ_max. Dianggap ", options: { bullet: true } },
        { text: "kebetulan acak", options: { bold: true } },
        { text: ".", options: { breakLine: true } },
        { text: "SIGNAL JALUR: Jika Eigenvalue > λ_max. Dianggap ", options: { bullet: true } },
        { text: "ikatan fundamental", options: { bold: true } },
        { text: ".", options: { breakLine: true } },
        { text: "4. Pembersihan & Rekonstruksi:", options: { bold: true, breakLine: true } },
        { text: "Hanya nilai signal yang dipertahankan untuk membangun ", options: { bullet: true } },
        { text: "matriks korelasi bersih", options: { bold: true } },
        { text: "." }
    ], { x: 0.5, y: 1.1, w: "90%", h: 4.5, fontSize: 18, color: "333333", valign: "top" });
    slide10.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slide10.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Slide 11: Lampiran - Menghitung Nilai Eigen ---
    let slide11 = pres.addSlide();
    slide11.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slide11.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slide11.addText("Lampiran: Bagaimana Menghitung Nilai Eigen?", { x: 0.5, y: 0.5, w: "90%", fontSize: 26, bold: true, color: "003366" });
    slide11.addText([
        { text: "1. Matriks Korelasi (C):", options: { bold: true, breakLine: true } },
        { text: "Membentuk tabel yang menjabarkan seluruh ", options: { bullet: true } },
        { text: "korelasi pergerakan", options: { bold: true } },
        { text: " harga antar sepasang koin.", options: { breakLine: true } },
        { text: "2. Konsep Persamaan Karakteristik:", options: { bold: true, breakLine: true } },
        { text: "Mencari besaran skalar ", options: { bullet: true } },
        { text: "λ (eigenvalue)", options: { italic: true, bold: true } },
        { text: " dan vektor arah yang memenuhi: ", options: {} },
        { text: "C × v = λ × v", options: { bold: true, color: "c0392b", breakLine: true } },
        { text: "3. Solusi Determinan:", options: { bold: true, breakLine: true } },
        { text: "Nilai λ adalah akar dari persamaan determinan: ", options: { bullet: true } },
        { text: "Det(C - λI) = 0", options: { bold: true, color: "c0392b", breakLine: true } },
        { text: "4. Arti dari Spektrum Hasil:", options: { bold: true, breakLine: true } },
        { text: "Nilai λ terbesar mewakili ", options: { bullet: true } },
        { text: "penggerak pasar", options: { bold: true } },
        { text: " utama (Market Factor).", options: { breakLine: true } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 18, color: "333333", valign: "top" });
    slide11.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slide11.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Slide 12: Lampiran - Contoh Praktek (Dummy Data) ---
    let slide12 = pres.addSlide();
    slide12.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slide12.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slide12.addText("Lampiran: Praktek Sederhana Menghitung Eigenvalue", { x: 0.5, y: 0.5, w: "90%", fontSize: 26, bold: true, color: "003366" });
    slide12.addText([
        { text: "Konteks Dummy: 2 Koin, BTC (Bitcoin) & ETH (Ethereum), korelasi = 0.5", options: { bold: true, breakLine: true } },
        { text: "1. Matriks Korelasi (C):", options: { bold: true, breakLine: true } },
        { text: "   C = [ 1.0  0.5 ]", options: { fontFace: "Courier New", breakLine: true } },
        { text: "       [ 0.5  1.0 ]", options: { fontFace: "Courier New", breakLine: true } },
        { text: "2. Persamaan: Det(C - λI) = 0", options: { bold: true, breakLine: true } },
        { text: "   (1 - λ)² - (0.5)² = 0", options: { breakLine: true } },
        { text: "   λ² - 2λ + 0.75 = 0  ", options: { breakLine: true } },
        { text: "   (λ - 1.5)(λ - 0.5) = 0", options: { breakLine: true } },
        { text: "3. Hasil Akar Nilai Eigen:", options: { bold: true, breakLine: true } },
        { text: "   • λ₁ = 1.5 : ", options: { bullet: true } },
        { text: "Market Factor", options: { bold: true, color: "27ae60" } },
        { text: " (Signal Kuat).", options: { breakLine: true } },
        { text: "   • λ₂ = 0.5 : ", options: { bullet: true } },
        { text: "Idiosyncratic Risk", options: { bold: true, color: "c0392b" } },
        { text: " (Noise).", options: { breakLine: true } },
        { text: "Kesimpulan Filtering:", options: { bold: true, breakLine: true } },
        { text: "   Jika RMT (Random Matrix Theory) mematok batas λ_max = 1.0, maka λ₂ dianggap ", options: { bullet: true } },
        { text: "Noise", options: { bold: true } },
        { text: " lalu dinolkan, sementara λ₁ dijaga sebagai sinyal sejati.", options: {} }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 16, color: "333333", valign: "top" });
    slide12.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slide12.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Slide 13: Lampiran - Bagaimana Menghitung Korelasi? ---
    let slide13 = pres.addSlide();
    slide13.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slide13.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slide13.addText("Lampiran: Bagaimana Menghitung Korelasi?", { x: 0.5, y: 0.5, w: "90%", fontSize: 26, bold: true, color: "003366" });
    slide13.addText([
        { text: "1. Data Historis (Returns):", options: { bold: true, breakLine: true } },
        { text: "Input berupa runtut waktu dari ", options: { bullet: true } },
        { text: "return harian", options: { bold: true } },
        { text: " aset kripto.", options: { breakLine: true } },
        { text: "2. Library & Metode Python:", options: { bold: true, breakLine: true } },
        { text: "Dihitung menggunakan ", options: { bullet: true } },
        { text: "Pandas (df.corr())", options: { fontFace: "Courier New", color: "c0392b", bold: true } },
        { text: " berbasis koefisien Pearson.", options: { breakLine: true } },
        { text: "3. Formula Pearson Correlation:", options: { bold: true, breakLine: true } },
        { text: "ρ(X,Y) = Cov(X,Y) / (σX × σY)", options: { bold: true, color: "27ae60", breakLine: true } },
        { text: "4. Output Matriks (N x N):", options: { bold: true, breakLine: true } },
        { text: "Nilai berkisar antara -1 hingga ", options: { bullet: true } },
        { text: "1 (searah)", options: { bold: true } },
        { text: ". Diagonal selalu 1.", options: { breakLine: true } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 16, color: "333333", valign: "top" });
    slide13.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slide13.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Slide 14: Lampiran - Apakah Nilai Eigen Statis? ---
    let slide14 = pres.addSlide();
    slide14.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slide14.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slide14.addText("Lampiran: Apakah Nilai Eigen Statis?", { x: 0.5, y: 0.5, w: "90%", fontSize: 26, bold: true, color: "003366" });
    slide14.addText([
        { text: "Apakah Nilai Eigen Sudah Ditentukan (Statis)?", options: { bold: true, breakLine: true, color: "c0392b" } },
        { text: "TIDAK. Nilai Eigen diekstrak langsung dari matriks korelasi ", options: { bullet: true } },
        { text: "saat rebalancing", options: { bold: true } },
        { text: ".", options: { breakLine: true } },
        { text: "Proses Dinamis:", options: { bold: true, breakLine: true } },
        { text: "Matriks korelasi dihitung dari data terbaru via ", options: { bullet: true } },
        { text: "rolling window", options: { bold: true } },
        { text: ".", options: { breakLine: true } },
        { text: "Batas filter Marchenko-Pastur ikut ", options: { bullet: true } },
        { text: "dihitung ulang", options: { bold: true } },
        { text: " mengikuti rasio data.", options: { breakLine: true } },
        { text: "Kesimpulan: Hasil tuning secara ", options: { bullet: true } },
        { text: "real-time", options: { bold: true, italic: true } },
        { text: " merespons perubahan rezim pasar dengan cepat.", options: {} }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 16, color: "333333", valign: "top" });
    slide14.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slide14.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Slide 15: Lampiran - Batas Noise Marchenko-Pastur ---
    let slide15 = pres.addSlide();
    slide15.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slide15.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slide15.addText("Lampiran: Menentukan Batas Noise (Marchenko-Pastur)", { x: 0.5, y: 0.5, w: "90%", fontSize: 26, bold: true, color: "003366" });
    slide15.addText([
        { text: "Apa itu Batas Marchenko-Pastur (MP)?", options: { bold: true, breakLine: true } },
        { text: "Prediksi bentuk distribusi dari matriks yang ", options: { bullet: true } },
        { text: "100% acak", options: { bold: true } },
        { text: " (noise).", options: { breakLine: true } },
        { text: "Menghitung Batas Atas Noise (λ_max):", options: { bold: true, breakLine: true } },
        { text: "λ_max = 1 + (1/Q) + 2√(1/Q)", options: { bold: true, color: "c0392b", breakLine: true } },
        { text: "Apa itu Rasio Q?", options: { bold: true, breakLine: true } },
        { text: "Q = T / N", options: { bold: true, color: "27ae60" } },
        { text: " (Baris Data / Aset).", options: { breakLine: true } },
        { text: "Mekanisme Filtering:", options: { bold: true, breakLine: true } },
        { text: "Eigenvalue < λ_max : ", options: { bullet: true } },
        { text: "Dihapus (Noise)", options: { bold: true } },
        { text: ".", options: { breakLine: true } },
        { text: "Eigenvalue > λ_max : ", options: { bullet: true } },
        { text: "Dipertahankan (Sinyal)", options: { bold: true } },
        { text: "." }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 16, color: "333333", valign: "top" });
    slide15.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slide15.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Slide 16: Lampiran - Analogi Minimum Spanning Tree (Bagian 1) ---
    let slide16 = pres.addSlide();
    slide16.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slide16.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slide16.addText("Lampiran: Analogi Minimum Spanning Tree (MST)", { x: 0.5, y: 0.5, w: "90%", fontSize: 26, bold: true, color: "003366" });
    slide16.addText([
        { text: "Membangun Jaringan Jalan Tol Antar Kota:", options: { bold: true, breakLine: true } },
        { text: "Hubungkan aset yang ", options: { bullet: true } },
        { text: "paling dekat", options: { bold: true } },
        { text: " (korelasi terkuat).", options: { breakLine: true } },
        { text: "Semua koin harus ", options: { bullet: true } },
        { text: "terhubung satu jaringan", options: { bold: true } },
        { text: ".", options: { breakLine: true } },
        { text: "Dilarang membuat ", options: { bullet: true } },
        { text: "jalan memutar", options: { bold: true } },
        { text: " (tanpa loop/redundansi).", options: { breakLine: true } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 18, color: "333333", valign: "top" });
    slide16.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slide16.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Slide 17: Lampiran - Analogi Minimum Spanning Tree (Bagian 2) ---
    let slide17 = pres.addSlide();
    slide17.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slide17.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slide17.addText("Lampiran: Mengapa Kita Membutuhkan MST?", { x: 0.5, y: 0.5, w: "90%", fontSize: 26, bold: true, color: "003366" });
    slide17.addText([
        { text: "Menemukan Titik Kemacetan (Hub Centrality):", options: { bold: true, color: "8e44ad", breakLine: true } },
        { text: "Di jaringan tol, kota besar menjadi pusat persimpangan. Di kripto, ini mewakili koin yang ", options: { bullet: true } },
        { text: "sangat sentral", options: { bold: true, italic: true } },
        { text: ".", options: { breakLine: true } },
        { text: "Jika terjadi kecelakaan di koin sentral, efeknya akan ", options: { bullet: true } },
        { text: "langsung menular", options: { bold: true } },
        { text: " ke seluruh jaringan.", options: { breakLine: true } },
        { text: "Solusi Network Markowitz:", options: { bold: true, breakLine: true } },
        { text: "Koin sentral akan diberi ", options: { bullet: true } },
        { text: "hukuman penalti", options: { bold: true } },
        { text: " agar portofolio tetap kokoh saat koin tersebut crash." }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 18, color: "333333", valign: "top" });
    slide17.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slide17.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Slide 18: Lampiran - Penalti (Gamma) Optimal ---
    let slide18 = pres.addSlide();
    slide18.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slide18.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slide18.addText("Lampiran: Berapa Nilai \"Penalti\" (Gamma) yang Optimal?", { x: 0.5, y: 0.5, w: "90%", fontSize: 26, bold: true, color: "003366" });
    slide18.addText([
        { text: "Apakah Ada Satu Angka Penalti yang Sempurna?", options: { bold: true, breakLine: true, color: "c0392b" } },
        { text: "TIDAK. Nilai gamma optimal berubah tergantung ", options: { bullet: true } },
        { text: "siklus fluktuasi", options: { bold: true } },
        { text: ".", options: { breakLine: true } },
        { text: "1. Fase Bull Market (Pasar Menguat):", options: { bold: true, breakLine: true, color: "27ae60" } },
        { text: "Penalti berat berisiko menghilangkan ", options: { bullet: true } },
        { text: "peluang untung", options: { bold: true } },
        { text: ". Gamma optimal cenderung rendah.", options: { breakLine: true } },
        { text: "2. Fase Bear Market / Crash (Pasar Jatuh):", options: { bold: true, breakLine: true, color: "c0392b" } },
        { text: "Keruntuhan koin sentral bisa memicu penularan (", options: { bullet: true } },
        { text: "Tail-Risk", options: { bold: true, italic: true } },
        { text: "). Gamma optimal cenderung tinggi.", options: { breakLine: true } },
        { text: "Kesimpulan:", options: { bold: true, breakLine: true, color: "8e44ad" } },
        { text: "Grid Search membiarkan komputer menyesuaikan gamma secara otomatis dengan ", options: { bullet: true } },
        { text: "data terbaru", options: { bold: true } },
        { text: "." }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 16, color: "333333", valign: "top" });
    slide18.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slide18.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Slide 18.1: Justifikasi Parameter Gamma (Giudici, 2020) ---
    let slideGammaJust = pres.addSlide();
    slideGammaJust.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slideGammaJust.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slideGammaJust.addText("Lampiran: Justifikasi Parameter Gamma (\u03b3)", { x: 0.5, y: 0.5, w: "90%", fontSize: 26, bold: true, color: "003366" });
    slideGammaJust.addText([
        { text: "Berdasarkan Teori Giudici et al. (2020):", options: { bold: true, breakLine: true, color: "27ae60" } },
        { text: "Parameter \u03b3 berfungsi sebagai pengukur aversi tingkat risiko sistemik investor.", options: { breakLine: true } },

        { text: "1. Batas Bawah (\u03b3 = 0):", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "   Model kembali ke Markowitz Tradisional. Investor sepenuhnya mengabaikan posisi aset dalam jaringan.", options: { breakLine: true } },

        { text: "2. Batas Atas (\u03b3 = 1):", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "   Investor memberikan bobot yang SETARA antara risiko individual (volatilitas) dan risiko sistemik (sentralitas).", options: { breakLine: true } },

        { text: "Signifikansi Nilai 1:", options: { bold: true, breakLine: true, color: "c0392b" } },
        { text: "   \u2713 Batas \"Aversi Risiko Tinggi\" yang masuk akal secara konseptual.", options: { breakLine: true } },
        { text: "   \u2713 Risiko penularan (contagion) dianggap sama pentingnya dengan fluktuasi harga aset.", options: { breakLine: true } },

        { text: "Optimalitas:", options: { bold: true, breakLine: true, color: "8e44ad" } },
        { text: "Menjaga model agar tidak hanya mengejar return, tapi juga memastikan portofolio tidak hancur saat hub utama pasar kolaps.", options: { italic: true } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 18, color: "333333", valign: "top" });
    slideGammaJust.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slideGammaJust.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Slide 19: Lampiran - Classical Markowitz (Bagian 1) ---
    let slide19 = pres.addSlide();
    slide19.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slide19.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slide19.addText("Lampiran: Apa itu Classical Markowitz (CM)? (1/2)", { x: 0.5, y: 0.5, w: "90%", fontSize: 26, bold: true, color: "003366" });
    slide19.addText([
        { text: "Modern Portfolio Theory (MPT) / Mean-Variance Optimization:", options: { bold: true, breakLine: true } },
        { text: "Merupakan teori klasik (ditemukan oleh Harry Markowitz tahun 1952) yang mencoba meramu komposisi/bobot aset dalam portofolio dengan tujuan matematika murni:", options: { bullet: true } },
        { text: "Memaksimalkan tingkat keuntungan (Return) pada tingkat risiko tertentu, ATAU", options: { bullet: true } },
        { text: "Meminimalkan risiko (Variance) pada tingkat keuntungan tertentu.", options: { bullet: true } },
        { text: "Asumsi Dasar Classical Markowitz:", options: { bold: true, breakLine: true } },
        { text: "Investor diasumsikan sepenuhnya rasional dan benci risiko (", options: { bullet: true } },
        { text: "Risk-averse", options: { italic: true, bold: true } },
        { text: ").", options: { breakLine: true } },
        { text: "Model ini sangat bergantung pada ", options: { bullet: true } },
        { text: "matriks kovarians historis", options: { bold: true } },
        { text: " sebagai pedoman utama memprediksi masa depan.", options: {} }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 18, color: "333333", valign: "top" });
    slide19.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slide19.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Slide 20: Lampiran - Classical Markowitz (Bagian 2) ---
    let slide20 = pres.addSlide();
    slide20.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slide20.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slide20.addText("Lampiran: Mengapa Classical Markowitz Kesulitan? (2/2)", { x: 0.5, y: 0.5, w: "90%", fontSize: 26, bold: true, color: "003366" });
    slide20.addText([
        { text: "Kelemahan Klasik di Pasar Kripto:", options: { bold: true, breakLine: true, color: "c0392b" } },
        { text: "Pasar kripto bersifat ", options: { bullet: true } },
        { text: "hyper-volatile", options: { bold: true, italic: true } },
        { text: " dengan korelasi ekor tebal.", options: { breakLine: true } },
        { text: "Model CM (Classical Markowitz) memakan mentah-mentah ", options: { bullet: true } },
        { text: "noise acak", options: { bold: true } },
        { text: " tanpa filter, berujung pada kegagalan prediksi.", options: { breakLine: true } },
        { text: "Evolusi → Network Markowitz:", options: { bold: true, breakLine: true, color: "27ae60" } },
        { text: "Mengembangkan CM (Classical Markowitz) dengan membersihkan noise menggunakan RMT (Random Matrix Theory) dan menghukum koin yang rawan menderita ", options: { bullet: true } },
        { text: "efek contagion", options: { bold: true } },
        { text: " menggunakan MST (Minimum Spanning Tree).", options: {} }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 18, color: "333333", valign: "top" });
    slide20.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slide20.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Slide 20.9: Lampiran - Detail Perhitungan Korelasi (Pearson) ---
    let slide20b = pres.addSlide();
    slide20b.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slide20b.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slide20b.addText("Lampiran: Detail Perhitungan Korelasi (Pearson)", { x: 0.5, y: 0.5, w: "90%", fontSize: 24, bold: true, color: "003366" });
    slide20b.addText([
        { text: "Bagaimana angka 0.50 didapatkan? (Dummy 3 Hari)", options: { bold: true, breakLine: true } },
        { text: "1. Data Return (X=BTC, Y=ETH):", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "   • Hari 1: X=2%, Y=2% | Hari 2: X=-2%, Y=0% | Hari 3: X=0%, Y=-2%", options: { breakLine: true } },
        { text: "   • Rata-rata (Mean): X̄ = 0, Ȳ = 0", options: { breakLine: true } },

        { text: "2. Tabel Deviasi & Perkalian:", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "   • Σ(x-x̄)(y-ȳ) = (2)(2) + (-2)(0) + (0)(-2) = ", options: {} },
        { text: "4", options: { bold: true, breakLine: true } },
        { text: "   • Σ(x-x̄)² = (2)² + (-2)² + (0)² = ", options: {} },
        { text: "8", options: { bold: true, breakLine: true } },
        { text: "   • Σ(y-ȳ)² = (2)² + (0)² + (-2)² = ", options: {} },
        { text: "8", options: { bold: true, breakLine: true } },

        { text: "3. Rumus Korelasi (ρ):", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "   ρ = Σ(dev_x * dev_y) / √(Σdev_x² * Σdev_y²)", options: { fontFace: "Courier New", color: "c0392b", breakLine: true } },
        { text: "   ρ = 4 / √(8 * 8) = 4 / 8 = ", options: { fontFace: "Courier New" } },
        { text: "0.50", options: { bold: true, color: "27ae60", fontFace: "Courier New" } },

        { text: "Kesimpulan:", options: { bold: true, breakLine: true, color: "c0392b" } },
        { text: "Angka ini menunjukkan arah pergerakan yang searah (positif) namun tidak identik, yang kemudian digunakan sebagai input matriks kovarians.", options: { italic: true } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 16, color: "333333", valign: "top" });
    slide20b.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slide20b.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Slide 20.8: Lampiran - Contoh Sederhana Equally Weighted (EW) ---
    let slideEWExample = pres.addSlide();
    slideEWExample.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slideEWExample.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slideEWExample.addText("Lampiran: Contoh Sederhana Equally Weighted (EW)", { x: 0.5, y: 0.5, w: "90%", fontSize: 24, bold: true, color: "003366" });
    slideEWExample.addText([
        { text: "Strategi 1/N: Alokasi Tanpa Rumit", options: { bold: true, breakLine: true } },
        { text: "1. Logika Dasar:", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "   Membagi seluruh modal secara merata tanpa melihat performa masa lalu.", options: { breakLine: true } },

        { text: "2. Simulasi (Dummy):", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "   • Total Modal: Rp 100.000.000", options: { breakLine: true } },
        { text: "   • Jumlah Aset (N): 5 (BTC, ETH, XRP, LTC, USDT)", options: { breakLine: true } },
        { text: "   • Bobot Tiap Aset: 1/5 = ", options: {} },
        { text: "20%", options: { bold: true, color: "27ae60" } },
        { text: " (Tetap)", options: { breakLine: true } },

        { text: "3. Sebaran Modal:", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "   • BTC: Rp 20jt | ETH: Rp 20jt | XRP: Rp 20jt | dst.", options: { breakLine: true } },

        { text: "4. Mengapa EW Masuk Baseline?", options: { bold: true, breakLine: true, color: "c0392b" } },
        { text: "   • Tidak menderita ", options: {} },
        { text: "error estimasi", options: { bold: true } },
        { text: " karena tidak menghitung korelasi.", options: { breakLine: true } },
        { text: "   • Benchmark yang sangat tangguh; model kompleks harus bisa mengalahkan EW untuk dianggap valid.", options: { italic: true } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 18, color: "333333", valign: "top" });
    slideEWExample.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slideEWExample.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });


    // --- Slide 21: Lampiran - Simulasi Sederhana Classical Markowitz (2 Aset) ---

    let slide21 = pres.addSlide();
    slide21.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slide21.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slide21.addText("Lampiran: Simulasi Sederhana Classical Markowitz (2 Aset)", { x: 0.5, y: 0.5, w: "90%", fontSize: 24, bold: true, color: "003366" });
    slide21.addText([
        { text: "Tujuan: Mencari bobot (w) untuk risiko terendah (Minimum Variance).", options: { bold: true, breakLine: true } },
        { text: "1. Data Input (Dummy):", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "   • Aset A (BTC): Volatilitas (σ₁) = 20% (0.20)", options: { breakLine: true } },
        { text: "   • Aset B (ETH): Volatilitas (σ₂) = 30% (0.30)", options: { breakLine: true } },
        { text: "   • Korelasi (ρ₁₂): 0.50 (Didapat dari return harian historis)", options: { breakLine: true } },
        { text: "   • Kovarians (σ₁₂): ", options: { bold: true } },
        { text: "ρ₁₂ × σ₁ × σ₂", options: { italic: true } },
        { text: " = 0.50 × 0.20 × 0.30 = ", options: {} },
        { text: "0.03", options: { bold: true, breakLine: true } },


        { text: "2. Rumus Bobot Aset A (w₁):", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "   w₁ = (σ₂² - Cov) / (σ₁² + σ₂² - 2·Cov)", options: { fontFace: "Courier New", color: "c0392b", bold: true, breakLine: true } },
        { text: "   w₁ = (0.09 - 0.03) / (0.04 + 0.09 - 0.06)", options: { fontFace: "Courier New", breakLine: true } },
        { text: "   w₁ = 0.06 / 0.07 ≈ ", options: { fontFace: "Courier New" } },
        { text: "0.85 (85%)", options: { bold: true, color: "27ae60", breakLine: true } },

        { text: "3. Hasil Alokasi:", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "   • Bobot BTC: 85%", options: { bullet: true } },
        { text: "   • Bobot ETH: 15%", options: { bullet: true } },

        { text: "Kesimpulan Klasik:", options: { bold: true, breakLine: true, color: "c0392b" } },
        { text: "Markowitz akan memilih aset dengan volatilitas lebih rendah (BTC) secara dominan. Namun, jika angka σ₁ dan σ₂ ini mengandung \"noise\", maka alokasi ini menjadi tidak optimal (Over-concentration).", options: { italic: true } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 16, color: "333333", valign: "top" });
    slide21.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slide21.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Slide 22: Lampiran - Contoh GLasso (1/2: Pembersihan Korelasi) ---
    let slide22 = pres.addSlide();
    slide22.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slide22.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slide22.addText("Lampiran: Contoh GLasso (1/2: Korelasi)", { x: 0.5, y: 0.5, w: "90%", fontSize: 24, bold: true, color: "003366" });
    slide22.addText([
        { text: "Tujuan: Membuang korelasi palsu (noise) untuk mendapatkan sinyal pasar murni.", options: { bold: true, breakLine: true } },
        { text: "1. Matriks Korelasi Sebelum GLasso (Dirty):", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "          BTC    ETH    DOGE", options: { fontFace: "Courier New", breakLine: true } },
        { text: "   BTC  [ 1.00   0.70   0.08 ]  <-- 0.08 (Noise)", options: { fontFace: "Courier New", breakLine: true } },
        { text: "   ETH  [ 0.70   1.00   0.05 ]  <-- 0.05 (Noise)", options: { fontFace: "Courier New", breakLine: true } },
        { text: "   DOGE [ 0.08   0.05   1.00 ]", options: { fontFace: "Courier New", breakLine: true } },

        { text: "2. Proses Penalti (λ):", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "   Algoritma GLasso menekan korelasi lemah menjadi nol.", options: { breakLine: true } },

        { text: "3. Matriks Hasil GLasso (Clean/Sparse):", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "          BTC    ETH    DOGE", options: { fontFace: "Courier New", color: "27ae60", breakLine: true } },
        { text: "   BTC  [ 1.00   0.65   0.00 ]", options: { fontFace: "Courier New", color: "27ae60", breakLine: true } },
        { text: "   ETH  [ 0.65   1.00   0.00 ]", options: { fontFace: "Courier New", color: "27ae60", breakLine: true } },
        { text: "   DOGE [ 0.00   0.00   1.00 ]", options: { fontFace: "Courier New", color: "27ae60", breakLine: true } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 16, color: "333333", valign: "top" });
    slide22.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slide22.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Slide 23: Lampiran - Contoh GLasso (2/2: Dampak Bobot) ---
    let slide23 = pres.addSlide();
    slide23.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slide23.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slide23.addText("Lampiran: Contoh GLasso (2/2: Dampak Bobot)", { x: 0.5, y: 0.5, w: "90%", fontSize: 24, bold: true, color: "003366" });
    slide23.addText([
        { text: "Bagaimana 'Pembersihan' mengubah alokasi modal?", options: { bold: true, breakLine: true } },
        { text: "4. Simulasi Perubahan Bobot (Weight Shift):", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "   • Alokasi Tanpa GLasso:", options: { bold: true, breakLine: true } },
        { text: "     BTC: 60%, ETH: 30%, ", options: {} },
        { text: "DOGE: 10%", options: { bold: true, color: "c0392b" } },
        { text: " (Tertipu korelasi palsu).", options: { breakLine: true } },

        { text: "   • Alokasi Dengan GLasso:", options: { bold: true, breakLine: true } },
        { text: "     BTC: 65%, ETH: 35%, ", options: {} },
        { text: "DOGE: 0%", options: { bold: true, color: "27ae60" } },
        { text: " (Fokus pada korelasi sejati).", options: { breakLine: true } },

        { text: "Kesimpulan & Manfaat:", options: { bold: true, breakLine: true, color: "27ae60" } },
        { text: "Pembersihan noise melalui GLasso memastikan modal tidak dialokasikan ke aset yang hanya terlihat menguntungkan secara statistik sesaat (spurious divergence), melainkan tetap pada struktur pasar yang kokoh.", options: { italic: true } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 18, color: "333333", valign: "top" });
    slide23.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slide23.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Slide 23.5: Lampiran - Contoh Sederhana Network Markowitz (NW) ---
    let slideNWExample = pres.addSlide();
    slideNWExample.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slideNWExample.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slideNWExample.addText("Lampiran: Contoh Sederhana Network Markowitz (NW)", { x: 0.5, y: 0.5, w: "90%", fontSize: 24, bold: true, color: "003366" });
    slideNWExample.addText([
        { text: "Tujuan: Mengurangi risiko sistemik dengan menghukum koin 'pusat' (hub).", options: { bold: true, breakLine: true } },
        { text: "1. Skenario Jaringan (MST):", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "   Bayangkan 3 Koin: BTC, ETH, dan ADA.", options: { breakLine: true } },
        { text: "   • BTC terhubung ke ETH dan ADA (BTC adalah ", options: {} },
        { text: "Hub utama", options: { bold: true } },
        { text: ").", options: { breakLine: true } },
        { text: "   • Centrality (EC): BTC=0.8, ETH=0.4, ADA=0.4.", options: { breakLine: true } },

        { text: "2. Perbandingan Alokasi:", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "   • Classical Markowitz: Melirik BTC karena paling stabil, bobot bisa ", options: {} },
        { text: "70%", options: { bold: true, color: "c0392b" } },
        { text: ".", options: { breakLine: true } },
        { text: "   • Network Markowitz (γ=1.0): Mendeteksi BTC sebagai titik bahaya penularan. Bobot BTC dipangkas menjadi ", options: {} },
        { text: "45%", options: { bold: true, color: "27ae60" } },
        { text: ".", options: { breakLine: true } },

        { text: "3. Mekanisme 'Hukuman' (Penalty):", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "   Risiko Baru = Risiko Harga + (γ × Centrality)", options: { fontFace: "Courier New", color: "8e44ad", bold: true, breakLine: true } },

        { text: "Hasil Akhir:", options: { bold: true, breakLine: true, color: "27ae60" } },
        { text: "Modal dialihkan ke ETH/ADA yang lebih 'pinggiran' (peripheral). Portofolio tidak hancur total jika 'Hub' (BTC) mengalami crash ekstrem.", options: { italic: true } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 16, color: "333333", valign: "top" });
    slideNWExample.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slideNWExample.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Slide 23.6: Lampiran - Detail Kalkulasi Penentuan Bobot NW ---
    let slideNWDetail = pres.addSlide();
    slideNWDetail.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slideNWDetail.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slideNWDetail.addText("Lampiran: Detail Kalkulasi Penentuan Bobot NW", { x: 0.5, y: 0.5, w: "90%", fontSize: 24, bold: true, color: "003366" });
    slideNWDetail.addText([
        { text: "Bagaimana 'Penalti' mengubah angka bobot secara konkret?", options: { bold: true, breakLine: true } },
        { text: "1. Data Aset (Dummy):", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "   • Aset A (Hub): Volatilitas = 20%, Centrality (EC) = 0.8", options: { breakLine: true } },
        { text: "   • Aset B (Peripheral): Volatilitas = 30%, Centrality (EC) = 0.2", options: { breakLine: true } },
        { text: "   • Parameter Gamma (γ) = 0.5", options: { breakLine: true } },

        { text: "2. Perhitungan Risiko Penyesuaian (Adjusted Risk):", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "   • Risiko A = 0.20 + (0.5 × 0.8) = ", options: {} },
        { text: "0.60 (Meningkat)", options: { bold: true, color: "c0392b", breakLine: true } },
        { text: "   • Risiko B = 0.30 + (0.5 × 0.2) = ", options: {} },
        { text: "0.40 (Relatif Stabil)", options: { bold: true, color: "27ae60", breakLine: true } },

        { text: "3. Pergeseran Bobot (Weight Shift):", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "   • Bobot A (Kuno): ", options: {} },
        { text: "60% → 40% (NW)", options: { bold: true, color: "c0392b", breakLine: true } },
        { text: "   • Bobot B (Kuno): ", options: {} },
        { text: "40% → 60% (NW)", options: { bold: true, color: "27ae60", breakLine: true } },

        { text: "Inti Logika NW:", options: { bold: true, breakLine: true, color: "8e44ad" } },
        { text: "Aset A dihukum bukan karena harganya tidak stabil, tapi karena ia adalah 'pusat kemacetan' risiko. Model secara matematis memindahkan modal ke Aset B untuk proteksi sistemik.", options: { italic: true } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 16, color: "333333", valign: "top" });
    slideNWDetail.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slideNWDetail.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Slide 23.7: Lampiran - Rumus & Penjabaran Bobot (60% vs 40%) ---
    let slideNWMath = pres.addSlide();
    slideNWMath.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slideNWMath.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slideNWMath.addText("Lampiran: Rumus & Penjabaran Bobot (60% vs 40%)", { x: 0.5, y: 0.5, w: "90%", fontSize: 24, bold: true, color: "003366" });
    slideNWMath.addText([
        { text: "Menggunakan Metode Alokasi Volatilitas Terbalik (Inverse Volatility):", options: { bold: true, breakLine: true } },
        { text: "Rumus: Bobot (w) = (1 / Risiko) / Total (1 / Risiko)", options: { italic: true, breakLine: true, color: "003366" } },

        { text: "1. Perhitungan Skenario Klasik (Tanpa Penalti):", options: { bold: true, breakLine: true, color: "2c3e50" } },
        { text: "   • Aset A (Risiko 0.20) → 1 / 0.20 = ", options: {} },
        { text: "5.00", options: { bold: true } },
        { text: " (Daya Tarik)", options: { breakLine: true } },
        { text: "   • Aset B (Risiko 0.30) → 1 / 0.30 = ", options: {} },
        { text: "3.33", options: { bold: true } },
        { text: " (Daya Tarik)", options: { breakLine: true } },
        { text: "   • Bobot A = 5.00 / (5.00 + 3.33) = 5.00 / 8.33 ≈ ", options: {} },
        { text: "60%", options: { bold: true, color: "c0392b", breakLine: true } },

        { text: "2. Perhitungan Skenario Network (Gamma 0.5):", options: { bold: true, breakLine: true, color: "2c3e50" } },
        { text: "   • Aset A (Risiko 0.60) → 1 / 0.60 = ", options: {} },
        { text: "1.67", options: { bold: true } },
        { text: " (Daya Tarik Turun)", options: { breakLine: true } },
        { text: "   • Aset B (Risiko 0.40) → 1 / 0.40 = ", options: {} },
        { text: "2.50", options: { bold: true } },
        { text: " (Daya Tarik Naik)", options: { breakLine: true } },
        { text: "   • Bobot A = 1.67 / (1.67 + 2.50) = 1.67 / 4.17 ≈ ", options: {} },
        { text: "40%", options: { bold: true, color: "27ae60", breakLine: true } },

        { text: "Kesimpulan Matematis:", options: { bold: true, breakLine: true, color: "8e44ad" } },
        { text: "Model ini secara adil memberikan porsi lebih besar pada aset yang memiliki skor gabungan 'Risiko + Penalti' yang paling kecil.", options: { italic: true } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 16, color: "333333", valign: "top" });
    slideNWMath.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slideNWMath.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Slide 23.8: Lampiran - NW vs CM: Apakah Tetap Memakai Korelasi? ---
    let slideNWCorr = pres.addSlide();
    slideNWCorr.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slideNWCorr.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slideNWCorr.addText("Lampiran: NW vs CM - Hubungan dengan Korelasi", { x: 0.5, y: 0.5, w: "90%", fontSize: 24, bold: true, color: "003366" });
    slideNWCorr.addText([
        { text: "Pertanyaan Penting: Apakah NW mengesampingkan korelasi?", options: { bold: true, breakLine: true, color: "c0392b" } },
        { text: "Jawaban: TIDAK. NW justru menggunakan korelasi secara lebih cerdas.", options: { bold: true, breakLine: true, color: "27ae60" } },

        { text: "Perbandingan Alur Kerja:", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "❑ Classical Markowitz (CM):", options: { bold: true, breakLine: true } },
        { text: "   Korelasi Mentah → Matriks Kovarians → Bobot Portofolio.", options: { breakLine: true } },

        { text: "❑ Network Markowitz (NW):", options: { bold: true, breakLine: true } },
        { text: "   1. Korelasi Mentah → ", options: {} },
        { text: "Saring Noise (RMT)", options: { bold: true, color: "8e44ad" } },
        { text: " (Membersihkan korelasi palsu).", options: { breakLine: true } },
        { text: "   2. Korelasi Bersih → ", options: {} },
        { text: "Bangun MST (Network)", options: { bold: true, color: "8e44ad" } },
        { text: " (Melihat peta hubungan antar koin).", options: { breakLine: true } },
        { text: "   3. Struktur Network → ", options: {} },
        { text: "Penalty Centrality", options: { bold: true, color: "8e44ad" } },
        { text: " (Menimbang risiko penularan).", options: { breakLine: true } },
        { text: "   4. Gabungan → Optimasi Markowitz Akhir.", options: { breakLine: true } },

        { text: "Kesimpulan:", options: { bold: true, breakLine: true, color: "2c3e50" } },
        { text: "NW tidak membuang korelasi; ia 'mengolah' korelasi menjadi peta jaringan untuk mendeteksi risiko sistemik yang tidak terlihat oleh CM biasa.", options: { italic: true } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 18, color: "333333", valign: "top" });
    slideNWCorr.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slideNWCorr.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Slide 23.9: Lampiran - Contoh Detail Filter RMT (Pembersihan Noise) ---
    let slideRMTExample = pres.addSlide();
    slideRMTExample.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slideRMTExample.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slideRMTExample.addText("Lampiran: Contoh Detail Filter RMT (Pembersihan Noise)", { x: 0.5, y: 0.5, w: "90%", fontSize: 24, bold: true, color: "003366" });
    slideRMTExample.addText([
        { text: "Analogi: Menghilangkan suara 'kresek' radio agar lagu terdengar jernih.", options: { bold: true, breakLine: true } },

        { text: "1. Skenario Korelasi Mentah (Dirty Matrix):", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "   • BTC vs ETH   : 0.75 (Ikatan Fundamental)", options: { breakLine: true } },
        { text: "   • BTC vs DOGE  : 0.12 (Kebetulan Spekulatif/Noise)", options: { breakLine: true } },
        { text: "   • ETH vs PEPE  : 0.08 (Kebetulan Spekulatif/Noise)", options: { breakLine: true } },

        { text: "2. Uji Batas RMT (Marchenko-Pastur Distribution):", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "   Rumus RMT menghitung 'Ambang Batas Keacakan' (Threshold).", options: { breakLine: true } },
        { text: "   Misalkan hasil hitung batas noise (λ_max) = ", options: {} },
        { text: "0.15", options: { bold: true, color: "c0392b" } },
        { text: ".", options: { breakLine: true } },

        { text: "3. Proses Pembersihan (Filtering):", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "   • 0.75 > 0.15  → ", options: {} },
        { text: "DIJAGA", options: { bold: true, color: "27ae60" } },
        { text: " (Ini adalah Sinyal Pasar).", options: { breakLine: true } },
        { text: "   • 0.12 < 0.15  → ", options: {} },
        { text: "DIBUANG (JADI 0)", options: { bold: true, color: "c0392b" } },
        { text: " (Ini adalah Noise).", options: { breakLine: true } },
        { text: "   • 0.08 < 0.15  → ", options: {} },
        { text: "DIBUANG (JADI 0)", options: { bold: true, color: "c0392b" } },
        { text: " (Ini adalah Noise).", options: { breakLine: true } },

        { text: "Kesimpulan Akhir:", options: { bold: true, breakLine: true, color: "8e44ad" } },
        { text: "Dengan RMT, model hanya akan membangun portofolio berdasarkan 'Gema Fundamental' aset, bukan berdasarkan 'Kebetulan Statistik' yang sering menjebak investor di pasar kripto.", options: { italic: true } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 16, color: "333333", valign: "top" });
    slideRMTExample.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slideRMTExample.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Slide 23.10: Lampiran - Cara Menghitung Batas Noise (λ_max) ---
    let slideLambdaMax = pres.addSlide();
    slideLambdaMax.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slideLambdaMax.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slideLambdaMax.addText("Lampiran: Cara Menghitung Batas Noise (λ_max)", { x: 0.5, y: 0.5, w: "90%", fontSize: 24, bold: true, color: "003366" });
    slideLambdaMax.addText([
        { text: "Batas ini ditentukan menggunakan Distribusi Marchenko-Pastur:", options: { bold: true, breakLine: true } },
        { text: "Rumus: λ_max = σ² × (1 + √(1/Q))²", options: { bold: true, color: "c0392b", breakLine: true, fontFace: "Courier New" } },

        { text: "Komponen Rasio Q:", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "   • Q = T / N", options: { fontFace: "Courier New", breakLine: true } },
        { text: "   • T = Jumlah hari (Observasi)", options: { breakLine: true } },
        { text: "   • N = Jumlah koin (Aset)", options: { breakLine: true } },

        { text: "Contoh Simulasi:", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "   Kita punya data 10 koin (N=10) selama 1000 hari (T=1000).", options: { breakLine: true } },
        { text: "   1. Hitung Q: 1000 / 10 = ", options: {} },
        { text: "100", options: { bold: true, breakLine: true } },
        { text: "   2. Hitung Akar: √(1/100) = ", options: {} },
        { text: "0.1", options: { bold: true, breakLine: true } },
        { text: "   3. Hitung λ_max: (1 + 0.1)² = 1.1² = ", options: {} },
        { text: "1.21", options: { bold: true, color: "27ae60", breakLine: true } },

        { text: "Penerapan:", options: { bold: true, breakLine: true, color: "8e44ad" } },
        { text: "Setiap Nilai Eigen dari matriks korelasi yang nilainya di bawah 1.21 akan dianggap sebagai noise dan dibersihkan dari model.", options: { italic: true } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 16, color: "333333", valign: "top" });
    slideLambdaMax.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slideLambdaMax.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Slide 23.11: Lampiran - Contoh Membangun MST (Network Mapping) ---
    let slideMSTBuild = pres.addSlide();
    slideMSTBuild.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slideMSTBuild.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slideMSTBuild.addText("Lampiran: Contoh Membangun MST (Network Mapping)", { x: 0.5, y: 0.5, w: "90%", fontSize: 24, bold: true, color: "003366" });
    slideMSTBuild.addText([
        { text: "Bagaimana korelasi berubah menjadi peta jaringan?", options: { bold: true, breakLine: true } },

        { text: "1. Konversi Korelasi ke Jarak (Metric Distance):", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "   Semakin tinggi korelasi → Semakin dekat jaraknya (riferensi: Mantegna, 1999).", options: { breakLine: true } },
        { text: "   • BTC - ETH (Korelasi 0.90) → Jarak: ", options: {} },
        { text: "0.45 (Sangat Dekat)", options: { bold: true, color: "27ae60", breakLine: true } },
        { text: "   • BTC - LTC (Korelasi 0.85) → Jarak: ", options: {} },
        { text: "0.55 (Dekat)", options: { bold: true, color: "27ae60", breakLine: true } },
        { text: "   • ETH - XRP (Korelasi 0.40) → Jarak: ", options: {} },
        { text: "1.10 (Jauh)", options: { bold: true, color: "c0392b", breakLine: true } },

        { text: "2. Menghubungkan Titik (Kruskal's Algorithm):", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "   • Langkah 1: Hubungkan BTC & ETH (Jalur Terkuat).", options: { breakLine: true } },
        { text: "   • Langkah 2: Hubungkan BTC & LTC (Jalur Kedua).", options: { breakLine: true } },
        { text: "   • Langkah 3: Hubungkan ETH & XRP (Jalur Ketiga).", options: { breakLine: true } },

        { text: "Ciri Khas MST:", options: { bold: true, breakLine: true, color: "8e44ad" } },
        { text: "Hanya menyisakan (N-1) koneksi terkuat dan dilarang membentuk loop. Di sini terlihat BTC menjadi 'Hub' karena ia yang menghubungkan banyak koin.", options: { italic: true } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 16, color: "333333", valign: "top" });
    slideMSTBuild.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slideMSTBuild.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Slide 23.12: Lampiran - Rumus Konversi Korelasi ke Jarak ---
    let slideDistFormula = pres.addSlide();
    slideDistFormula.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slideDistFormula.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slideDistFormula.addText("Lampiran: Rumus Konversi Korelasi ke Jarak", { x: 0.5, y: 0.5, w: "90%", fontSize: 24, bold: true, color: "003366" });
    slideDistFormula.addText([
        { text: "Untuk membangun jaringan, korelasi harus diubah menjadi skor jarak (Metric Space).", options: { bold: true, breakLine: true } },
        { text: "Rumus (Mantegna, 1999): d_ij = √(2(1 - ρ_ij))", options: { bold: true, color: "c0392b", breakLine: true, fontFace: "Courier New" } },

        { text: "Simulasi Tabel Jarak:", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "   • Korelasi 1.00 (Identik) → d = √(2(1-1)) = ", options: {} },
        { text: "0", options: { bold: true, color: "27ae60", breakLine: true } },

        { text: "   • Korelasi 0.50 (Kuat) → d = √(2(1-0.5)) = √(1) = ", options: {} },
        { text: "1.00", options: { bold: true, color: "2c3e50", breakLine: true } },

        { text: "   • Korelasi 0.00 (Acak) → d = √(2(1-0)) = √(2) ≈ ", options: {} },
        { text: "1.41", options: { bold: true, color: "c0392b", breakLine: true } },

        { text: "   • Korelasi -1.00 (Berlawanan) → d = √(2(1-(-1))) = √(4) = ", options: {} },
        { text: "2.00", options: { bold: true, color: "c0392b", breakLine: true } },

        { text: "Logika Utama:", options: { bold: true, breakLine: true, color: "8e44ad" } },
        { text: "1. Hubungan searah → Jarak Pendek (Aset sering bergerak bersama).", options: { bullet: true } },
        { text: "2. Hubungan berlawanan → Jarak Jauh (Aset saling menjauh).", options: { bullet: true } },
        { text: "MST hanya mengambil jalur-jalur dengan Jarak (d) paling kecil agar efisien.", options: { italic: true } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 18, color: "333333", valign: "top" });
    slideDistFormula.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slideDistFormula.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Slide 23.13: Lampiran - Apakah Jaringan Selalu Berbentuk Star? ---
    let slideTopology = pres.addSlide();
    slideTopology.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slideTopology.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slideTopology.addText("Lampiran: Apakah Jaringan Selalu Berbentuk Star?", { x: 0.5, y: 0.5, w: "90%", fontSize: 24, bold: true, color: "003366" });
    slideTopology.addText([
        { text: "Jawaban: TIDAK. Topologi jaringan bersifat dinamis mengikuti rezim pasar.", options: { bold: true, breakLine: true, color: "c0392b" } },

        { text: "1. Tipe Star (Sentralistik):", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "   • Muncul saat pasar panik atau crash.", options: { breakLine: true } },
        { text: "   • Semua aset mengekor pada satu koin dominan (Hub).", options: { breakLine: true } },
        { text: "   • Penalti NW akan sangat berat pada koin pusat tersebut.", options: { breakLine: true } },

        { text: "2. Tipe Terdistribusi (Cluster):", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "   • Muncul saat pasar tenang atau aset memiliki 'narasi' berbeda.", options: { breakLine: true } },
        { text: "   • Hubungan terbagi ke beberapa kelompok (misal: DeFi Group, Stablecoin Group).", options: { breakLine: true } },
        { text: "   • Penalti NW akan lebih menyebar dan diversifikasi lebih alami.", options: { breakLine: true } },

        { text: "Dinamika dalam NW:", options: { bold: true, breakLine: true, color: "8e44ad" } },
        { text: "Kekuatan Network Markowitz adalah kemampuannya mendeteksi transisi bentuk ini secara real-time melalui data historis.", options: { italic: true } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 18, color: "333333", valign: "top" });
    slideTopology.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slideTopology.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Slide 23.14: Lampiran - Berapa Banyak Perhitungan Korelasi? ---
    let slideCorrCount = pres.addSlide();
    slideCorrCount.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slideCorrCount.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slideCorrCount.addText("Lampiran: Berapa Banyak Perhitungan Korelasi?", { x: 0.5, y: 0.5, w: "90%", fontSize: 24, bold: true, color: "003366" });
    slideCorrCount.addText([
        { text: "Untuk membangun satu matriks utuh, setiap aset harus dipasangkan satu sama lain.", options: { bold: true, breakLine: true } },
        { text: "Rumus Kombinasi: C(N,2) = (N × (N - 1)) / 2", options: { bold: true, color: "c0392b", breakLine: true, fontFace: "Courier New" } },

        { text: "Simulasi untuk 10 Aset (N=10):", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "   Jumlah Pasangan Unik = (10 × 9) / 2 = ", options: {} },
        { text: "45 Perhitungan", options: { bold: true, color: "27ae60", breakLine: true } },

        { text: "Detail Penjabaran:", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "   • Koin ke-1 (BTC) berpasangan dengan 9 koin lainnya.", options: { breakLine: true } },
        { text: "   • Koin ke-2 (ETH) berpasangan dengan 8 koin sisa.", options: { breakLine: true } },
        { text: "   • Koin ke-3 (XRP) berpasangan dengan 7 koin sisa.", options: { breakLine: true } },
        { text: "   • ... Seterusnya.", options: { breakLine: true } },
        { text: "   • Total: 9 + 8 + 7 + 6 + 5 + 4 + 3 + 2 + 1 = ", options: {} },
        { text: "45", options: { bold: true } },

        { text: "Fakta Matriks 10x10:", options: { bold: true, breakLine: true, color: "8e44ad" } },
        { text: "Meskipun ada 100 kotak di matriks, komputer hanya perlu menghitung 45 angka unik karena korelasi bersifat cermin (A-B sama dengan B-A) dan tengahnya selalu 1 (A-A).", options: { italic: true } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 16, color: "333333", valign: "top" });
    slideCorrCount.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slideCorrCount.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Slide 23.15: Lampiran - Alur Transformasi Matriks ---
    let slidePipeline = pres.addSlide();
    slidePipeline.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slidePipeline.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slidePipeline.addText("Lampiran: Alur Transformasi Matriks", { x: 0.5, y: 0.5, w: "90%", fontSize: 24, bold: true, color: "003366" });
    slidePipeline.addText([
        { text: "Proses ini memastikan data statistik mentah bisa divisualisasikan menjadi peta risiko.", options: { bold: true, breakLine: true } },

        { text: "Langkah 1: Matriks Korelasi (Pearson)", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "   Berisi angka -1 s/d 1. Menunjukkan 'Kemiripan' gerak aset.", options: { breakLine: true } },

        { text: "Langkah 2: Matriks Jarak (Metric Distance)", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "   Angka korelasi diubah menjadi angka 0 s/d 2 (Jarak). Semakin mirip aset, semakin 'nempel' (jarak mendekati 0).", options: { breakLine: true } },

        { text: "Langkah 3: Bangun Jaringan (MST Algorithm)", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "   Komputer membaca Matriks Jarak sebagai peta jalan tol, lalu memilih rute terpendek untuk menghubungkan seluruh aset.", options: { breakLine: true } },

        { text: "Hasil Akhir:", options: { bold: true, breakLine: true, color: "27ae60" } },
        { text: "Jaringan (Network) yang siap digunakan untuk menghitung skor penalti penularan risiko sistemik.", options: { italic: true } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 18, color: "333333", valign: "top" });
    slidePipeline.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slidePipeline.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Slide 23.16: Lampiran - Mekanisme Penalty Centrality ---
    let slidePenalty = pres.addSlide();
    slidePenalty.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slidePenalty.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slidePenalty.addText("Lampiran: Mekanisme Penalty Centrality (Risiko Penularan)", { x: 0.5, y: 0.5, w: "90%", fontSize: 24, bold: true, color: "003366" });
    slidePenalty.addText([
        { text: "Kenapa koin yang 'populer' di jaringan justru diberi hukuman (penalti)?", options: { bold: true, breakLine: true } },

        { text: "Analogi Bandara Transit (Hub):", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "   Jika Bandara transit utama (misal: Singapura) terkena badai, maka seluruh penerbangan di Asia Tenggara akan ikut kacau. Singapura adalah 'Hub' dengan Centrality tinggi.", options: { breakLine: true } },

        { text: "Skenario di Kripto (Dominasi BTC):", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "   • BTC terhubung ke semua koin lain dalam MST.", options: { breakLine: true } },
        { text: "   • Skor Centrality BTC = ", options: {} },
        { text: "1.00", options: { bold: true, color: "c0392b" } },
        { text: " (Maksimal).", options: { breakLine: true } },

        { text: "Cara Model NW Menghitung Risiko:", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "   Risiko Baru = Risiko Harga + (γ × ", options: {} },
        { text: "1.00", options: { bold: true, color: "c0392b" } },
        { text: ")", options: { breakLine: true } },

        { text: "Tujuan Penalti:", options: { bold: true, breakLine: true, color: "8e44ad" } },
        { text: "Agar portofolio tidak menaruh terlalu banyak modal pada koin yang bisa memicu 'Efek Domino'. Jika BTC rontok, penalti memastikan kita sudah punya cadangan di koin-koin 'pinggiran' (peripheral) yang lebih mandiri.", options: { italic: true } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 18, color: "333333", valign: "top" });
    slidePenalty.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slidePenalty.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Slide 23.17: Lampiran - Cara Menghitung Centrality (Degree Centrality) ---
    let slideCalcCentrality = pres.addSlide();
    slideCalcCentrality.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slideCalcCentrality.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slideCalcCentrality.addText("Lampiran: Cara Menghitung Centrality (Degree Centrality)", { x: 0.5, y: 0.5, w: "90%", fontSize: 24, bold: true, color: "003366" });
    slideCalcCentrality.addText([
        { text: "Metode Degree Centrality mengukur seberapa banyak 'tangan' yang dimiliki sebuah aset untuk memegang aset lain.", options: { bold: true, breakLine: true } },

        { text: "Rumus: Centrality = Jumlah Koneksi / (N - 1)", options: { bold: true, color: "c0392b", breakLine: true, fontFace: "Courier New" } },

        { text: "Simulasi untuk 5 Aset (BTC, ETH, XRP, LTC, USDT):", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "   • Aset Utama (N) = 5. Pembagi (N-1) = 4.", options: { breakLine: true } },

        { text: "1. Aset BTC (Terkoneksi ke ETH, XRP, LTC):", options: { breakLine: true } },
        { text: "   3 Koneksi / 4 = ", options: {} },
        { text: "0.75 (Centrality Tinggi)", options: { bold: true, color: "c0392b", breakLine: true } },

        { text: "2. Aset LTC (Terkoneksi ke BTC & USDT):", options: { breakLine: true } },
        { text: "   2 Koneksi / 4 = ", options: {} },
        { text: "0.50 (Centrality Sedang)", options: { bold: true, color: "f39c12", breakLine: true } },

        { text: "3. Aset ETH, XRP, USDT (Hanya 1 Koneksi):", options: { breakLine: true } },
        { text: "   1 Koneksi / 4 = ", options: {} },
        { text: "0.25 (Centrality Rendah)", options: { bold: true, color: "27ae60", breakLine: true } },

        { text: "Dampak Penalti NW:", options: { bold: true, breakLine: true, color: "8e44ad" } },
        { text: "Karena BTC memiliki skor 0.75, ia akan menerima penalti 3x lebih besar daripada ETH (0.25). Ini menjaga portofolio tetap terdiversifikasi dari pusat jaringan.", options: { italic: true } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 16, color: "333333", valign: "top" });
    slideCalcCentrality.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slideCalcCentrality.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Slide 23.18: Lampiran - Contoh Sederhana Rolling Window Grid Search ---
    let slideGridSearchEx = pres.addSlide();
    slideGridSearchEx.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slideGridSearchEx.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slideGridSearchEx.addText("Lampiran: Mekanisme 2-Stage Grid Search (Coarse-to-Fine)", { x: 0.5, y: 0.5, w: "90%", fontSize: 24, bold: true, color: "003366" });
    slideGridSearchEx.addText([
        { text: "Tujuan: Menemukan kombinasi (Window, γ) paling presisi dengan efisiensi tinggi.", options: { bold: true, breakLine: true } },
        { text: "1. Stage 1: Coarse Search (Pemindaian Makro)", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "   ❑ Mencari kandidat di seluruh ruang parameter dengan langkah (step) lebar.", options: { breakLine: true } },
        { text: "   ❑ Contoh: Window dipantau setiap kelipatan 5 hari, dan Gamma setiap 0.2 unit.", options: { breakLine: true } },
        { text: "   ❑ Output: Menentukan 'zona potensial' terbaik secara cepat.", options: { breakLine: true } },
        { text: "2. Stage 2: Fine Search (Zoom-in Mikro)", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "   ❑ Melakukan pencarian mendalam hanya di sekitar koordinat terbaik Stage 1.", options: { breakLine: true } },
        { text: "   ❑ Contoh: Jika koordinat terbaik Stage 1 adalah (W:40, γ:0.6), Stage 2 mencari di rentang presisi (W: 36-44) dan (γ: 0.55-0.65).", options: { breakLine: true } },
        { text: "   ❑ Output: Mendapatkan titik absolut paling optimal.", options: { breakLine: true } },
        { text: "Keunggulan Eksperimen:", options: { bold: true, breakLine: true, color: "8e44ad" } },
        { text: "✔ Akurasi: Jauh lebih detail dibanding grid search tunggal konvensional.", options: { breakLine: true } },
        { text: "✔ Efisiensi: Mengurangi total percobaan komputasi hingga >60% dibanding brute-force.", options: { breakLine: true } },
        { text: "✔ Multi-Target: Parameter disesuaikan spesifik untuk target VAR, Sharpe, atau Rachev.", options: { italic: true } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 16, color: "333333", valign: "top" });
    slideGridSearchEx.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slideGridSearchEx.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Slide 23.19: Lampiran - Justifikasi Pemilihan Rolling Window ---
    let slideWindowJust = pres.addSlide();
    slideWindowJust.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slideWindowJust.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slideWindowJust.addText("Lampiran: Justifikasi Pemilihan Rolling Window", { x: 0.5, y: 0.5, w: "90%", fontSize: 24, bold: true, color: "003366" });
    slideWindowJust.addText([
        { text: "Kenapa rentang 30 s/d 60 hari dipilih sebagai 'Sweet Spot'?", options: { bold: true, breakLine: true } },

        { text: "1. Alasan Teknis (Kenapa Tidak < 30 Hari?):", options: { bold: true, breakLine: true, color: "c0392b" } },
        { text: "   • Ill-Conditioned Matrix: Data terlalu sedikit (T ≈ N) membuat kalkulasi bobot tidak stabil dan acak.", options: { breakLine: true } },
        { text: "   • RMT Failure: Batas noise menjadi sangat lebar, filter RMT akan menghapus hampir semua sinyal pasar.", options: { breakLine: true } },
        { text: "   • High Turnover: Bobot berubah terlalu liar, keuntungan habis dimakan biaya transaksi (fees).", options: { breakLine: true } },

        { text: "2. Alasan Strategis (Kenapa Tidak > 120 Hari?):", options: { bold: true, breakLine: true, color: "f39c12" } },
        { text: "   • Information Lag: Model terlambat mendeteksi crash atau perubahan rezim karena terbebani data setahun lalu.", options: { breakLine: true } },
        { text: "   • Regime Blurring: Mencampur data Bull & Bear Market dalam satu hitungan mengaburkan risiko asli saat ini.", options: { breakLine: true } },

        { text: "3. Solusi 30-60 Hari (Optimal):", options: { bold: true, breakLine: true, color: "27ae60" } },
        { text: "   • Responsif: Cepat menangkap narasi pasar baru (misal: siklus Altcoin).", options: { breakLine: true } },
        { text: "   • Stabil: Cukup data untuk menghasilkan peta jaringan MST yang jujur dan korelasi yang valid.", options: { breakLine: true } },

        { text: "Kesimpulan:", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "Window 30-60 hari adalah keseimbangan antara akurasi statistik dan kecepatan respons pasar.", options: { italic: true } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 16, color: "333333", valign: "top" });
    slideWindowJust.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slideWindowJust.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });
    slideCalcCentrality.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slideCalcCentrality.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });













    // --- Slide 24: Lampiran - Dua Tipe Grid Search ---
    let slide24 = pres.addSlide();
    slide24.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slide24.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slide24.addText("Lampiran: Dua Tipe Pendekatan Grid Search", { x: 0.5, y: 0.5, w: "90%", fontSize: 26, bold: true, color: "003366" });
    slide24.addText([

        { text: "Dua Objektif Optimasi (Return vs Risk):", options: { bold: true, breakLine: true } },
        { text: "1. Network Markowitz dengan Target Return (NW - Return Grid Search):", options: { bold: true, breakLine: true, color: "27ae60" } },
        { text: "Memaksimalkan capaian ", options: { bullet: true } },
        { text: "imbal hasil", options: { bold: true } },
        { text: " portofolio.", options: { breakLine: true } },
        { text: "Lebih Agresif untuk mengeksploitasi reli pada pasar ", options: { bullet: true } },
        { text: "Bullish", options: { bold: true } },
        { text: ".", options: { breakLine: true } },
        { text: "2. Network Markowitz dengan Target Risiko (NW - Risk Grid Search):", options: { bold: true, breakLine: true, color: "c0392b" } },
        { text: "Menekan parameter ", options: { bullet: true } },
        { text: "risiko total", options: { bold: true } },
        { text: " hingga minimal.", options: { breakLine: true } },
        { text: "Lebih Defensif untuk meredam fluktuasi saat ", options: { bullet: true } },
        { text: "Crypto Winter", options: { bold: true } },
        { text: ".", options: { breakLine: true } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 18, color: "333333", valign: "top" });
    slide24.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slide24.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });
    // --- Slide 25: Lampiran - Penanganan Missing Value ---
    let slideOut1 = pres.addSlide();
    slideOut1.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slideOut1.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slideOut1.addText("Lampiran: Penanganan Data Kosong (Missing Values)", { x: 0.5, y: 0.5, w: "90%", fontSize: 26, bold: true, color: "003366" });
    slideOut1.addText([
        { text: "Pertanyaan: ", options: { bold: true, color: "c0392b" } },
        { text: "\"Beberapa koin seperti Binance/EOS belum rilis di awal 2017 sehingga datanya kosong. Bukankah backward-fill memalsukan harga dan merusak matriks?\"", options: { italic: true, breakLine: true } },

        { text: "Jawaban (Kenapa RMT sangat Krusial):", options: { bold: true, breakLine: true, color: "27ae60" } },
        { text: "   1. Praktik pengisian data diam (", options: {} },
        { text: "backward fill", options: { italic: true, bold: true } },
        { text: ") memang menciptakan rentetan nilai harga yang statis.", options: { breakLine: true } },
        { text: "   2. Namun, kehebatan ", options: {} },
        { text: "Random Matrix Theory (RMT)", options: { bold: true } },
        { text: " diuji di sini! Karena data yang datar sama sekali tidak punya korelasi nyata.", options: { breakLine: true } },
        { text: "   3. RMT otomatis akan mendeteksi korelasi buatan tersebut sebagai probabilitas ", options: {} },
        { text: "Noise Acak", options: { bold: true, italic: true } },
        { text: ", lalu membuangnya menjadi 0.", options: { breakLine: true } },
        { text: "   4. Hasilnya, matriks korelasi ", options: {} },
        { text: "terselamatkan", options: { bold: true } },
        { text: " dan tidak tercemar oleh cacat kelengkapan data.", options: { breakLine: true } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 18, color: "333333", valign: "top" });
    slideOut1.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slideOut1.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Slide 26: Lampiran - Peran USDT (Bagian 1) ---
    let slideOut2 = pres.addSlide();
    slideOut2.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slideOut2.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slideOut2.addText("Lampiran: Mengapa Menyertakan Tether (USDT)? (1/2)", { x: 0.5, y: 0.5, w: "90%", fontSize: 26, bold: true, color: "003366" });
    slideOut2.addText([
        { text: "Pertanyaan: ", options: { bold: true, color: "c0392b" } },
        { text: "\"Tether (USDT) itu stablecoin yang nilainya selalu fix ke 1 USD. Apakah tidak berbuat curang?\"", options: { italic: true, breakLine: true } },

        { text: "Jawaban (Dinamika Portofolio Cerdas):", options: { bold: true, breakLine: true, color: "27ae60" } },
        { text: "Model ini mensimulasikan perilaku ", options: { bullet: true } },
        { text: "Robo-Advisory", options: { bold: true, italic: true } },
        { text: " institusional.", options: { breakLine: true } },
        { text: "Ketika pasar anjlok ekstrem, investor akan melarikan senjatanya ke posisi ", options: { bullet: true } },
        { text: "tunai (USDT)", options: { bold: true } },
        { text: " sebagai evakuasi risiko.", options: {} }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 18, color: "333333", valign: "top" });
    slideOut2.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slideOut2.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Slide 27: Lampiran - Peran USDT (Bagian 2) ---
    let slideOut3 = pres.addSlide();
    slideOut3.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slideOut3.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slideOut3.addText("Lampiran: Mengapa Menyertakan Tether (USDT)? (2/2)", { x: 0.5, y: 0.5, w: "90%", fontSize: 26, bold: true, color: "003366" });
    slideOut3.addText([
        { text: "3. Algoritma jaringan ", options: {} },
        { text: "Network Markowitz GS (Grid Search)", options: { bold: true } },
        { text: " dilatih secara matematis; jika mendeteksi korelasi ancaman kolaps merambat ke semua altcoin, ia akan melempar alokasi modalnya menuju node ", options: {} },
        { text: "USDT (Tether)", options: { bold: true } },
        { text: " sebagai langkah ", options: {} },
        { text: "evakuasi otomatis", options: { bold: true } },
        { text: " (Shock-Absorber).", options: { breakLine: true } },
        { text: "   4. Hal ini yang membuat performa Risk-GS (Grid Search) sangat ", options: {} },
        { text: "tangguh", options: { bold: true } },
        { text: " dari serangan Crypto Winter, suatu kapabilitas pertahanan yang tidak dipahami oleh model ortodoks murni Markowitz.", options: { breakLine: true } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 18, color: "333333", valign: "top" });
    slideOut3.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slideOut3.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Slide 28: Lampiran - Justifikasi Akademik 1: Non-Stationarity ---
    let slideOut4 = pres.addSlide();
    slideOut4.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slideOut4.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slideOut4.addText("Lampiran: Bukti Empiris Non-Stationarity Pasar", { x: 0.5, y: 0.5, w: "90%", fontSize: 26, bold: true, color: "003366" });
    slideOut4.addText([
        { text: "Masalah: ", options: { bold: true, color: "c0392b" } },
        { text: "Kenapa harus menggunakan model dengan Tuning Parameter?", options: { italic: true, breakLine: true } },
        { text: "Bukti dari Grid Search:", options: { bold: true, breakLine: true, color: "27ae60" } },
        { text: "Data menunjukkan nilai Gamma (γ) optimal ", options: { bullet: true } },
        { text: "terus bergeser", options: { bold: true } },
        { text: " setiap periode rebalancing.", options: { breakLine: true } },
        { text: "Penggunaan γ statis tidak cukup untuk menangkap perubahan ", options: { bullet: true } },
        { text: "struktur korelasi", options: { bold: true } },
        { text: " yang sangat cepat di pasar kripto.", options: { breakLine: true } },
        { text: "Ini membenarkan bahwa pasar kripto membutuhkan ", options: { bullet: true } },
        { text: "kalibrasi otomatis", options: { bold: true } },
        { text: " secara temporal.", options: {} }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 18, color: "333333", valign: "top" });
    slideOut4.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slideOut4.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Slide 29: Lampiran - Justifikasi Akademik 2: Strategi Shock-Absorber ---
    let slideOut5 = pres.addSlide();
    slideOut5.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slideOut5.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });
    slideOut5.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slideOut5.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slideOut5.addText("Lampiran: Jaringan sebagai 'Shock-Absorber'", { x: 0.5, y: 0.5, w: "90%", fontSize: 26, bold: true, color: "003366" });
    slideOut5.addText([
        { text: "Konsep Teoritis:", options: { bold: true, breakLine: true, color: "27ae60" } },
        { text: "Eigenvector Centrality mengidentifikasi koin yang menjadi ", options: { bullet: true } },
        { text: "hub risiko", options: { bold: true } },
        { text: ".", options: { breakLine: true } },
        { text: "Penalti Gamma memaksa portofolio ", options: { bullet: true } },
        { text: "menjauhi aset", options: { bold: true } },
        { text: " yang terlalu dominan secara sistemik saat volatilitas tinggi.", options: { breakLine: true } },
        { text: "Hasil pada ", options: { bullet: true } },
        { text: "Tabel 5", options: { bold: true } },
        { text: " membuktikan bahwa saat crash, distribusi kerugian ekor model NW (Network Markowitz) jauh lebih terjaga.", options: { breakLine: true } },
        { text: "Kesimpulan: Topologi jaringan memberikan sinyal ", options: { bullet: true } },
        { text: "diversifikasi akurat", options: { bold: true } },
        { text: " daripada sekadar variansi harga.", options: {} }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 18, color: "333333", valign: "top" });
    // --- Slide 30: Lampiran - Simulasi Kalkulasi Cumulative P&L ---
    let slidePnLSim = pres.addSlide();
    slidePnLSim.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slidePnLSim.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slidePnLSim.addText("Lampiran: Simulasi Kalkulasi Cumulative P&L", { x: 0.5, y: 0.5, w: "90%", fontSize: 24, bold: true, color: "003366" });
    slidePnLSim.addText([
        { text: "Bagaimana modal berkembang melalui bunga majemuk (compounding)?", options: { bold: true, breakLine: true } },
        { text: "1. Konsep Dasar:", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "   Cumulative P&L mengukur total keuntungan atau kerugian bersih dari awal investasi.", options: { breakLine: true } },

        { text: "2. Simulasi 3 Hari (Modal Awal: Rp 10.000.000):", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "   • Hari 1: Return +2% ", options: {} },
        { text: "→ Untung Rp 200.000 (Modal: 10,2jt)", options: { italic: true, breakLine: true } },
        { text: "   • Hari 2: Return -1% ", options: {} },
        { text: "→ Rugi Rp 102.000 (Modal: 10,098jt)", options: { italic: true, breakLine: true } },
        { text: "   • Hari 3: Return +5% ", options: {} },
        { text: "→ Untung Rp 504.900 (Modal: 10,6029jt)", options: { italic: true, breakLine: true } },

        { text: "3. Rumus Akumulasi (Compounding Index):", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "   Cum_Return = (1 + r1) × (1 + r2) × ... × (1 + rn) - 1", options: { fontFace: "Courier New", color: "c0392b", bold: true, breakLine: true } },
        { text: "   Hasil Simulasi: (1.02 × 0.99 × 1.05) - 1 ≈ ", options: { fontFace: "Courier New" } },
        { text: "6.03%", options: { bold: true, color: "27ae60", fontFace: "Courier New" } },

        { text: "Kelebihan Metrik Ini:", options: { bold: true, breakLine: true, color: "8e44ad" } },
        { text: "Memberikan gambaran riil 'kekuatan bertahan' sebuah strategi. Meskipun ada hari-hari rugi (drawdown), akumulasi positif menunjukkan resiliensi portofolio.", options: { italic: true } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 16, color: "333333", valign: "top" });
    slidePnLSim.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slidePnLSim.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Slide 32: Lampiran - Simulasi Kalkulasi Value at Risk (VaR) ---
    let slideVaRSim = pres.addSlide();
    slideVaRSim.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slideVaRSim.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slideVaRSim.addText("Lampiran: Simulasi Kalkulasi Value at Risk (VaR)", { x: 0.5, y: 0.5, w: "90%", fontSize: 24, bold: true, color: "003366" });
    slideVaRSim.addText([
        { text: "Berapa potensi kerugian maksimal dalam kondisi pasar normal?", options: { bold: true, breakLine: true } },
        { text: "1. Konsep Dasar:", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "   Value at Risk (VaR) mengukur ambang batas kerugian maksimal pada tingkat kepercayaan tertentu (Confidence Level 95%) dalam periode harian.", options: { breakLine: true } },

        { text: "2. Metode Simulasi Historis (Paling Sederhana):", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "   Mengurutkan 100 hari return dari terburuk ke terbaik, lalu mengambil nilai persentil ke-5.", options: { breakLine: true } },

        { text: "3. Simulasi Angka (Modal: Rp 10.000.000):", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "   • Return terburuk urutan ke-5 (P5): -8% (-0.08)", options: { breakLine: true } },
        { text: "   • Perhitungan VaR (95%): Rp 10.000.000 × 0.08 = ", options: {} },
        { text: "Rp 800.000", options: { bold: true, color: "c0392b", breakLine: true } },

        { text: "4. Interpretasi Hasil:", options: { bold: true, breakLine: true, color: "8e44ad" } },
        { text: "Kita yakin 95% bahwa kerugian harian tidak akan melebihi Rp 800.000. Namun, ada 5% risiko (ekstrem) di masa depan di mana kerugian bisa lebih besar dari angka tersebut.", options: { italic: true } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 16, color: "333333", valign: "top" });
    slideVaRSim.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slideVaRSim.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Slide 33: Lampiran - Simulasi Kalkulasi Sharpe Ratio ---
    let slideSharpeSim = pres.addSlide();
    slideSharpeSim.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slideSharpeSim.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slideSharpeSim.addText("Lampiran: Simulasi Kalkulasi Sharpe Ratio", { x: 0.5, y: 0.5, w: "90%", fontSize: 24, bold: true, color: "003366" });
    slideSharpeSim.addText([
        { text: "Bagaimana mengukur kualitas imbal hasil per unit risiko?", options: { bold: true, breakLine: true } },
        { text: "1. Konsep Dasar:", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "   Sharpe Ratio mengukur kelebihan imbal hasil (Excess Return) dibandingkan aset aman (Risk-Free) per satu satuan risiko (Volatilitas).", options: { breakLine: true } },

        { text: "2. Rumus Sederhana:", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "   SR = (Return Portofolio - Risk-Free Rate) / Standar Deviasi", options: { fontFace: "Courier New", color: "c0392b", bold: true, breakLine: true } },

        { text: "3. Simulasi Angka (Dummy):", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "   • Return Portofolio (Rp): 15% (0.15)", options: { breakLine: true } },
        { text: "   • Suku Bunga Bebas Risiko (Rf): 2% (0.02) ", options: { italic: true } },
        { text: " (Basis: USDT Lending Rate)", options: { fontSize: 12, breakLine: true } },
        { text: "   • Volatilitas Portofolio (σ): 10% (0.10)", options: { breakLine: true } },

        { text: "4. Kalkulasi:", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "   SR = (0.15 - 0.02) / 0.10 = 0.13 / 0.10 = ", options: { fontFace: "Courier New" } },
        { text: "1.30", options: { bold: true, color: "27ae60", fontFace: "Courier New" } },

        { text: "Interpretasi:", options: { bold: true, breakLine: true, color: "8e44ad" } },
        { text: "Nilai 1.30 berarti untuk setiap 1% risiko yang diambil, portofolio memberikan imbal hasil tambahan sebesar 1.30%. Semakin tinggi nilai SR, semakin efisien portofolio tersebut.", options: { italic: true } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 16, color: "333333", valign: "top" });
    slideSharpeSim.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slideSharpeSim.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Slide 34: Lampiran - Simulasi Kalkulasi Rachev Ratio ---
    let slideRachevSim = pres.addSlide();
    slideRachevSim.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slideRachevSim.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slideRachevSim.addText("Lampiran: Simulasi Kalkulasi Rachev Ratio", { x: 0.5, y: 0.5, w: "90%", fontSize: 24, bold: true, color: "003366" });
    slideRachevSim.addText([
        { text: "1. Konsep Dasar:", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "   Rachev Ratio mengukur rasio potensi 'hadiah ekstrem' (Top 5% average) terhadap 'risiko ekstrem' (Bottom 5% average loss).", options: { breakLine: true } },

        { text: "2. Skenario A: Normal/Bullish (Menguntungkan):", options: { bold: true, breakLine: true, color: "27ae60" } },
        { text: "   • Avg. Profit (Top 5%): +12% | Avg. Loss (Bottom 5%): -8%", options: { breakLine: true } },
        { text: "   • Kalkulasi: 12% / 8% = ", options: {} },
        { text: "1.50", options: { bold: true, color: "27ae60", breakLine: true } },

        { text: "3. Skenario B: Extreme Bear Market (Keduanya Rugi):", options: { bold: true, breakLine: true, color: "c0392b" } },
        { text: "   • Jika kondisi pasar hancur total, bahkan return terbaik pun bisa negatif.", options: { fontSize: 12, italic: true, breakLine: true } },
        { text: "   • Avg. Return (Top 5%): ", options: {} },
        { text: "-2% (Rugi)", options: { bold: true } },
        { text: " | Avg. Loss (Bottom 5%): ", options: {} },
        { text: "-12% (Rugi)", options: { bold: true, breakLine: true } },
        { text: "   • Kalkulasi: -2% / 12% = ", options: {} },
        { text: "-0.16", options: { bold: true, color: "c0392b", breakLine: true } },

        { text: "Interpretasi:", options: { bold: true, breakLine: true, color: "8e44ad" } },
        { text: "Nilai negatif (Skenario B) menunjukkan strategi gagal total memberikan profit. Sebaliknya, nilai > 1 (Skenario A) menunjukkan portofolio memiliki potensi pemulihan yang jauh lebih besar daripada risiko kejatuhannya.", options: { italic: true } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 16, color: "333333", valign: "top" });
    slideRachevSim.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slideRachevSim.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Akhir Presentasi: Terima Kasih ---
    let slide7 = pres.addSlide();
    slide7.addImage({ path: "bg_watermark.png", x: 0, y: 0, w: "100%", h: "100%" });
    slide7.addImage({ path: "logo_unm.png", x: 9.1, y: 0.1, w: 0.7, h: 0.7 });
    slide7.addText("Terima Kasih", { x: 0.5, y: 2.7, w: "90%", fontSize: 40, bold: true, align: "center", color: "003366" });
    slide7.addText("📂 Lampiran", { x: 7.3, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '3' }, align: "right" });
    slide7.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });


    // --- Simpan File ---
    const outputFilename = "Presentasi_Proposal_Update.pptx";

    try {
        await pres.writeFile({ fileName: outputFilename });
        console.log(`Bagus! Presentasi berhasil disimpan sebagai: ${outputFilename}`);
    } catch (error) {
        console.error("Terjadi kesalahan saat menyimpan presentasi:", error);
    }
}

// Menjalankan fungsi
createPresentation();
