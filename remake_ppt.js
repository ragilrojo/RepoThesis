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

    // Kolom Kiri: I & II
    slideTOC.addText([
        { text: "I. PENDAHULUAN & MASALAH", options: { bold: true, color: "003366", fontSize: 16, breakLine: true } },
        { text: "  \u2022 ", options: { color: "7f8c8d" } },
        { text: "Latar Belakang", options: { fontSize: 13, color: "0563C1", underline: true, hyperlink: { slide: '3' } } },
        { text: "\n", options: { breakLine: true } },
        { text: "  \u2022 Rumusan Masalah\n", options: { fontSize: 13, color: "333333" } },
        { text: "  \u2022 Tujuan Penelitian\n\n", options: { fontSize: 13, color: "333333" } },

        { text: "II. LANDASAN & SIMULASI", options: { bold: true, color: "003366", fontSize: 16, breakLine: true } },
        { text: "  \u2022 ", options: { color: "7f8c8d" } },
        { text: "Landasan Teori (Markowitz)", options: { fontSize: 13, color: "0563C1", underline: true, hyperlink: { slide: '4' } } },
        { text: "\n", options: { breakLine: true } },
        { text: "  \u2022 Kerangka Pemikiran\n", options: { fontSize: 13, color: "333333" } }
    ], { x: 0.5, y: 1.2, w: 4.5, h: 4.0, valign: "top" });

    // Kolom Kanan: III & IV
    slideTOC.addText([
        { text: "III. METODOLOGI & SAC-NET", options: { bold: true, color: "003366", fontSize: 16, breakLine: true } },
        { text: "  \u2022 ", options: { color: "7f8c8d" } },
        { text: "Network Markowitz", options: { fontSize: 13, color: "0563C1", underline: true, hyperlink: { slide: '5' } } },
        { text: "\n", options: { breakLine: true } },
        { text: "  \u2022 Mekanisme SAC Agent\n", options: { fontSize: 13, color: "333333" } },
        { text: "  \u2022 Parameter Eksperimen\n\n", options: { fontSize: 13, color: "333333" } },

        { text: "IV. EVALUASI & HASIL", options: { bold: true, color: "003366", fontSize: 16, breakLine: true } },
        { text: "  \u2022 ", options: { color: "7f8c8d" } },
        { text: "Evaluasi (Sharpe Ratio)", options: { fontSize: 13, color: "0563C1", underline: true, hyperlink: { slide: '6' } } },
        { text: "\n", options: { breakLine: true } },
        { text: "  \u2022 Analisis Performa\n", options: { fontSize: 13, color: "333333" } },
        { text: "  \u2022 Kesimpulan & Saran", options: { fontSize: 13, color: "333333" } }
    ], { x: 5.2, y: 1.2, w: 4.5, h: 4.0, valign: "top" });

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

    slideNet.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

    // --- Slide 6: Simulasi Perhitungan Sharpe Ratio ---
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

    // Box Penjabaran Kalkulasi
    slideSharpe.addShape(pres.ShapeType.rect, { x: 6.2, y: 1.5, w: 3.5, h: 3.7, fill: { color: "fff8ef" }, line: { color: "e67e22", width: 1.5 } });
    slideSharpe.addText([
        { text: "Langkah Kalkulasi:\n", options: { bold: true, color: "d35400", fontSize: 13, breakLine: true } },
        { text: "1. Mean Return (Rp):\n", options: { bold: true, fontSize: 10 } },
        { text: "(0.7+1.1+0.2+0.6)/4 = ", options: { fontSize: 10 } }, { text: "0.65%\n\n", options: { bold: true } },

        { text: "2. Std Dev / Risk (\u03c3p):\n", options: { bold: true, fontSize: 10 } },
        { text: "Mengukur fluktuasi harian.\n", options: { fontSize: 10 } },
        { text: "Misal \u03c3p = ", options: { fontSize: 10 } }, { text: "0.33%\n\n", options: { bold: true } },

        { text: "3. Sharpe Ratio (Rf = 0%):\n", options: { bold: true, fontSize: 10 } },
        { text: "Sharpe = (0.65 - 0) / 0.33\n", options: { fontSize: 10 } },
        { text: "Sharpe = ", options: { fontSize: 11 } }, { text: "1.97", options: { bold: true, color: "27ae60", fontSize: 13 } },

        { text: "\n------------------------------------------\n", options: {} },
        { text: "Kesimpulan:\n", options: { bold: true, fontSize: 10 } },
        { text: "Aset ini sangat efisien karena return rata-ratanya hampir 2x lipat dari fluktuasi risikonya.", options: { fontSize: 9 } }
    ], { x: 6.4, y: 1.7, w: 3.2, color: "333333", valign: "top" });

    slideSharpe.addText("🏠 Daftar Isi", { x: 8.5, y: 5.3, w: 1.2, fontSize: 10, color: "0563C1", underline: true, hyperlink: { slide: '2' }, align: "right" });

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
