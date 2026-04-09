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
    slide1.addText("Proposal Tesis:\nOptimalisasi Portofolio Adaptif", { 
        x: 0.5, y: 1.2, w: "90%", fontSize: 40, bold: true, align: "center", color: "003366" 
    });
    slide1.addText([
        { text: "Berbasis Pendekatan\n" },
        { text: "Network Markowitz", options: { italic: true } },
        { text: " dengan " },
        { text: "Rolling Window Grid Search", options: { italic: true } }
    ], { 
        x: 0.5, y: 2.6, w: "90%", fontSize: 24, align: "center", color: "34495e" 
    });
    slide1.addText("Oleh: Ragil Yulianto", { 
        x: 0.5, y: 4.5, w: "90%", fontSize: 18, align: "center", color: "7f8c8d" 
    });

    // --- Slide 2: Latar Belakang ---
    let slide2 = pres.addSlide();
    slide2.addText("Latar Belakang", { x: 0.5, y: 0.5, w: "90%", fontSize: 28, bold: true, color: "003366" });
    slide2.addText([
        { text: "Tingginya volatilitas ekstrem dan ", options: { bullet: true } },
        { text: "noise", options: { italic: true, bold: true } },
        { text: " memerlukan manajemen risiko yang presisi.", options: { breakLine: true } },
        { text: "Kelemahan Markowitz: Rawan terhadap ", options: { bullet: true } },
        { text: "estimation error", options: { italic: true, bold: true } },
        { text: " pada matriks kovarians.", options: { breakLine: true } },
        { text: "Network Markowitz: Menyaring noise menggunakan ", options: { bullet: true } },
        { text: "Random Matrix Theory", options: { italic: true, bold: true } },
        { text: " (RMT) untuk rekonstruksi stabilitas." }
    ], { x: 0.5, y: 1.1, w: "90%", h: 3, fontSize: 20, color: "333333", valign: "top" });

    // --- Slide 3: Konsep "Noise" dalam Cryptocurrency ---
    let slide3 = pres.addSlide();
    slide3.addText("Apa itu \"Noise\" di Pasar Kripto?", { x: 0.5, y: 0.5, w: "90%", fontSize: 28, bold: true, color: "003366" });
    slide3.addText([
        { text: "Noise (Kebisingan) Pasar:", options: { bold: true, breakLine: true } },
        { text: "Fluktuasi harga acak", options: { bold: true } },
        { text: " akibat sentimen sesaat, rumor, FOMO (Fear Of Missing Out), atau spekulasi yang tidak mencerminkan nilai fundamental aset.", options: { breakLine: true } },
        { text: "", options: { breakLine: true } }, // Spasi antar poin

        { text: "Estimation Error (Korelasi Palsu):", options: { bold: true, breakLine: true } },
        { text: "Model tradisional seringkali menangkap pergerakan acak ini sebagai korelasi tinggi antar aset, menghasilkan ", options: { } },
        { text: "matriks kovarians", options: { bold: true } },
        { text: " yang ", options: { } },
        { text: "berisik", options: { italic: true } },
        { text: " dan tidak stabil.", options: { breakLine: true } },
        { text: "", options: { breakLine: true } }, // Spasi antar poin

        { text: "Solusi Random Matrix Theory (", options: { bold: true } },
        { text: "RMT", options: { bold: true, hyperlink: { slide: '9' }, color: "0563C1", underline: true } },
        { text: "):", options: { bold: true, breakLine: true } },
        { text: "Berfungsi sebagai ", options: { } },
        { text: "filter", options: { bold: true } },
        { text: " untuk memisahkan ", options: { } },
        { text: "korelasi sejati", options: { bold: true } },
        { text: " (sinyal struktur pasar) dari sekadar pergerakan kebetulan (", options: { } },
        { text: "noise", options: { italic: true } },
        { text: "), memastikan alokasi portofolio tidak tertipu oleh ", options: { } },
        { text: "fluktuasi semu", options: { bold: true } },
        { text: "." }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5, fontSize: 18, color: "333333", valign: "top" });

    // --- Slide 4: Landasan Teori (Dua Kolom) ---
    let slide4 = pres.addSlide();
    slide4.addText("Landasan Teori Utama", { x: 0.5, y: 0.5, w: "90%", fontSize: 28, bold: true, color: "003366" });
    
    // Kolom Kiri: Portofolio & Risiko
    slide4.addText([
        { text: "Tinjauan Portofolio & Risiko:", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "➤ ", options: { } },
        { text: "Portofolio", options: { bold: true, underline: true } },
        { text: ": Diversifikasi aset untuk optimasi return-risiko.", options: { breakLine: true } },
        { text: "➤ ", options: { } },
        { text: "Volatilitas", options: { bold: true, underline: true } },
        { text: ": Ukuran fluktuasi harga pasar.", options: { breakLine: true } },
        { text: "➤ ", options: { } },
        { text: "Kovarians", options: { bold: true, underline: true } },
        { text: ": Ukuran pergerakan bersama aset.", options: { breakLine: true } }
    ], { x: 0.5, y: 1.2, w: "45%", h: 4, fontSize: 20, color: "333333", valign: "top" });

    // Kolom Kanan: Struktur Jaringan
    slide4.addText([
        { text: "Pendekatan Struktur Jaringan:", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "➤ ", options: { } },
        { text: "RMT (Random Matrix Theory)", options: { bold: true, underline: true } },
        { text: ": Filter noise untuk kestabilan matriks.", options: { breakLine: true } },
        { text: "➤ ", options: { } },
        { text: "MST (Minimum Spanning Tree)", options: { bold: true, underline: true } },
        { text: ": Jaringan korelasi terkuat tanpa loop.", options: { breakLine: true } },
        { text: "➤ ", options: { } },
        { text: "Centrality", options: { bold: true, underline: true } },
        { text: ": Metrik risiko penularan sistemik.", options: { breakLine: true } }
    ], { x: 5.2, y: 1.2, w: "45%", h: 4, fontSize: 20, color: "333333", valign: "top" });

    // --- Slide 4.1: Penelitian Terdahulu ---
    let slidePrev = pres.addSlide();
    slidePrev.addText("Penelitian Terdahulu (State of the Art)", { x: 0.5, y: 0.5, w: "90%", fontSize: 28, bold: true, color: "003366" });
    slidePrev.addText([
        { text: "1. Giudici et al. (2020):", options: { bold: true, breakLine: true, color: "27ae60" } },
        { text: "Pelopor model Network Markowitz yang memadukan ", options: { } },
        { text: "RMT (Random Matrix Theory)", options: { bold: true } },
        { text: " dan ", options: { } },
        { text: "MST (Minimum Spanning Tree)", options: { bold: true } },
        { text: " di kripto.", options: { breakLine: true } },
        { text: "2. Kitanovski et al. (2022):", options: { bold: true, breakLine: true, color: "8e44ad" } },
        { text: "Mendemonstrasikan diversifikasi berbasis ", options: { } },
        { text: "deteksi komunitas", options: { bold: true } },
        { text: " jaringan.", options: { breakLine: true } },
        { text: "3. Jing & Rocha (2023):", options: { bold: true, breakLine: true, color: "c0392b" } },
        { text: "Pemilihan koin topologi MST (Minimum Spanning Tree) mengalahkan ", options: { } },
        { text: "semua benchmark", options: { bold: true } },
        { text: ".", options: { breakLine: true } },
        { text: "4. Kitanovski et al. (2024):", options: { bold: true, breakLine: true, color: "16a085" } },
        { text: "Penalti graf sangat resilien meredam ", options: { } },
        { text: "eksposur ekstrem", options: { bold: true } },
        { text: ".", options: { breakLine: true } },
        { text: "5. Jing et al. (2025):", options: { bold: true, breakLine: true, color: "f39c12" } },
        { text: "Penggabungan Network-MPT (Modern Portfolio Theory) memberikan ", options: { } },
        { text: "prediksi stabil", options: { bold: true } },
        { text: " di fase terbaru.", options: { breakLine: true } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.0, fontSize: 16, color: "333333", valign: "top" });

    // --- Slide 4.5: Kerangka Penelitian ---
    let slideFramework = pres.addSlide();
    slideFramework.addText("Kerangka Pemikiran / Penelitian", { x: 0.5, y: 0.5, w: "90%", fontSize: 28, bold: true, color: "003366" });
    slideFramework.addImage({ path: "framwrok.jpg", x: 1.0, y: 1.1, w: 8.0, h: 4.0 });

    // --- Slide 4.6: Dataset - 10 Aset Kripto Utama ---
    let slideData = pres.addSlide();
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
        x: 0.5, y: 1.1, w: 9.0,
        colWidths: [1.2, 2.5, 5.3],
        border: { type: "solid", color: "cccccc", pt: 1 },
        fontSize: 16,
        color: "333333"
    });

    // --- Slide 5: Strategi yang Dibandingkan ---
    let slide5 = pres.addSlide();
    slide5.addText("Strategi Portofolio yang Disimulasikan", { x: 0.5, y: 0.5, w: "90%", fontSize: 28, bold: true, color: "003366" });
    slide5.addText([
        { text: "1. Kelompok Baseline:", options: { bold: true, color: "003366", breakLine: true } },
        { text: "   • EW (Equally Weighted)", options: { breakLine: true } },
        { text: "   • CM (Classical Markowitz)", options: { breakLine: true } },
        { text: "", options: { breakLine: true } },

        { text: "2. Kelompok Regularisasi:", options: { bold: true, color: "003366", breakLine: true } },
        { text: "   • GM (Glasso Markowitz)", options: { breakLine: true } },
        { text: "", options: { breakLine: true } },

        { text: "3. Kelompok Network (Statis):", options: { bold: true, color: "003366", breakLine: true } },
        { text: "   • NW Statis (γ fixed)", options: { breakLine: true } },
        { text: "", options: { breakLine: true } },

        { text: "4. Kelompok Network (Adaptif):", options: { bold: true, color: "003366", breakLine: true } },
        { text: "   • NW Adaptif (Grid Search)", options: {} }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5, fontSize: 22, color: "333333", valign: "top" });

    // --- Slide 5.1: Equally Weighted (EW) ---
    let slideEW = pres.addSlide();
    slideEW.addText("1.1. Equally Weighted (EW)", { x: 0.5, y: 0.5, w: "90%", fontSize: 28, bold: true, color: "003366" });
    slideEW.addText([
        { text: "Konsep Dasar:", options: { bold: true, breakLine: true } },
        { text: "Strategi alokasi ", options: { bullet: true } },
        { text: "1/N", options: { bold: true } },
        { text: " tanpa mempertimbangkan ", options: { } },
        { text: "parameter risiko", options: { bold: true } },
        { text: ".", options: { breakLine: true } },
        { text: "Keunggulan:", options: { bold: true, breakLine: true } },
        { text: "Berfungsi sebagai ", options: { bullet: true } },
        { text: "benchmark naif", options: { bold: true } },
        { text: " yang tangguh.", options: { breakLine: true } },
        { text: "Tidak memiliki ", options: { bullet: true } },
        { text: "estimation risk", options: { bold: true } },
        { text: " karena minim statistik.", options: { } }
    ], { x: 0.5, y: 1.2, w: "90%", h: 4, fontSize: 20, color: "333333", valign: "top" });

    // --- Slide 5.2: Classical Markowitz (CM) ---
    let slideCM = pres.addSlide();
    slideCM.addText("1.2. Classical Markowitz (CM)", { x: 0.5, y: 0.5, w: "90%", fontSize: 28, bold: true, color: "003366" });
    slideCM.addText([
        { text: "Konsep Dasar:", options: { bold: true, breakLine: true } },
        { text: "Meminimalkan variansi untuk tingkat ", options: { bullet: true } },
        { text: "imbal hasil", options: { bold: true } },
        { text: " tertentu.", options: { breakLine: true } },
        { text: "Kelemahan:", options: { bold: true, breakLine: true } },
        { text: "Menderita ", options: { bullet: true } },
        { text: "ketidakstabilan numerik", options: { bold: true } },
        { text: " pada data yang ", options: { } },
        { text: "sangat berisik", options: { bold: true } },
        { text: " (noisy).", options: { breakLine: true } },
        { text: "Pondasi dasar sebagai ", options: { bullet: true } },
        { text: "teori tradisional", options: { bold: true } },
        { text: " dalam penelitian ini.", options: { } }
    ], { x: 0.5, y: 1.2, w: "90%", h: 4, fontSize: 20, color: "333333", valign: "top" });

    // --- Slide 5.3: Graphical Lasso Markowitz (GM) ---
    let slideGM = pres.addSlide();
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
        { text: " (spurious correlations).", options: { } }
    ], { x: 0.5, y: 1.2, w: "90%", h: 4, fontSize: 20, color: "333333", valign: "top" });

    // --- Slide 5.4: Network Markowitz (NW) Statis ---
    let slideNWStatic = pres.addSlide();
    slideNWStatic.addText("3. Network Markowitz (NW) Statis", { x: 0.5, y: 0.5, w: "90%", fontSize: 28, bold: true, color: "003366" });
    slideNWStatic.addText([
        { text: "Konsep Dasar:", options: { bold: true, breakLine: true } },
        { text: "Model jaringan original (Giudici et al., 2020) yang menggabungkan ", options: { bullet: true } },
        { text: "filter RMT (Random Matrix Theory)", options: { bold: true } },
        { text: " dan ", options: { } },
        { text: "penalti sentralitas graf", options: { bold: true } },
        { text: ".", options: { breakLine: true } },
        { text: "Karakteristik:", options: { bold: true, breakLine: true } },
        { text: "Menggunakan parameter penalti (gamma) yang bersifat ", options: { bullet: true } },
        { text: "statis/tetap", options: { bold: true } },
        { text: " (hard-coded).", options: { breakLine: true } },
        { text: "Digunakan sebagai ", options: { bullet: true } },
        { text: "pembanding langsung", options: { bold: true } },
        { text: " untuk menguji efisiensi parameter adaptif.", options: { } }
    ], { x: 0.5, y: 1.2, w: "90%", h: 4, fontSize: 20, color: "333333", valign: "top" });

    // --- Slide 5.5: Network Markowitz (NW) Adaptif (Dua Kolom) ---
    let slideNWAdaptive = pres.addSlide();
    slideNWAdaptive.addText("4. Network Markowitz (NW) Adaptif", { x: 0.5, y: 0.5, w: "90%", fontSize: 28, bold: true, color: "003366" });
    // Kolom Kiri: Mekanisme
    slideNWAdaptive.addText([
        { text: "Mekanisme Optimasi Dinamis:", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "❑ Rolling Window: 30, 60, 90, 120 hari.", options: { breakLine: true } },
        { text: "❑ Rebalancing: Setiap 7 hari.", options: { breakLine: true } },
        { text: "❑ Transaction Cost: 0.1% (10 basis points).", options: { breakLine: true } },
        { text: "❑ Grid Search γ: Rentang [0.0 - 2.0].", options: { breakLine: true } }
    ], { x: 0.5, y: 1.2, w: "45%", h: 4, fontSize: 20, color: "333333", valign: "top" });

    // Kolom Kanan: Validasi
    slideNWAdaptive.addText([
        { text: "Split Validasi Internal (80/20):", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "❑ Training: Estimasi bobot gamma.", options: { breakLine: true } },
        { text: "❑ Validation: Seleksi performa optimal.", options: { breakLine: true } },
        { text: "❑ Fallback: Menggunakan strategi EW (Equally Weighted).", options: { breakLine: true } },
        { text: "❑ Tujuan: Adaptasi rezim pasar.", options: { breakLine: true } }
    ], { x: 5.2, y: 1.2, w: "45%", h: 4, fontSize: 20, color: "333333", valign: "top" });

    // --- Slide 6: Matriks Evaluasi Performa ---
    let slide6 = pres.addSlide();
    slide6.addText("Matriks Evaluasi Performa", { x: 0.5, y: 0.5, w: "90%", fontSize: 28, bold: true, color: "003366" });
    slide6.addText([
        { text: "1. Sharpe Ratio:", options: { bold: true, breakLine: true, color: "27ae60" } },
        { text: "   Imbal hasil per unit risiko. Semakin besar menunjukkan kualitas ", options: { } },
        { text: "efisiensi portofolio", options: { bold: true } },
        { text: ".", options: { breakLine: true } },
        { text: "2. Value at Risk (VaR):", options: { bold: true, breakLine: true, color: "c0392b" } },
        { text: "   Batas ", options: { } },
        { text: "kerugian maksimal", options: { bold: true } },
        { text: " kondiri crash. Semakin kecil tandanya perisai sukses.", options: { breakLine: true } },
        { text: "3. Rachev Ratio:", options: { bold: true, breakLine: true, color: "8e44ad" } },
        { text: "   Membandingkan potensi 'Profit Ekstrem' vs ", options: { } },
        { text: "Ancaman Loss", options: { bold: true } },
        { text: ". Menilai asimetri ekor.", options: { breakLine: true } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.0, fontSize: 18, color: "333333", valign: "top" });

    // --- Slide 7: Terima Kasih ---
    let slide7 = pres.addSlide();
    slide7.addText("Terima Kasih", { x: 0.5, y: 2.7, w: "90%", fontSize: 40, bold: true, align: "center", color: "003366" });

    // --- Slide 9: Lampiran - Analogi RMT ---
    let slide9 = pres.addSlide();
    slide9.addText([
        { text: "Lampiran: Analogi " },
        { text: "RMT (Random Matrix Theory)", options: { hyperlink: { slide: '9' } } },
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

    // --- Slide 10: Lampiran - Signal vs Noise ---
    let slide10 = pres.addSlide();
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

    // --- Slide 11: Lampiran - Menghitung Nilai Eigen ---
    let slide11 = pres.addSlide();
    slide11.addText("Lampiran: Bagaimana Menghitung Nilai Eigen?", { x: 0.5, y: 0.5, w: "90%", fontSize: 26, bold: true, color: "003366" });
    slide11.addText([
        { text: "1. Matriks Korelasi (C):", options: { bold: true, breakLine: true } },
        { text: "Membentuk tabel yang menjabarkan seluruh ", options: { bullet: true } },
        { text: "korelasi pergerakan", options: { bold: true } },
        { text: " harga antar sepasang koin.", options: { breakLine: true } },
        { text: "2. Konsep Persamaan Karakteristik:", options: { bold: true, breakLine: true } },
        { text: "Mencari besaran skalar ", options: { bullet: true } },
        { text: "λ (eigenvalue)", options: { italic: true, bold: true } },
        { text: " dan vektor arah yang memenuhi: ", options: { } },
        { text: "C × v = λ × v", options: { bold: true, color: "c0392b", breakLine: true } },
        { text: "3. Solusi Determinan:", options: { bold: true, breakLine: true } },
        { text: "Nilai λ adalah akar dari persamaan determinan: ", options: { bullet: true } },
        { text: "Det(C - λI) = 0", options: { bold: true, color: "c0392b", breakLine: true } },
        { text: "4. Arti dari Spektrum Hasil:", options: { bold: true, breakLine: true } },
        { text: "Nilai λ terbesar mewakili ", options: { bullet: true } },
        { text: "penggerak pasar", options: { bold: true } },
        { text: " utama (Market Factor).", options: { breakLine: true } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 18, color: "333333", valign: "top" });

    // --- Slide 12: Lampiran - Contoh Praktek (Dummy Data) ---
    let slide12 = pres.addSlide();
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
        { text: " lalu dinolkan, sementara λ₁ dijaga sebagai sinyal sejati.", options: { } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 16, color: "333333", valign: "top" });

    // --- Slide 13: Lampiran - Bagaimana Menghitung Korelasi? ---
    let slide13 = pres.addSlide();
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

    // --- Slide 14: Lampiran - Apakah Nilai Eigen Statis? ---
    let slide14 = pres.addSlide();
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
        { text: "Kesimpulan: Sifat adaptif secara ", options: { bullet: true } },
        { text: "real-time", options: { bold: true, italic: true } },
        { text: " merespons perubahan rezim pasar dengan cepat.", options: { } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 16, color: "333333", valign: "top" });

    // --- Slide 15: Lampiran - Batas Noise Marchenko-Pastur ---
    let slide15 = pres.addSlide();
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

    // --- Slide 16: Lampiran - Analogi Minimum Spanning Tree (Bagian 1) ---
    let slide16 = pres.addSlide();
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

    // --- Slide 17: Lampiran - Analogi Minimum Spanning Tree (Bagian 2) ---
    let slide17 = pres.addSlide();
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

    // --- Slide 18: Lampiran - Penalti (Gamma) Optimal ---
    let slide18 = pres.addSlide();
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

    // --- Slide 19: Lampiran - Classical Markowitz (Bagian 1) ---
    let slide19 = pres.addSlide();
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
        { text: " sebagai pedoman utama memprediksi masa depan.", options: { } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 18, color: "333333", valign: "top" });

    // --- Slide 20: Lampiran - Classical Markowitz (Bagian 2) ---
    let slide20 = pres.addSlide();
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
        { text: " menggunakan MST (Minimum Spanning Tree).", options: { } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 18, color: "333333", valign: "top" });

    // --- Slide 20.9: Lampiran - Detail Perhitungan Korelasi (Pearson) ---
    let slide20b = pres.addSlide();
    slide20b.addText("Lampiran: Detail Perhitungan Korelasi (Pearson)", { x: 0.5, y: 0.5, w: "90%", fontSize: 24, bold: true, color: "003366" });
    slide20b.addText([
        { text: "Bagaimana angka 0.50 didapatkan? (Dummy 3 Hari)", options: { bold: true, breakLine: true } },
        { text: "1. Data Return (X=BTC, Y=ETH):", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "   • Hari 1: X=2%, Y=2% | Hari 2: X=-2%, Y=0% | Hari 3: X=0%, Y=-2%", options: { breakLine: true } },
        { text: "   • Rata-rata (Mean): X̄ = 0, Ȳ = 0", options: { breakLine: true } },
        
        { text: "2. Tabel Deviasi & Perkalian:", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "   • Σ(x-x̄)(y-ȳ) = (2)(2) + (-2)(0) + (0)(-2) = ", options: { } },
        { text: "4", options: { bold: true, breakLine: true } },
        { text: "   • Σ(x-x̄)² = (2)² + (-2)² + (0)² = ", options: { } },
        { text: "8", options: { bold: true, breakLine: true } },
        { text: "   • Σ(y-ȳ)² = (2)² + (0)² + (-2)² = ", options: { } },
        { text: "8", options: { bold: true, breakLine: true } },

        { text: "3. Rumus Korelasi (ρ):", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "   ρ = Σ(dev_x * dev_y) / √(Σdev_x² * Σdev_y²)", options: { fontFace: "Courier New", color: "c0392b", breakLine: true } },
        { text: "   ρ = 4 / √(8 * 8) = 4 / 8 = ", options: { fontFace: "Courier New" } },
        { text: "0.50", options: { bold: true, color: "27ae60", fontFace: "Courier New" } },
        
        { text: "Kesimpulan:", options: { bold: true, breakLine: true, color: "c0392b" } },
        { text: "Angka ini menunjukkan arah pergerakan yang searah (positif) namun tidak identik, yang kemudian digunakan sebagai input matriks kovarians.", options: { italic: true } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 16, color: "333333", valign: "top" });

    // --- Slide 20.8: Lampiran - Contoh Sederhana Equally Weighted (EW) ---
    let slideEWExample = pres.addSlide();
    slideEWExample.addText("Lampiran: Contoh Sederhana Equally Weighted (EW)", { x: 0.5, y: 0.5, w: "90%", fontSize: 24, bold: true, color: "003366" });
    slideEWExample.addText([
        { text: "Strategi 1/N: Alokasi Tanpa Rumit", options: { bold: true, breakLine: true } },
        { text: "1. Logika Dasar:", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "   Membagi seluruh modal secara merata tanpa melihat performa masa lalu.", options: { breakLine: true } },
        
        { text: "2. Simulasi (Dummy):", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "   • Total Modal: Rp 100.000.000", options: { breakLine: true } },
        { text: "   • Jumlah Aset (N): 5 (BTC, ETH, XRP, LTC, USDT)", options: { breakLine: true } },
        { text: "   • Bobot Tiap Aset: 1/5 = ", options: { } },
        { text: "20%", options: { bold: true, color: "27ae60" } },
        { text: " (Tetap)", options: { breakLine: true } },
        
        { text: "3. Sebaran Modal:", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "   • BTC: Rp 20jt | ETH: Rp 20jt | XRP: Rp 20jt | dst.", options: { breakLine: true } },

        { text: "4. Mengapa EW Masuk Baseline?", options: { bold: true, breakLine: true, color: "c0392b" } },
        { text: "   • Tidak menderita ", options: { } },
        { text: "error estimasi", options: { bold: true } },
        { text: " karena tidak menghitung korelasi.", options: { breakLine: true } },
        { text: "   • Benchmark yang sangat tangguh; model kompleks harus bisa mengalahkan EW untuk dianggap valid.", options: { italic: true } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 18, color: "333333", valign: "top" });


    // --- Slide 21: Lampiran - Simulasi Sederhana Classical Markowitz (2 Aset) ---

    let slide21 = pres.addSlide();
    slide21.addText("Lampiran: Simulasi Sederhana Classical Markowitz (2 Aset)", { x: 0.5, y: 0.5, w: "90%", fontSize: 24, bold: true, color: "003366" });
    slide21.addText([
        { text: "Tujuan: Mencari bobot (w) untuk risiko terendah (Minimum Variance).", options: { bold: true, breakLine: true } },
        { text: "1. Data Input (Dummy):", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "   • Aset A (BTC): Volatilitas (σ₁) = 20% (0.20)", options: { breakLine: true } },
        { text: "   • Aset B (ETH): Volatilitas (σ₂) = 30% (0.30)", options: { breakLine: true } },
        { text: "   • Korelasi (ρ₁₂): 0.50 (Didapat dari return harian historis)", options: { breakLine: true } },
        { text: "   • Kovarians (σ₁₂): ", options: { bold: true } },
        { text: "ρ₁₂ × σ₁ × σ₂", options: { italic: true } },
        { text: " = 0.50 × 0.20 × 0.30 = ", options: { } },
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

    // --- Slide 22: Lampiran - Contoh GLasso (1/2: Pembersihan Korelasi) ---
    let slide22 = pres.addSlide();
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

    // --- Slide 23: Lampiran - Contoh GLasso (2/2: Dampak Bobot) ---
    let slide23 = pres.addSlide();
    slide23.addText("Lampiran: Contoh GLasso (2/2: Dampak Bobot)", { x: 0.5, y: 0.5, w: "90%", fontSize: 24, bold: true, color: "003366" });
    slide23.addText([
        { text: "Bagaimana 'Pembersihan' mengubah alokasi modal?", options: { bold: true, breakLine: true } },
        { text: "4. Simulasi Perubahan Bobot (Weight Shift):", options: { bold: true, breakLine: true, color: "003366" } },
        { text: "   • Alokasi Tanpa GLasso:", options: { bold: true, breakLine: true } },
        { text: "     BTC: 60%, ETH: 30%, ", options: { } },
        { text: "DOGE: 10%", options: { bold: true, color: "c0392b" } },
        { text: " (Tertipu korelasi palsu).", options: { breakLine: true } },
        
        { text: "   • Alokasi Dengan GLasso:", options: { bold: true, breakLine: true } },
        { text: "     BTC: 65%, ETH: 35%, ", options: { } },
        { text: "DOGE: 0%", options: { bold: true, color: "27ae60" } },
        { text: " (Fokus pada korelasi sejati).", options: { breakLine: true } },
        
        { text: "Kesimpulan & Manfaat:", options: { bold: true, breakLine: true, color: "27ae60" } },
        { text: "Pembersihan noise melalui GLasso memastikan modal tidak dialokasikan ke aset yang hanya terlihat menguntungkan secara statistik sesaat (spurious divergence), melainkan tetap pada struktur pasar yang kokoh.", options: { italic: true } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 18, color: "333333", valign: "top" });

    // --- Slide 24: Lampiran - Dua Tipe Grid Search ---
    let slide24 = pres.addSlide();
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
    // --- Slide 25: Lampiran - Penanganan Missing Value ---
    let slideOut1 = pres.addSlide();
    slideOut1.addText("Lampiran: Penanganan Data Kosong (Missing Values)", { x: 0.5, y: 0.5, w: "90%", fontSize: 26, bold: true, color: "003366" });
    slideOut1.addText([
        { text: "Pertanyaan: ", options: { bold: true, color: "c0392b" } },
        { text: "\"Beberapa koin seperti Binance/EOS belum rilis di awal 2017 sehingga datanya kosong. Bukankah backward-fill memalsukan harga dan merusak matriks?\"", options: { italic: true, breakLine: true } },
        
        { text: "Jawaban (Kenapa RMT sangat Krusial):", options: { bold: true, breakLine: true, color: "27ae60" } },
        { text: "   1. Praktik pengisian data diam (", options: { } },
        { text: "backward fill", options: { italic: true, bold: true } },
        { text: ") memang menciptakan rentetan nilai harga yang statis.", options: { breakLine: true } },
        { text: "   2. Namun, kehebatan ", options: { } },
        { text: "Random Matrix Theory (RMT)", options: { bold: true } },
        { text: " diuji di sini! Karena data yang datar sama sekali tidak punya korelasi nyata.", options: { breakLine: true } },
        { text: "   3. RMT otomatis akan mendeteksi korelasi buatan tersebut sebagai probabilitas ", options: { } },
        { text: "Noise Acak", options: { bold: true, italic: true } },
        { text: ", lalu membuangnya menjadi 0.", options: { breakLine: true } },
        { text: "   4. Hasilnya, matriks korelasi ", options: { } },
        { text: "terselamatkan", options: { bold: true } },
        { text: " dan tidak tercemar oleh cacat kelengkapan data.", options: { breakLine: true } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 18, color: "333333", valign: "top" });

    // --- Slide 26: Lampiran - Peran USDT (Bagian 1) ---
    let slideOut2 = pres.addSlide();
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
        { text: " sebagai evakuasi risiko.", options: { } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 18, color: "333333", valign: "top" });

    // --- Slide 27: Lampiran - Peran USDT (Bagian 2) ---
    let slideOut3 = pres.addSlide();
    slideOut3.addText("Lampiran: Mengapa Menyertakan Tether (USDT)? (2/2)", { x: 0.5, y: 0.5, w: "90%", fontSize: 26, bold: true, color: "003366" });
    slideOut3.addText([
        { text: "3. Algoritma jaringan ", options: { } },
        { text: "Network Markowitz GS (Grid Search)", options: { bold: true } },
        { text: " dilatih secara matematis; jika mendeteksi korelasi ancaman kolaps merambat ke semua altcoin, ia akan melempar alokasi modalnya menuju node ", options: { } },
        { text: "USDT (Tether)", options: { bold: true } },
        { text: " sebagai langkah ", options: { } },
        { text: "evakuasi otomatis", options: { bold: true } },
        { text: " (Shock-Absorber).", options: { breakLine: true } },
        { text: "   4. Hal ini yang membuat performa Risk-GS (Grid Search) sangat ", options: { } },
        { text: "tangguh", options: { bold: true } },
        { text: " dari serangan Crypto Winter, suatu kapabilitas pertahanan yang tidak dipahami oleh model ortodoks murni Markowitz.", options: { breakLine: true } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 18, color: "333333", valign: "top" });

    // --- Slide 28: Lampiran - Justifikasi Akademik 1: Non-Stationarity ---
    let slideOut4 = pres.addSlide();
    slideOut4.addText("Lampiran: Bukti Empiris Non-Stationarity Pasar", { x: 0.5, y: 0.5, w: "90%", fontSize: 26, bold: true, color: "003366" });
    slideOut4.addText([
        { text: "Masalah: ", options: { bold: true, color: "c0392b" } },
        { text: "Kenapa harus menggunakan model Adaptif (Dynamic Gamma)?", options: { italic: true, breakLine: true } },
        { text: "Bukti dari Grid Search:", options: { bold: true, breakLine: true, color: "27ae60" } },
        { text: "Data menunjukkan nilai Gamma (γ) optimal ", options: { bullet: true } },
        { text: "terus bergeser", options: { bold: true } },
        { text: " setiap periode rebalancing.", options: { breakLine: true } },
        { text: "Penggunaan γ statis tidak cukup untuk menangkap perubahan ", options: { bullet: true } },
        { text: "struktur korelasi", options: { bold: true } },
        { text: " yang sangat cepat di pasar kripto.", options: { breakLine: true } },
        { text: "Ini membenarkan bahwa pasar kripto membutuhkan ", options: { bullet: true } },
        { text: "kalibrasi otomatis", options: { bold: true } },
        { text: " secara temporal.", options: { } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 18, color: "333333", valign: "top" });

    // --- Slide 29: Lampiran - Justifikasi Akademik 2: Strategi Shock-Absorber ---
    let slideOut5 = pres.addSlide();
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
        { text: " daripada sekadar variansi harga.", options: { } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 18, color: "333333", valign: "top" });


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
