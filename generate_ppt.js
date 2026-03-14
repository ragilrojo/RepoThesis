const pptxgen = require("pptxgenjs");

async function createPresentation() {
    console.log("Memulai pembuatan presentasi sesuai hasil riset...");
    
    // Inisialisasi presentasi baru
    let pres = new pptxgen();

    // Set layout (opsional, defaulnya 16x9)
    pres.layout = "LAYOUT_16x9";

    // --- Slide 1: Judul ---
    let slide1 = pres.addSlide();
    slide1.addText("Proposal Tesis:\nOptimalisasi Portofolio Adaptif", { 
        x: 0.5, y: 1.2, w: "90%", fontSize: 40, bold: true, align: "center", color: "2c3e50" 
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
    slide2.addText("Latar Belakang", { x: 0.5, y: 0.5, w: "90%", fontSize: 28, bold: true, color: "2980b9" });
    slide2.addText([
        { text: "Tingginya volatilitas ekstrem dan ", options: { bullet: true } },
        { text: "noise", options: { italic: true } },
        { text: " di pasar " },
        { text: "cryptocurrency", options: { italic: true } },
        { text: " memerlukan manajemen rekam jejak risiko yang presisi.", options: { breakLine: true } },

        { text: "Kelemahan Markowitz Tradisional: Rawan terhadap ", options: { bullet: true } },
        { text: "estimation error", options: { italic: true } },
        { text: " pada matriks kovarians, terutama ketika korelasi aset sangat berisik.", options: { breakLine: true } },

        { text: "Potensi ", options: { bullet: true } },
        { text: "Network Markowitz", options: { italic: true } },
        { text: ": Menyaring " },
        { text: "noise", options: { italic: true } },
        { text: " menggunakan " },
        { text: "Random Matrix Theory", options: { italic: true } },
        { text: " (" },
        { text: "RMT", options: { hyperlink: { slide: '9' }, color: "0563C1", underline: true } },
        { text: ") dan memetakan struktur pasar lewat " },
        { text: "Minimum Spanning Tree", options: { italic: true } },
        { text: " (MST) untuk penentuan penalti sentralitas." }
    ], { x: 0.5, y: 1.1, w: "90%", h: 3, fontSize: 20, color: "333333", valign: "top" });

    // --- Slide 3: Konsep "Noise" dalam Cryptocurrency ---
    let slide3 = pres.addSlide();
    slide3.addText("Apa itu \"Noise\" di Pasar Kripto?", { x: 0.5, y: 0.5, w: "90%", fontSize: 28, bold: true, color: "2980b9" });
    slide3.addText([
        { text: "Noise (Kebisingan) Pasar:", options: { bold: true, breakLine: true } },
        { text: "   Fluktuasi harga acak akibat sentimen sesaat, rumor, FOMO, atau spekulasi yang tidak mencerminkan nilai fundamental aset.", options: { breakLine: true } },
        
        { text: "Estimation Error (Korelasi Palsu):", options: { bold: true, breakLine: true } },
        { text: "   Model tradisional seringkali menangkap pergerakan acak ini sebagai korelasi tinggi antar aset, menghasilkan matriks kovarians yang ", options: { } },
        { text: "berisik", options: { italic: true } },
        { text: " dan tidak stabil.", options: { breakLine: true } },

        { text: "Solusi Random Matrix Theory (" },
        { text: "RMT", options: { bold: true, hyperlink: { slide: '9' }, color: "0563C1", underline: true } },
        { text: "):", options: { bold: true, breakLine: true } },
        { text: "   Berfungsi sebagai filter untuk memisahkan korelasi sejati (sinyal struktur pasar) dari sekadar pergerakan kebetulan (", options: { } },
        { text: "noise", options: { italic: true } },
        { text: "), memastikan alokasi portofolio tidak tertipu oleh fluktuasi semu." }
    ], { x: 0.5, y: 1.1, w: "90%", h: 4, fontSize: 18, color: "333333", valign: "top" });

    // --- Slide 4: Landasan Teori ---
    let slide4 = pres.addSlide();
    slide4.addText("Landasan Teori Utama", { x: 0.5, y: 0.5, w: "90%", fontSize: 28, bold: true, color: "2980b9" });
    slide4.addText([
        { text: "Tinjauan Portofolio & Risiko:", options: { bold: true, breakLine: true } },
        { text: "   Portofolio: Kumpulan aset finansial yang dikelola bersama untuk mengoptimalkan profil risiko-imbal hasil melalui diversifikasi.", options: { breakLine: true } },
        { text: "   Volatilitas: Alat statistik pemetaan fluktuasi harga yang menjadi basis penentuan tingkat risiko pasar.", options: { breakLine: true } },
        { text: "   Matriks Kovarians: Alat matematis yang mengukur arah pergerakan bersama (korelasi) serta tingkat fluktuasi seluruh aset dalam portofolio.", options: { breakLine: true } },
        
        { text: "Pendekatan Struktur Jaringan:", options: { bold: true, breakLine: true } },
        { text: "   Random Matrix Theory (" },
        { text: "RMT", options: { hyperlink: { slide: '9' }, color: "0563C1", underline: true } },
        { text: "): Metode fisika statistik pemfilter ", options: { } },
        { text: "noise", options: { italic: true } },
        { text: " untuk merekonstruksi kestabilan matriks korelasi.", options: { breakLine: true } },
        { text: "   Minimum Spanning Tree (MST): Konstruksi jaringan antar aset berdasarkan jarak korelasi terpendek (paling kuat); menyaring informasi redundan tanpa membentuk siklus (loop).", options: { breakLine: true } },
        { text: "   Network Centrality: Metrik penghukuman pada aset pasar yang letaknya terpusat untuk menekan probabilitas risiko penularan letupan harga yang sistemik.", options: { } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5, fontSize: 13, color: "333333", valign: "top" });

    // --- Slide 5: Strategi yang Dibandingkan ---
    let slide5 = pres.addSlide();
    slide5.addText("Strategi Portofolio yang Disimulasikan", { x: 0.5, y: 0.5, w: "90%", fontSize: 28, bold: true, color: "2980b9" });
    slide5.addText([
        { text: "1. EW (" },
        { text: "Equally Weighted", options: { italic: true } },
        { text: "): " },
        { text: "Baseline", options: { italic: true } },
        { text: " naif 1/N.", options: { breakLine: true } },

        { text: "2. CM (" },
        { text: "Classical Markowitz", options: { italic: true } },
        { text: "): " },
        { text: "Mean-Variance Optimization", options: { italic: true } },
        { text: " murni.", options: { breakLine: true } },

        { text: "3. GM (" },
        { text: "Glasso Markowitz", options: { italic: true } },
        { text: "): Regularisasi L1 (" },
        { text: "Graphical Lasso", options: { italic: true } },
        { text: ") pada matriks kovarians.", options: { breakLine: true } },

        { text: "4. NW (" },
        { text: "Network Markowitz", options: { italic: true } },
        { text: ") Statis: Dengan parameter gamma statis (0, 1.0, 2.0).", options: { breakLine: true } },

        { text: "5. NW (" },
        { text: "Grid Search", options: { italic: true } },
        { text: ") Adaptif: Menggunakan " },
        { text: "rolling window", options: { italic: true } },
        { text: " untuk optimasi parameter dinamis." }
    ], { x: 0.5, y: 1.1, w: "90%", h: 3.5, fontSize: 20, color: "333333", valign: "top" });

    // --- Slide 6: Temuan Utama: Performa ---
    let slide6 = pres.addSlide();
    slide6.addText("Temuan Utama: Evaluasi Performa", { x: 0.5, y: 0.5, w: "90%", fontSize: 28, bold: true, color: "2980b9" });
    slide6.addText([
        { text: "Fase ", options: { bold: true, bullet: true } },
        { text: "Bearish & Crypto Winter", options: { bold: true, italic: true } },
        { text: " Ekstrem (2018-2019):", options: { bold: true, breakLine: true } },

        { text: "   " },
        { text: "Network Markowitz", options: { italic: true } },
        { text: " (gamma = 2.0) meredam " },
        { text: "drawdowns", options: { italic: true } },
        { text: " secara signifikan dibandingkan EW dan CM.", options: { breakLine: true } },

        { text: "Ketahanan Risiko (", options: { bold: true, bullet: true } },
        { text: "Value at Risk", options: { bold: true, italic: true } },
        { text: " / VaR):", options: { bold: true, breakLine: true } },

        { text: "   Pembatasan eksposur pada sentralitas tinggi (MST) membantu mengamankan modal dengan VaR sangat stabil.", options: { breakLine: true } },

        { text: "Pengelolaan ", options: { bold: true, bullet: true } },
        { text: "Tail-Risk", options: { bold: true, italic: true } },
        { text: " (", options: { bold: true } },
        { text: "Rachev Ratio", options: { bold: true, italic: true } },
        { text: "):", options: { bold: true, breakLine: true } },

        { text: "   Mengindikasikan " },
        { text: "Network Markowitz", options: { italic: true } },
        { text: " sukses bertindak sebagai " },
        { text: "shock-absorber", options: { italic: true } },
        { text: " dalam meredam " },
        { text: "market crash", options: { italic: true } },
        { text: "." }
    ], { x: 0.5, y: 1.1, w: "90%", h: 4, fontSize: 18, color: "333333", valign: "top" });

    // --- Slide 7: Kesimpulan ---
    let slide7 = pres.addSlide();
    slide7.addText("Kesimpulan", { x: 0.5, y: 0.5, w: "90%", fontSize: 28, bold: true, color: "2980b9" });
    slide7.addText([
        { text: "Superioritas ", options: { bold: true, bullet: true } },
        { text: "Grid Search", options: { bold: true, italic: true } },
        { text: " Adaptif:", options: { bold: true, breakLine: true } },

        { text: "   Strategi NW (Return GS) dan NW (Risk GS) konsisten menjadi " },
        { text: "all-terrain model", options: { italic: true } },
        { text: " dari puncak " },
        { text: "bull", options: { italic: true } },
        { text: " hingga dasar " },
        { text: "winter", options: { italic: true } },
        { text: ".", options: { breakLine: true } },

        { text: "Fleksibilitas Struktur Jaringan:", options: { bold: true, bullet: true, breakLine: true } },

        { text: "   Teori graf terbukti tidak hanya meredam risiko keruntuhan berantai, tetapi mengungguli formulasi statis masa lalu.", options: { breakLine: true } },

        { text: "Potensi Eksploitasi ", options: { bold: true, bullet: true } },
        { text: "Market Recovery", options: { bold: true, italic: true } },
        { text: ":", options: { bold: true, breakLine: true } },

        { text: "   " },
        { text: "Rachev ratio", options: { italic: true } },
        { text: " membuktikan model jaringan memaksimalkan ceruk " },
        { text: "gain", options: { italic: true } },
        { text: " ekstrem saat pasar mulai " },
        { text: "rebound", options: { italic: true } },
        { text: "." }
    ], { x: 0.5, y: 1.1, w: "90%", h: 4, fontSize: 18, color: "333333", valign: "top" });

    // --- Slide 8: Terima Kasih ---
    let slide8 = pres.addSlide();
    slide8.addText("Terima Kasih", { x: 0.5, y: 2.2, w: "90%", fontSize: 40, bold: true, align: "center", color: "2c3e50" });
    slide8.addText("Sesi Tanya Jawab", { x: 0.5, y: 3.2, w: "90%", fontSize: 24, align: "center", color: "7f8c8d" });

    // --- Slide 9: Lampiran - Analogi RMT ---
    let slide9 = pres.addSlide();
    slide9.addText([
        { text: "Lampiran: Analogi " },
        { text: "RMT", options: { hyperlink: { slide: '9' } } },
        { text: " sebagai \"Noise-Canceling\"" }
    ], { x: 0.5, y: 0.5, w: "90%", fontSize: 28, bold: true, color: "2980b9" });
    slide9.addText([
        { text: "Pasar Kripto = Pesta yang Bising:", options: { bold: true, breakLine: true } },
        { text: "   Banyak fluktuasi harga karena sentimen sesaat / kebetulan (", options: { } },
        { text: "noise", options: { italic: true } },
        { text: ").", options: { breakLine: true } },
        
        { text: "Sinyal Korelasi Asli = Suara yang Ingin Didengar:", options: { bold: true, breakLine: true } },
        { text: "   Hubungan nyata antar-aset yang stabil dan berbobot.", options: { breakLine: true } },

        { text: "Random Matrix Theory (" },
        { text: "RMT", options: { bold: true, hyperlink: { slide: '9' } } },
        { text: ") = Headphone Noise-Canceling:", options: { bold: true, breakLine: true } },
        { text: "Membedakan gelombang statistik acak (noise) dari pola suara asli (signal) menggunakan distribusi Marchenko-Pastur.", options: { bullet: true, breakLine: true } },
        { text: "\"Meredam\" spekulasi jangka pendek untuk mencegah ", options: { bullet: true } },
        { text: "estimation error", options: { italic: true } },
        { text: ".", options: { breakLine: true } },
        { text: "Matriks tersisa adalah hubungan antar-aset yang bersih, kuat, & terpercaya.", options: { bullet: true } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5, fontSize: 22, color: "333333", valign: "top" });

    // --- Slide 10: Lampiran - Signal vs Noise ---
    let slide10 = pres.addSlide();
    slide10.addText("Lampiran: Membedakan Hubungan Sejati (Signal) vs Noise", { x: 0.5, y: 0.5, w: "90%", fontSize: 26, bold: true, color: "2980b9" });
    slide10.addText([
        { text: "1. Mencari Nilai Eigen (Eigenvalues):", options: { bold: true, breakLine: true } },
        { text: "   Mengekstrak angka dari matriks korelasi yang mewakili kekuatan pola pergerakan bersama antar-aset.", options: { breakLine: true } },
        
        { text: "2. Batas Noise (Marchenko-Pastur/MP):", options: { bold: true, breakLine: true } },
        { text: "   RMT menghitung batas teoretis maksimum (", options: { } },
        { text: "λ_max", options: { italic: true } },
        { text: ") dari sebuah matriks yang diasumsikan 100% acak tanpa pola.", options: { breakLine: true } },

        { text: "3. Uji Coba Signal vs Noise:", options: { bold: true, breakLine: true } },
        { text: "   • JALUR NOISE: Jika Eigenvalue < λ_max. Dianggap hanya kebetulan acak (hubungan palsu).", options: { breakLine: true } },
        { text: "   • JALUR SIGNAL: Jika Eigenvalue > λ_max. Dianggap memiliki ikatan fundamental (hubungan sejati).", options: { breakLine: true } },

        { text: "4. Pembersihan & Rekonstruksi:", options: { bold: true, breakLine: true } },
        { text: "   Eigenvalues yang tergolong noise dibersihkan (dinolkan) dan hanya nilai signal yang dipertahankan untuk membangun ulang matriks korelasi yang bersih.", options: { breakLine: true } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 4.5, fontSize: 18, color: "333333", valign: "top" });

    // --- Slide 11: Lampiran - Menghitung Nilai Eigen ---
    let slide11 = pres.addSlide();
    slide11.addText("Lampiran: Bagaimana Menghitung Nilai Eigen?", { x: 0.5, y: 0.5, w: "90%", fontSize: 26, bold: true, color: "2980b9" });
    slide11.addText([
        { text: "1. Matriks Korelasi (C):", options: { bold: true, breakLine: true } },
        { text: "   Membentuk matriks (N x N) yang menjabarkan seluruh korelasi pergerakan harga antar sepasang mata uang kripto secara historis.", options: { breakLine: true } },
        
        { text: "2. Konsep Persamaan Karakteristik:", options: { bold: true, breakLine: true } },
        { text: "   Mencari sebuah besaran skalar ", options: { } },
        { text: "λ (lambda/eigenvalue)", options: { italic: true } },
        { text: " dan vektor ", options: { } },
        { text: "v (eigenvector)", options: { italic: true } },
        { text: " yang dapat memenuhi ekuivalensi matriks linier: ", options: { } },
        { text: "C × v = λ × v", options: { bold: true, color: "c0392b", breakLine: true } },

        { text: "3. Solusi Determinan:", options: { bold: true, breakLine: true } },
        { text: "   Secara matematis, nilai λ tersebut adalah akar yang dicari dengan menyelesaikan persamaan determinan: ", options: { } },
        { text: "Det(C - λI) = 0", options: { bold: true, color: "c0392b", breakLine: true } },
        { text: "   (di mana I mewakili matriks Indentitas).", options: { italic: true, breakLine: true } },

        { text: "4. Arti dari Spektrum Hasil:", options: { bold: true, breakLine: true } },
        { text: "   Mesin (seperti metode Eigen-Decomposition) akan menemukan sekumpulan nilai λ yang memuaskan persamaan di atas. Nilai λ yang paling besar mewakili penggerak pasar terbesar (Market Factor).", options: { breakLine: true } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 18, color: "333333", valign: "top" });

    // --- Slide 12: Lampiran - Contoh Praktek (Dummy Data) ---
    let slide12 = pres.addSlide();
    slide12.addText("Lampiran: Praktek Sederhana Menghitung Eigenvalue", { x: 0.5, y: 0.5, w: "90%", fontSize: 26, bold: true, color: "2980b9" });
    slide12.addText([
        { text: "Konteks Dummy: 2 Koin (BTC & ETH) dengan korelasi = 0.5", options: { bold: true, breakLine: true } },
        
        { text: "1. Matriks Korelasi (C):", options: { bold: true, breakLine: true } },
        { text: "   C = [ 1.0  0.5 ]", options: { fontFace: "Courier New", breakLine: true } },
        { text: "       [ 0.5  1.0 ]", options: { fontFace: "Courier New", breakLine: true } },
        
        { text: "2. Persamaan: Det(C - λI) = 0", options: { bold: true, breakLine: true } },
        { text: "   (1 - λ)² - (0.5)² = 0", options: { breakLine: true } },
        { text: "   λ² - 2λ + 0.75 = 0  ", options: { breakLine: true } },
        { text: "   (λ - 1.5)(λ - 0.5) = 0", options: { breakLine: true } },

        { text: "3. Hasil Akar Nilai Eigen:", options: { bold: true, breakLine: true } },
        { text: "   • λ₁ = 1.5 (Market Factor / Signal Kuat)", options: { bold: true, color: "27ae60", breakLine: true } },
        { text: "   • λ₂ = 0.5 (Idiosyncratic Risk / Noise)", options: { bold: true, color: "c0392b", breakLine: true } },

        { text: "Kesimpulan Filtering:", options: { bold: true, breakLine: true } },
        { text: "   Jika RMT mematok batas λ_max = 1.0, maka λ₂ (0.5) akan dianggap sebagai Noise lalu dinolkan, sementara λ₁ (1.5) dijaga sebagai korelasi sejati (Signal).", options: { breakLine: true } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 16, color: "333333", valign: "top" });

    // --- Slide 13: Lampiran - Bagaimana Menghitung Korelasi? ---
    let slide13 = pres.addSlide();
    slide13.addText("Lampiran: Bagaimana Menghitung Korelasi?", { x: 0.5, y: 0.5, w: "90%", fontSize: 26, bold: true, color: "2980b9" });
    slide13.addText([
        { text: "1. Data Historis (Returns):", options: { bold: true, breakLine: true } },
        { text: "   Input dasar berupa data runtut waktu (time-series) dari return harian aset-aset cryptocurrency.", options: { breakLine: true } },
        
        { text: "2. Library & Metode Python:", options: { bold: true, breakLine: true } },
        { text: "   Dalam script ", options: { } },
        { text: "strategy_comparison.ipynb", options: { italic: true } },
        { text: ", korelasi dihitung menggunakan library ", options: { } },
        { text: "Pandas (df.corr())", options: { fontFace: "Courier New", color: "c0392b" } },
        { text: " atau ", options: { } },
        { text: "NumPy (np.corrcoef())", options: { fontFace: "Courier New", color: "c0392b", breakLine: true } },
        { text: "   Secara bawaan (default), fungsi ini menghitung algoritma ", options: { } },
        { text: "Koefisien Korelasi Pearson", options: { bold: true, breakLine: true } },

        { text: "3. Formula Pearson Correlation:", options: { bold: true, breakLine: true } },
        { text: "   ρ(X,Y) = Cov(X,Y) / (σX × σY)", options: { bold: true, color: "27ae60", breakLine: true } },
        { text: "   (Cov = kovarians antara dua koin, σ = standar deviasi tingkat volatilitas).", options: { italic: true, breakLine: true } },

        { text: "4. Output Matriks (N x N):", options: { bold: true, breakLine: true } },
        { text: "   Menghasilkan tabel bersilang berisi nilai antara ", options: { } },
        { text: "-1 (Berkebalikan arah)", options: { bold: true } },
        { text: " hingga ", options: { } },
        { text: "1 (Bergerak searah)", options: { bold: true } },
        { text: ". Nilai diagonal matriks ini selalu 1 (korelasi koin terhadap dirinya sendiri).", options: { breakLine: true } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 16, color: "333333", valign: "top" });

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
