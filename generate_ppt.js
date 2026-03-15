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

    // --- Slide 4.5: Kerangka Penelitian ---
    let slideFramework = pres.addSlide();
    slideFramework.addText("Kerangka Pemikiran / Penelitian", { x: 0.5, y: 0.5, w: "90%", fontSize: 28, bold: true, color: "2980b9" });
    slideFramework.addImage({ path: "e:/ProjectNodeJs/temp_doc_build/framwrok.jpg", x: 1.0, y: 1.1, w: 8.0, h: 4.0 });

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

    // --- Slide 6: Matriks Evaluasi Performa ---
    let slide6 = pres.addSlide();
    slide6.addText("Matriks Evaluasi Performa", { x: 0.5, y: 0.5, w: "90%", fontSize: 28, bold: true, color: "2980b9" });
    slide6.addText([
        { text: "1. Sharpe Ratio (Risk-Adjusted Return):", options: { bold: true, breakLine: true, color: "27ae60" } },
        { text: "   Mengukur imbal hasil berlebih per unit risiko (Volatilitas) secara umum.", options: { breakLine: true } },
        { text: "   Target:", options: { bold: true } },
        { text: " Semakin besar nilainya semakin efisien kualitas portofolio tersebut.", options: { breakLine: true } },

        { text: "2. Value at Risk (VaR) / Downside Risk:", options: { bold: true, breakLine: true, color: "c0392b" } },
        { text: "   Mengkuantifikasi potensi kemungkinan batas 'Kerugian Maksimal' yang dapat diderita model (kondisi crash).", options: { breakLine: true } },
        { text: "   Target:", options: { bold: true } },
        { text: " Semakin kecil batas toleransi kerugiannya (mendekati 0), tandanya model sukses menjadi perisai.", options: { breakLine: true } },

        { text: "3. Rachev Ratio (Tail Risk & Reward):", options: { bold: true, breakLine: true, color: "8e44ad" } },
        { text: "   Karena kripto sering meledak tinggi/rendah secara instan (", options: { } },
        { text: "fat-tail", options: { italic: true } },
        { text: "), metrik ini secara spesifik membandingkan kuantil ekor ekstrem: Potensi 'Profit Ekstrem' vs Ancaman 'Loss Ekstrem'.", options: { breakLine: true } },
        { text: "   Target:", options: { bold: true } },
        { text: " Jika nilainya positif besar, artinya peluang profit jauh menutupi probabilitas loss.", options: { breakLine: true } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.0, fontSize: 18, color: "333333", valign: "top" });
    // --- Slide 7: Terima Kasih ---
    let slide7 = pres.addSlide();
    slide7.addText("Terima Kasih", { x: 0.5, y: 2.7, w: "90%", fontSize: 40, bold: true, align: "center", color: "2c3e50" });

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

    // --- Slide 14: Lampiran - Apakah Nilai Eigen Statis? ---
    let slide14 = pres.addSlide();
    slide14.addText("Lampiran: Apakah Nilai Eigen Statis?", { x: 0.5, y: 0.5, w: "90%", fontSize: 26, bold: true, color: "2980b9" });
    slide14.addText([
        { text: "Apakah Nilai Eigen Sudah Ditentukan (Statis)?", options: { bold: true, breakLine: true, color: "c0392b" } },
        { text: "   TIDAK. Nilai Eigen (Eigenvalue) adalah hasil perhitungan dinamis yang diekstrak langsung dari matriks korelasi aset-aset pada saat itu.", options: { breakLine: true } },
        
        { text: "Proses Dinamis dalam Framework Network Markowitz:", options: { bold: true, breakLine: true } },
        { text: "1. Matriks Korelasi Berubah:", options: { bold: true } },
        { text: " Dihitung ulang dari data pergerakan harga terbaru (rolling window).", options: { breakLine: true } },
        
        { text: "2. Dekomposisi Eigen Diperbarui:", options: { bold: true } },
        { text: " Matriks korelasi baru dipecah menjadi Nilai Eigen (kekuatan pola) dan Vektor Eigen (arah).", options: { breakLine: true } },
        
        { text: "3. Batas Filter MP Ikut Berubah:", options: { bold: true } },
        { text: " Batas noise (Marchenko-Pastur) juga dihitung ulang mengikuti rasio jumlah data harian dibagi jumlah aset.", options: { breakLine: true } },
        
        { text: "4. Korelasi Bersih Terbentuk:", options: { bold: true } },
        { text: " Nilai eigen yang masuk kategori noise dinolkan, dan nilai eigen sinyal (signal) digunakan untuk merekonstruksi korelasi yang stabil.", options: { breakLine: true } },
        
        { text: "→ Kesimpulan: Sifat adaptif secara ", options: { bold: true, color: "27ae60" } },
        { text: "real-time", options: { bold: true, italic: true, color: "27ae60" } },
        { text: " inilah yang membuat model sanggup merespons dengan cepat perubahan rezim dari bull ke bear market.", options: { bold: true, color: "27ae60", breakLine: true } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 16, color: "333333", valign: "top" });

    // --- Slide 15: Lampiran - Batas Noise Marchenko-Pastur ---
    let slide15 = pres.addSlide();
    slide15.addText("Lampiran: Menentukan Batas Noise (Marchenko-Pastur)", { x: 0.5, y: 0.5, w: "90%", fontSize: 26, bold: true, color: "2980b9" });
    slide15.addText([
        { text: "Apa itu Batas Marchenko-Pastur (MP)?", options: { bold: true, breakLine: true } },
        { text: "   Distribusi MP adalah teori probabilitas yang memprediksi seperti apa bentuk (distribusi eigenvalue) dari sebuah matriks yang 100% berisi angka acak (noise).", options: { breakLine: true } },
        
        { text: "Menghitung Batas Atas Noise (λ_max):", options: { bold: true, breakLine: true } },
        { text: "   λ_max = 1 + (1/Q) + 2√(1/Q)", options: { bold: true, color: "c0392b", breakLine: true } },
        
        { text: "Apa itu Rasio Q?", options: { bold: true, breakLine: true } },
        { text: "   ", options: { } },
        { text: "Q = T / N", options: { bold: true, color: "27ae60", breakLine: true } },
        { text: "   • T = Jumlah baris data historis (Misal: pergerakan harga selama 365 hari)", options: { bullet: true, breakLine: true } },
        { text: "   • N = Jumlah kolom / aset / koin (Misal: 10 koin kripto)", options: { bullet: true, breakLine: true } },
        
        { text: "Mekanisme Pemfilteran:", options: { bold: true, breakLine: true } },
        { text: "   Setiap nilai eigen (eigenvalue) dari matriks korelasi kripto akan dibandingkan dengan λ_max ini.", options: { breakLine: true } },
        { text: "   • Eigenvalue < λ_max : Dihapus (Dianggap Noise/Acak)", options: { bold: true, color: "7f8c8d", breakLine: true } },
        { text: "   • Eigenvalue > λ_max : Dipertahankan (Dianggap Sinyal Fundamental)", options: { bold: true, color: "27ae60", breakLine: true } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 16, color: "333333", valign: "top" });

    // --- Slide 16: Lampiran - Analogi Minimum Spanning Tree (Bagian 1) ---
    let slide16 = pres.addSlide();
    slide16.addText("Lampiran: Analogi Minimum Spanning Tree (MST)", { x: 0.5, y: 0.5, w: "90%", fontSize: 26, bold: true, color: "2980b9" });
    slide16.addText([
        { text: "Membangun Jaringan Jalan Tol Antar Kota:", options: { bold: true, breakLine: true } },
        { text: "   Bayangkan aset-aset kripto (BTC, ETH, BNB) adalah kota-kota yang ingin dihubungkan dengan jalan tol (korelasi).", options: { breakLine: true } },
        
        { text: "Aturan MST (Minimum Spanning Tree):", options: { bold: true, breakLine: true } },
        { text: "1. Hubungkan Kota yang Paling Dekat Dulu:", options: { bold: true, color: "27ae60" } },
        { text: " (Memprioritaskan korelasi yang paling kuat antar aset).", options: { breakLine: true } },
        
        { text: "2. Jangkau Semua Kota:", options: { bold: true, color: "27ae60" } },
        { text: " (Semua koin dalam portofolio harus terhubung dalam 1 jaringan yang sama).", options: { breakLine: true } },
        
        { text: "3. Dilarang Membuat Jalan Memutar (Tanpa Loop):", options: { bold: true, color: "c0392b" } },
        { text: " (Agar tidak ada informasi pergerakan harga yang bergema/berputar-putar secara redundan yang bisa memicu reaksi berlebihan).", options: { breakLine: true } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 18, color: "333333", valign: "top" });

    // --- Slide 17: Lampiran - Analogi Minimum Spanning Tree (Bagian 2) ---
    let slide17 = pres.addSlide();
    slide17.addText("Lampiran: Mengapa Kita Membutuhkan MST?", { x: 0.5, y: 0.5, w: "90%", fontSize: 26, bold: true, color: "2980b9" });
    slide17.addText([
        { text: "Menemukan Titik Kemacetan (Hub Centrality):", options: { bold: true, color: "8e44ad", breakLine: true } },
        { text: "   • Di sebuah jaringan jalan tol, pasti ada 1 kota besar yang menjadi pusat persimpangan (Banyak jalan terhubung ke sana).", options: { breakLine: true } },
        
        { text: "   • Di pasar kripto, kota pusat ini mewakili koin yang ", options: { breakLine: true } },
        { text: "Sangat Sentral", options: { bold: true, italic: true } },
        { text: ". Jika terjadi \"kecelakaan\" (harga anjlok) di koin sentral ini, efeknya akan langsung menular ke seluruh jaringan.", options: { breakLine: true } },
        
        { text: "Solusi Network Markowitz:", options: { bold: true, breakLine: true } },
        { text: "   Koin sentral/pusat ini akan diberi hukuman (penalti bobot alokasi) agar portofolio tidak hancur seketika saat koin tersebut crash.", options: { breakLine: true } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 18, color: "333333", valign: "top" });

    // --- Slide 18: Lampiran - Penalti (Gamma) Optimal ---
    let slide18 = pres.addSlide();
    slide18.addText("Lampiran: Berapa Nilai \"Penalti\" (Gamma) yang Optimal?", { x: 0.5, y: 0.5, w: "90%", fontSize: 26, bold: true, color: "2980b9" });
    slide18.addText([
        { text: "Apakah Ada Satu Angka Penalti yang Sempurna?", options: { bold: true, breakLine: true, color: "c0392b" } },
        { text: "   TIDAK. Secara teori, tidak ada nilai penalti statis yang selalu cocok sepanjang waktu akibat perubahan siklus (rezim) pada fluktuasi uang kripto.", options: { breakLine: true } },
        
        { text: "1. Ketika Fase Bull Market (Pasar Menguat):", options: { bold: true, breakLine: true, color: "27ae60" } },
        { text: "   Koin-koin cenderung naik bersamaan. Di fase ini, jika koin sentral diberi penalti terlalu berat (misal $\\gamma$ = 2.0), Anda berisiko kehilangan peluang mendulang untung yang besar.", options: { breakLine: true } },
        { text: "   • ", options: { } },
        { text: "Nilai Optimal:", options: { bold: true } },
        { text: " Cenderung rendah (mulai mendekati 0).", options: { breakLine: true } },
        
        { text: "2. Ketika Fase Bear Market / Crash (Pasar Jatuh):", options: { bold: true, breakLine: true, color: "c0392b" } },
        { text: "   Kepanikan massal membuat koreksi harga saling menular. Keruntuhan 1 koin sentral bisa mematikan seluruh portofolio Anda secara seketika (", options: { } },
        { text: "Tail-Risk", options: { italic: true } },
        { text: ").", options: { breakLine: true } },
        { text: "   • ", options: { } },
        { text: "Nilai Optimal:", options: { bold: true } },
        { text: " Cenderung tinggi (bisa $\\gamma$ = 1.0, 2.0, dst.) untuk melumpuhkan bobot eksposur koin yang krusial tersebut.", options: { breakLine: true } },
        
        { text: "Kesimpulan Strategi Grid Search (GS):", options: { bold: true, breakLine: true, color: "8e44ad" } },
        { text: "   Alih-alih menebak satu tebakan buta, Network Markowitz GS membiarkan komputer \"terus belajar dan menyesuaikan\" nilai Gamma yang paling sesuai dengan data harga harian terbaru.", options: { breakLine: true } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 16, color: "333333", valign: "top" });

    // --- Slide 19: Lampiran - Classical Markowitz (Bagian 1) ---
    let slide19 = pres.addSlide();
    slide19.addText("Lampiran: Apa itu Classical Markowitz (CM)? (1/2)", { x: 0.5, y: 0.5, w: "90%", fontSize: 26, bold: true, color: "2980b9" });
    slide19.addText([
        { text: "Modern Portfolio Theory (MPT) / Mean-Variance Optimization:", options: { bold: true, breakLine: true } },
        { text: "   Merupakan teori klasik (ditemukan oleh Harry Markowitz tahun 1952) yang mencoba meramu komposisi/bobot aset dalam portofolio dengan tujuan matematika murni:", options: { breakLine: true } },
        { text: "   • Memaksimalkan tingkat keuntungan (Return) pada tingkat risiko tertentu, ATAU", options: { breakLine: true } },
        { text: "   • Meminimalkan risiko (Variance) pada tingkat keuntungan tertentu.", options: { breakLine: true } },
        
        { text: "Asumsi Dasar Classical Markowitz:", options: { bold: true, breakLine: true } },
        { text: "   • Investor diasumsikan sepenuhnya rasional dan benci risiko (", options: { } },
        { text: "Risk-averse", options: { italic: true } },
        { text: ").", options: { breakLine: true } },
        { text: "   • Model ini sangat bergantung pada matriks kovarians historis sebagai pedoman utama memprediksi masa depan.", options: { breakLine: true } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 18, color: "333333", valign: "top" });

    // --- Slide 20: Lampiran - Classical Markowitz (Bagian 2) ---
    let slide20 = pres.addSlide();
    slide20.addText("Lampiran: Mengapa Classical Markowitz Kesulitan? (2/2)", { x: 0.5, y: 0.5, w: "90%", fontSize: 26, bold: true, color: "2980b9" });
    slide20.addText([
        { text: "Kelemahan Klasik di Pasar Kripto (Estimation Error):", options: { bold: true, breakLine: true, color: "c0392b" } },
        { text: "   1. Pasar kripto sangat liar (", options: { } },
        { text: "hyper-volatile", options: { italic: true } },
        { text: ") dan memiliki korelasi ekor tebal. Fluktuasi historis belum tentu berulang di masa depan.", options: { breakLine: true } },
        { text: "   2. CM memakan mentah-mentah noise (angka semu/acak) tanpa memfilternya, yang berujung pada portofolio gagal yang terlalu percaya diri pada korelasi historis palsu.", options: { breakLine: true } },

        { text: "Evolusi → Network Markowitz:", options: { bold: true, breakLine: true, color: "27ae60" } },
        { text: "   Oleh karena itu, CM dikembangkan menjadi Network Markowitz di penelitian ini: membersihkan noise (RMT) dan menghukum koin dominan yang rawan hancur (MST).", options: { breakLine: true } }
    ], { x: 0.5, y: 1.1, w: "90%", h: 5.5, fontSize: 18, color: "333333", valign: "top" });

    // --- Slide 21: Lampiran - Dua Tipe Grid Search ---
    let slide22 = pres.addSlide();
    slide22.addText("Lampiran: Dua Tipe Pendekatan Grid Search", { x: 0.5, y: 0.5, w: "90%", fontSize: 26, bold: true, color: "2980b9" });
    slide22.addText([
        { text: "Dua Objektif Optimasi (Return vs Risk):", options: { bold: true, breakLine: true } },
        { text: "   Dalam penelitian ini, Grid Search dibelah menjadi 2 pendekatan utama agar sejalan dengan tujuan dari masing-masing investor (Mau untung besar vs Cari aman).", options: { breakLine: true } },
        
        { text: "1. Network Markowitz dengan Target Return (NW - Return GS):", options: { bold: true, breakLine: true, color: "27ae60" } },
        { text: "   • ", options: { } },
        { text: "Tujuan:", options: { bold: true } },
        { text: " Memaksimalkan capaian tingkat imbal hasil (", options: { } },
        { text: "Expected Return", options: { italic: true } },
        { text: ") portofolio.", options: { breakLine: true } },
        { text: "   • ", options: { } },
        { text: "Sifat:", options: { bold: true } },
        { text: " Lebih Agresif. Grid search akan mencari kombinasi penalti (\u03b3) dan alokasi aset yang bisa menyerok keuntungan sebesar mungkin, sangat cocok untuk mengeksploitasi reli harga saat pasar ", options: { } },
        { text: "Bullish / Recovery", options: { italic: true } },
        { text: ".", options: { breakLine: true } },
        
        { text: "2. Network Markowitz dengan Target Risiko (NW - Risk GS):", options: { bold: true, breakLine: true, color: "c0392b" } },
        { text: "   • ", options: { } },
        { text: "Tujuan:", options: { bold: true } },
        { text: " Menekan parameter risiko total portofolio (", options: { } },
        { text: "Variance", options: { italic: true } },
        { text: ") hingga ke tingkat minimal.", options: { breakLine: true } },
        { text: "   • ", options: { } },
        { text: "Sifat:", options: { bold: true } },
        { text: " Lebih Defensif. Algoritma akan mencari nilai \u03b3 tinggi yang paling efektif meredam fluktuasi harga dan mengamankan modal (Mengerem penyebaran efek ", options: { } },
        { text: "Contagion", options: { italic: true } },
        { text: ") saat terjadi crash di fase ", options: { } },
        { text: "Crypto Winter", options: { italic: true } },
        { text: ".", options: { breakLine: true } }
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
