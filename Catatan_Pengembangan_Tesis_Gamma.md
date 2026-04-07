# Catatan Pengembangan Tesis: Analisis Ekstensif Parameter Gamma (γ) dan Hyperparameter Tuning

Dokumen ini disusun sebagai panduan intelektual dan landasan teoritis untuk menjawab pertanyaan-pertanyaan strategis baik saat penyusunan atau saat sidang tesis. Ini merangkum pemahaman tentang parameter **Gamma (Penalti)**, dinamika pemilihannya, metode *tuning*, dan analisis alternatif.

---

## 1. Pemahaman Konsep Dasar: Penalti dan Gamma (γ)

### Apa itu Penalti dalam Network Markowitz?
Penalti adalah "hukuman" matematis yang diberikan kepada suatu aset kripto karena karakteristiknya dalam jaringan. Dalam teori portofolio berbasis jaringan, kita menganalisis **Centrality** (khususnya *Eigenvector Centrality*). Koin yang sangat sentral berarti pergerakannya sangat mempengaruhi (dan dipengaruhi) oleh seluruh koin lain dalam jaringan. Jika terjadi guncangan pasar (*crash*), koin ini adalah penyebar kepanikan utama (*spreader*).

### Apa itu Gamma, Mempengaruhi Apa, dan Dipengaruhi Apa?
*   **Gamma (γ)** adalah koefisien skalar atau *hyperparameter* yang mengatur "kekuatan" dari penalti tersebut.
*   **Mempengaruhi:** Bobot alokasi (persentase dana) akhir sebuah koin. Semakin besar tingkat sentralitas suatu koin dan semakin tinggi nilai Gamma, maka bobot alokasi dana pada koin tersebut akan **semakin dikurangi**.
*   **Dipengaruhi:** Nilai Gamma otomatis dipengaruhi oleh kondisi / rezim pasar (*market regime*). Nilai optimalnya ditentukan oleh fungsi objektif (misal: memaksimalkan *Sharpe Ratio* atau meminimalkan variansi/risiko) dari performa simulasi (*rolling window*).

---

## 2. Dinamika Nilai Gamma (Statis vs Dinamis)

### Apakah Gamma Dinamis?
**Ya, secara operasional portofolio, pendekatan kita membuat nilai Gamma tersebut beradaptasi (dinamis).**  
Karena kita menggunakan skema *Rolling Window* dan *Rebalancing*, kita tidak mematok satu nilai Gamma untuk seluruh rentang waktu (berbeda dengan paper klasik yang statis). Setiap kali jendela waktu bergeser, algoritma mencari kembali nilai Gamma spesifik yang paling optimal khusus untuk kondisi pasar saat itu.

### Varian Nilai Gamma:
*   **γ = 0**: Nol Penalti. Jaringan sama sekali tidak dipertimbangkan. Model kembali menjadi *Classical Markowitz* murni.
*   **γ sedang (misal 0.5 - 1.0)**: Sinergi yang seimbang antara Markowitz dan profil perlindungan risiko jaringan.
*   **γ ekstrem (> 1.0)**: Portofolio berada pada tingkat defensif ultra. Ia sangat ketakutan akan penyebaran risiko (*contagion risk*) dan cenderung membuang altcoin sentral secara ekstrem, lebih memilih koin peripheral/pinggiran (contoh dalam simulasi ini: penempatan ke *safe haven* misal USDT).

---

## 3. Anatomi Rentang Parameter dan Apakah Gamma > 1?

### Mengapa Range-nya seperti itu (misal [0.0 - 2.0])?
Berdasarkan landasan teoretis stabilitas numerik:
1.  Jika Gamma bernilai negatif, itu menghancurkan asumsi rasionalitas portofolio, karena itu artinya sistem malah *menyukai* koin berisiko penularan tinggi.
2.  Jika Gamma terlalu besar (misal > 3 atau > 5), **matriks penalti menjadi "terlalu kuat/mendominasi" matriks fundamental (kovarians)**. Jika matriks penalti mendominasi, optimasi matematikanya akan gagal atau portofolio hanya akan memilih 1 aset non-sentral yang sepenuhnya merusak prinsip diversifikasi *Markowitz*. Range 0 - 2 dianggap secara akademis sebagai rentang penyesuaian (*fine-tuning*) yang logis tanpa merusak integritas matriks asli.

### Apakah Gamma Bisa Lebih dari 1?
Sangat Bisa. Ketika pasar kripto sedang di ambang kehancuran besar (*crypto winter* parah), algoritma akan mengkalkulasi bahwa memegang koin dengan interkoneksi tinggi terlalu berbahaya, sehingga ia membutuhkan gaya pendorong penalti di atas 100% (*multiplier* > 1) agar algoritma secara agresif mereduksi eksposur ke koin utama sentral.

---

## 4. Hyperparameter Tuning: Alternatif Algoritma Selain Grid Search

### Landasan Memilih Grid Search
Kita menggunakan **Grid Search** karena ruang pencariannya hanya berupa 1-Dimensi (hanya variabel Gamma saja) dan rentang pengukurannya sempit (dibagi hingga langkah 0.1). Grid search memastikan **optimasi yang pasti (exhaustive)**; mesin mencoba semua skenario yang mungkin dan mutlak memberikan titik paling optimal di rentang tersebut, dan memiliki transparansi (interpretability) terbesar.

### Mengapa Parameter Ini Perlu di-Tuning?
Dalam disiplin *Machine Learning / Data Science*, *hyperparameter* adalah kontrol atas *learning process*. Jika dibiarkan salah kalibrasi, mesin kehilangan keunggulan (bahkan performanya bisa berada di bawah model *Equally Weighted*).

### Algoritma Pencarian Lebih Lanjut (Alternatif selain Grid Search):
Ada banyak algoritma lain, sangat berguna untuk dicatat di bagian "Saran Pengembangan / Future Works" tesis jika selanjutnya ingin melakukan tuning multivariabel sekaligus (contoh: tuning ukuran *window*, batas nilai *eigen value* RMT, dan nilai Gamma secara berbarengan):
1.  **Bayesian Optimization (contoh: Pustaka `Optuna` atau `Hyperopt`)**: Alih-alih membabi buta mencari satu per satu (seperti Grid Search), Bayesian menggunakan teorema probabilitas *prior* untuk "menebak area" mana dari parameter yang paling mungkin memberikan performa tinggi, lalu memusatkan pencarian ke sana. **Algoritma ini jauh lebih cepat dan cerdas**.
2.  **Randomized Search / Random Search**: Memilih titik parameter secara acak dalam ruang distribusi. Cocok dan jauh lebih murah secara komputasi ketika jumlah dimensi hiperparameternya terlalu banyak melebihi batasan grid.
3.  **Algoritma Genetika (Genetic Algorithm / GA)**: Terinspirasi bio-reproduksi, mengawinsilangkan kumpulan portofolio dengan set parameter acak. Portofolio "unggulan" ber-Sharpe tertinggi dikawinsilangkan hingga mutasi menghasilkan parameter absolut.
4.  **Particle Swarm Optimization (PSO)**: Algoritma kawanan / swarm, mencoba memandu pergerakan arah pencarian parameter melalui iterasi hingga 'partikel' mendekati nilai paling prima.

---

## 5. Analisa Apa yang Bisa Dilakukan dalam Tesis Ini?

Dari semua catatan ini, berikut usulan analisis terstruktur yang dapat diolah di Bab Pembahasan/Hasil Tesis:

1.  **Analisis *Market Regime Trajectory***:
    Petakan grafik pergerakan harga aset / Market Index dengan grafik **fluktuasi nilai Gamma yang dipilih algoritma**. Analisa apakah benar secara empiris, algoritma secara mandiri melonjakkan nilai penalti (Gamma tinggi) ketika akan terjadi periode anjlok (*Crash*), dan menurunkan penalti (Gamma rendah mendekati 0) saat aset sedangreli subur (*Bull Run*) guna memaksimalkan profit? Jika analisis ini terbukti, hal ini menegaskan efektivitas mekanisme adaptif.
2.  **Analisis Sensitivitas Penalti (*Penalty Sensitivity Index*)**:
    Lakukan demonstrasi visual/tabel tentang bagaimana komposisi koin (alokasi modal) bergeser drastis saat membandingkan Gamma = 0 versus Gamma = 1.0 pada tanggal yang sama. Ini menyoroti utilitas langsung bagaimana penalti memaksa disinvestasi dari *hub crypto*.
3.  **Analisis Batas Fundamental (*Threshold Matrix Analysis*)**:
    Mengevaluasi secara spesifik batasan *filter* batas *noise* RMT. Meneliti apakah penggunaan RMT justru menjaga struktur matriks adjusment aman selama Grid Search (agar hasil penalti grid tidak membuat matriks gagal inversi/melanggar kelayakan non-singularity dalam optimasi aljabar).
