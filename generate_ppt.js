const pptxgen = require("pptxgenjs");

async function createPresentation() {
    console.log("Memulai pembuatan presentasi...");
    
    // 1. Inisialisasi presentasi baru
    let pres = new pptxgen();

    // Set layout (opsional, defaulnya 16x9)
    pres.layout = "LAYOUT_16x9";

    // --- Slide 1: Judul ---
    let slide1 = pres.addSlide();
    
    // Menambahkan Judul Utama
    slide1.addText("Proposal Tesis: Optimalisasi Portofolio", { 
        x: 0.5, 
        y: 1.5, 
        w: "90%", 
        fontSize: 36, 
        bold: true, 
        align: "center", 
        color: "2c3e50" 
    });
    
    // Menambahkan Sub-judul
    slide1.addText("Menggunakan Jaringan Saraf Tiruan dan Markowitz", { 
        x: 0.5, 
        y: 2.5, 
        w: "90%", 
        fontSize: 24, 
        align: "center", 
        color: "34495e" 
    });
    
    // Nama Penulis
    slide1.addText("Oleh: [Nama Anda]", { 
        x: 0.5, 
        y: 4, 
        w: "90%", 
        fontSize: 18, 
        align: "center", 
        color: "7f8c8d" 
    });

    // --- Slide 2: Latar Belakang ---
    let slide2 = pres.addSlide();
    
    // Judul Slide
    slide2.addText("Latar Belakang", { 
        x: 0.5, 
        y: 0.5, 
        w: "90%", 
        fontSize: 28, 
        bold: true, 
        color: "2980b9" 
    });
    
    // Konten List (Bullet points)
    slide2.addText([
        { text: "Tingginya volatilitas di pasar cryptocurrency memerlukan manajemen risiko.", options: { bullet: true, breakLine: true } },
        { text: "Kebutuhan terhadap optimasi portofolio yang dinamis sesuai pergerakan pasar.", options: { bullet: true, breakLine: true } },
        { text: "Potensi integrasi analisis jaringan dan Modern Portfolio Theory (Markowitz).", options: { bullet: true } }
    ], { 
        x: 0.5, 
        y: 1.5, 
        w: "80%", 
        h: 3,
        fontSize: 20, 
        color: "333333", 
        valign: "top"
    });

    // --- Slide 3: Metodologi / Alur ---
    let slide3 = pres.addSlide();
    
    slide3.addText("Metodologi Penelitian", { 
        x: 0.5, 
        y: 0.5, 
        w: "90%", 
        fontSize: 28, 
        bold: true, 
        color: "2980b9" 
    });
    
    // Menambahkan bentuk/shape (Kotak Proses)
    slide3.addShape(pres.ShapeType.rect, { 
        x: 1, 
        y: 1.5, 
        w: 3, 
        h: 1.5, 
        fill: { color: "ecf0f1" }, 
        line: { color: "bdc3c7" }
    });
    
    slide3.addText("Pengumpulan Data\n(Yahoo Finance)", { 
        x: 1, 
        y: 1.5, 
        w: 3, 
        h: 1.5,
        fontSize: 16, 
        align: "center", 
        valign: "middle",
        bold: true,
        color: "2c3e50"
    });

    // Panah
    slide3.addShape(pres.ShapeType.rightArrow, { 
        x: 4.2, 
        y: 2, 
        w: 1, 
        h: 0.5, 
        fill: { color: "3498db" } 
    });

    // Kotak Proses 2
    slide3.addShape(pres.ShapeType.rect, { 
        x: 5.4, 
        y: 1.5, 
        w: 3, 
        h: 1.5, 
        fill: { color: "ecf0f1" }, 
        line: { color: "bdc3c7" }
    });

    slide3.addText("Optimasi Portofolio\n(Markowitz Model)", { 
        x: 5.4, 
        y: 1.5, 
        w: 3, 
        h: 1.5,
        fontSize: 16, 
        align: "center", 
        valign: "middle",
        bold: true,
        color: "2c3e50"
    });

    // --- Simpan File ---
    const outputFilename = "Presentasi_Proposal.pptx";
    
    try {
        await pres.writeFile({ fileName: outputFilename });
        console.log(`Bagus! Presentasi berhasil disimpan sebagai: ${outputFilename}`);
    } catch (error) {
        console.error("Terjadi kesalahan saat menyimpan presentasi:", error);
    }
}

// Menjalankan fungsi
createPresentation();
