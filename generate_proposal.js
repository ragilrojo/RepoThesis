const docx = require('docx');
const {
    Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
    AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
    LevelFormat, PageNumber, PageBreak, TabStopType, TabStopPosition,
    VerticalAlign, ImageRun, Footer, NumberFormat
} = docx;

// Fallback for TabLeader which might be named differently or missing in some versions
const TabLeader = docx.TabLeader || docx.TabStopLeader || { DOT: "dot" };
const fs = require('fs');

const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };

function p(text, opts = {}) {
    return new Paragraph({
        ...opts,
        children: [new TextRun({ text, font: "Times New Roman", size: 24, ...opts.run })]
    });
}

function heading1(text) {
    return new Paragraph({
        heading: HeadingLevel.HEADING_1,
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text, font: "Times New Roman", size: 28, bold: true })]
    });
}

function heading2(text) {
    return new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun({ text, font: "Times New Roman", size: 26, bold: true })]
    });
}

function heading3(text) {
    return new Paragraph({
        heading: HeadingLevel.HEADING_3,
        children: [new TextRun({ text, font: "Times New Roman", size: 24, bold: true })]
    });
}

function body(text, opts = {}) {
    return new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { before: 0, after: 120, line: 360, lineRule: "auto" },
        indent: { firstLine: 720 },
        children: [new TextRun({ text, font: "Times New Roman", size: 24, ...opts })]
    });
}

function bodyNoIndent(text, opts = {}) {
    return new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { before: 0, after: 120, line: 360, lineRule: "auto" },
        children: [new TextRun({ text, font: "Times New Roman", size: 24, ...opts })]
    });
}

// Paragraf campuran italic dan non-italic
function mixedBody(segments) {
    return new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { before: 0, after: 120, line: 360, lineRule: "auto" },
        indent: { firstLine: 720 },
        children: segments.map(s => new TextRun({ text: s.text, font: "Times New Roman", size: 24, italics: s.italic || false }))
    });
}

function emptyLine() {
    return new Paragraph({ children: [new TextRun("")] });
}

function centeredBold(text, size = 24) {
    return new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 120 },
        children: [new TextRun({ text, font: "Times New Roman", size, bold: true })]
    });
}

function centered(text, size = 24) {
    return new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 120 },
        children: [new TextRun({ text, font: "Times New Roman", size })]
    });
}

// Numbered list item
function numItem(text, ref = "numbers") {
    return new Paragraph({
        numbering: { reference: ref, level: 0 },
        alignment: AlignmentType.JUSTIFIED,
        spacing: { before: 0, after: 80, line: 360, lineRule: "auto" },
        children: [new TextRun({ text, font: "Times New Roman", size: 24 })]
    });
}

// Letter list item (a, b, c)
function letterItem(text, ref = "letters") {
    return new Paragraph({
        numbering: { reference: ref, level: 0 },
        alignment: AlignmentType.JUSTIFIED,
        spacing: { before: 0, after: 80, line: 360, lineRule: "auto" },
        children: [new TextRun({ text, font: "Times New Roman", size: 24 })]
    });
}

function bulletItem(text) {
    return new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        spacing: { before: 0, after: 80 },
        children: [new TextRun({ text, font: "Times New Roman", size: 24 })]
    });
}

function tocRow(text, page, indentLevel = 0, isBold = false) {
    return new Paragraph({
        spacing: { before: 80, after: 80 },
        indent: { left: indentLevel * 400 },
        tabStops: [
            { type: TabStopType.RIGHT, position: 9000, leader: TabLeader.DOT }
        ],
        children: [
            new TextRun({ text, font: "Times New Roman", size: 24, bold: isBold }),
            new TextRun({ text: "\t", font: "Times New Roman", size: 24, bold: isBold }),
            new TextRun({ text: page, font: "Times New Roman", size: 24, bold: isBold }),
        ],
    });
}

function tocChapter(num, title, page) {
    return new Paragraph({
        spacing: { before: 120, after: 120 },
        tabStops: [
            { type: TabStopType.LEFT, position: 500 },
            { type: TabStopType.RIGHT, position: 9000, leader: TabLeader.DOT }
        ],
        children: [
            new TextRun({ text: num + "\t" + title, font: "Times New Roman", size: 24, bold: true }),
            new TextRun({ text: "\t", font: "Times New Roman", size: 24, bold: true }),
            new TextRun({ text: page, font: "Times New Roman", size: 24, bold: true }),
        ],
    });
}

const doc = new Document({
    numbering: {
        config: [
            {
                reference: "numbers",
                levels: [{
                    level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
                    style: { paragraph: { indent: { left: 720, hanging: 360 } } }
                }]
            },
            {
                reference: "letters",
                levels: [{
                    level: 0, format: LevelFormat.LOWER_LETTER, text: "%1.", alignment: AlignmentType.LEFT,
                    style: { paragraph: { indent: { left: 720, hanging: 360 } } }
                }]
            },
            {
                reference: "letters1",
                levels: [{
                    level: 0, format: LevelFormat.LOWER_LETTER, text: "%1.", alignment: AlignmentType.LEFT,
                    style: { paragraph: { indent: { left: 720, hanging: 360 } } }
                }]
            },
            {
                reference: "letters2",
                levels: [{
                    level: 0, format: LevelFormat.LOWER_LETTER, text: "%1.", alignment: AlignmentType.LEFT,
                    style: { paragraph: { indent: { left: 720, hanging: 360 } } }
                }]
            },
            {
                reference: "letters3",
                levels: [{
                    level: 0, format: LevelFormat.LOWER_LETTER, text: "%1.", alignment: AlignmentType.LEFT,
                    style: { paragraph: { indent: { left: 720, hanging: 360 } } }
                }]
            },
            {
                reference: "bullets",
                levels: [{
                    level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
                    style: { paragraph: { indent: { left: 720, hanging: 360 } } }
                }]
            }
        ]
    },
    styles: {
        default: {
            document: { run: { font: "Times New Roman", size: 24 } }
        },
        paragraphStyles: [
            {
                id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
                run: { size: 28, bold: true, font: "Times New Roman" },
                paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 0 }
            },
            {
                id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
                run: { size: 26, bold: true, font: "Times New Roman" },
                paragraph: { spacing: { before: 200, after: 100 }, outlineLevel: 1 }
            },
            {
                id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
                run: { size: 24, bold: true, font: "Times New Roman" },
                paragraph: { spacing: { before: 160, after: 80 }, outlineLevel: 2 }
            }
        ]
    },
    sections: [
        // ==================== COVER PAGE ====================
        {
            properties: {
                page: {
                    size: { width: 11906, height: 16838 },
                    margin: { top: 1440, right: 1440, bottom: 1440, left: 1800 }
                }
            },
            children: [
                emptyLine(),
                centeredBold("OPTIMASI DINAMIS PEMODELAN NETWORK MARKOWITZ", 28),
                centeredBold("UNTUK MANAJEMEN PORTOFOLIO MATA UANG KRIPTO", 28),
                emptyLine(),
                emptyLine(),
                emptyLine(),
                centered("[Logo Universitas Nusa Mandiri]", 20),
                emptyLine(),
                emptyLine(),
                centeredBold("PROPOSAL TESIS", 26),
                emptyLine(),
                centered("Diajukan sebagai salah satu syarat untuk memperoleh gelar"),
                centered("Magister Komputer (M.Kom)"),
                emptyLine(),
                emptyLine(),
                centeredBold("Ragil Yulianto", 24),
                centered("14240007"),
                emptyLine(),
                emptyLine(),
                centered("Program Studi Ilmu Komputer (S2)"),
                centered("Fakultas Teknologi Informasi"),
                centered("Universitas Nusa Mandiri"),
                centered("Jakarta"),
                centered("2026"),
            ]
        },
        // ==================== HALAMAN PENGESAHAN ====================
        {
            properties: {
                page: {
                    size: { width: 11906, height: 16838 },
                    margin: { top: 1440, right: 1440, bottom: 1440, left: 1800 }
                },
                pageNumber: { start: 2, formatType: NumberFormat.LOWER_ROMAN }
            },
            footers: {
                default: new Footer({
                    children: [
                        new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [new TextRun({ text: "ii", font: "Times New Roman", size: 24 })]
                        })
                    ]
                })
            },
            children: [
                emptyLine(),
                centeredBold("HALAMAN PENGESAHAN", 26),
                emptyLine(),
                centeredBold("PROPOSAL TESIS", 24),
                emptyLine(),
                emptyLine(),
                bodyNoIndent("Proposal tesis ini diajukan oleh:", { bold: true }),
                emptyLine(),
                // Tabel transparan 3 kolom: Label | : | Nilai
                (() => {
                    const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
                    const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };
                    const mkCell = (colW, text) => new TableCell({
                        borders: noBorders,
                        width: { size: colW, type: WidthType.DXA },
                        margins: { top: 40, bottom: 40, left: 0, right: 80 },
                        children: [new Paragraph({ children: [new TextRun({ text, font: "Times New Roman", size: 24, bold: true })] })]
                    });
                    const rows = [
                        ["Nama",          "Ragil Yulianto"],
                        ["NIM",           "14240007"],
                        ["Program Studi", "Ilmu Komputer"],
                        ["Fakultas",      "Teknologi Informasi"],
                        ["Jenjang",       "Strata Dua (S2)"],
                        ["Judul Tesis",   "OPTIMASI DINAMIS PEMODELAN NETWORK MARKOWITZ UNTUK MANAJEMEN PORTOFOLIO MATA UANG KRIPTO"],
                    ];
                    return new Table({
                        width: { size: 8666, type: WidthType.DXA },
                        columnWidths: [2100, 300, 6266],
                        rows: rows.map(([label, value]) => new TableRow({
                            children: [
                                mkCell(2100, label),
                                mkCell(300,  ":"),
                                mkCell(6266, value),
                            ]
                        }))
                    });
                })(),
                emptyLine(),
                emptyLine(),
                bodyNoIndent("Jakarta, 11 Maret 2026"),
                emptyLine(),
                emptyLine(),
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [new TextRun({ text: "Dosen Pembimbing", font: "Times New Roman", size: 24, bold: true })]
                }),
                emptyLine(),
                emptyLine(),
                emptyLine(),
                emptyLine(),
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                        new TextRun({ text: "____________________________________", font: "Times New Roman", size: 24, bold: true }),
                    ]
                }),
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                        new TextRun({ text: "Nama Dosen Pembimbing", font: "Times New Roman", size: 24, bold: true }),
                    ]
                }),
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                        new TextRun({ text: "NIDN. XXXXXXXXXX", font: "Times New Roman", size: 24, bold: true }),
                    ]
                }),
                emptyLine(),
                emptyLine(),
            ]
        },
        // ==================== DAFTAR ISI ====================
        {
            properties: {
                page: {
                    size: { width: 11906, height: 16838 },
                    margin: { top: 1440, right: 1440, bottom: 1440, left: 1800 }
                },
                pageNumber: { formatType: NumberFormat.LOWER_ROMAN }
            },
            footers: {
                default: new Footer({
                    children: [
                        new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [new TextRun({ text: "iii", font: "Times New Roman", size: 24 })]
                        })
                    ]
                })
            },
            children: [
                emptyLine(),
                centeredBold("DAFTAR ISI", 26),
                emptyLine(),
                emptyLine(),
                tocRow("HALAMAN PENGESAHAN", "ii", 0, true),
                emptyLine(),
                tocChapter("I", "PENDAHULUAN", "1"),
                tocRow("1.1 Latar Belakang", "1", 1),
                tocRow("1.2 Identifikasi Masalah", "1", 1),
                tocRow("1.3 Tujuan Penelitian", "2", 1),
                tocRow("1.4 Ruang Lingkup Penelitian", "2", 1),
                tocRow("1.5 Sistematika Penulisan", "3", 1),
                emptyLine(),
                tocChapter("II", "LANDASAN/KERANGKA PEMIKIRAN", "4"),
                tocRow("2.1 Kerangka Teori", "4", 1),
                tocRow("2.1.1 Modern Portfolio Theory (Markowitz)", "4", 2),
                tocRow("2.1.2 Random Matrix Theory (RMT) dan Kompleksitas Jaringan", "4", 2),
                tocRow("2.1.3 Network Markowitz", "4", 2),
                tocRow("2.1.4 Penelitian Terdahulu", "4", 2),
                emptyLine(),
                tocChapter("III", "METODOLOGI PENELITIAN", "6"),
                tocRow("3.1 Tahapan Penelitian", "6", 1),
                tocRow("3.2 Alat dan Bahan Penelitian", "7", 1),
                tocRow("3.2.1 Perangkat Lunak", "7", 2),
                tocRow("3.3 Dataset", "7", 1),
                tocRow("3.4 Metode/Algoritma yang Digunakan", "7", 1),
                tocRow("3.5 Rencana Jadwal Penelitian", "8", 1),
                emptyLine(),
                tocRow("DAFTAR REFERENSI", "9", 0, true),
                emptyLine(),
                emptyLine(),
            ]
        },
        // ==================== BAB I ====================
        {
            properties: {
                page: {
                    size: { width: 11906, height: 16838 },
                    margin: { top: 1440, right: 1440, bottom: 1440, left: 1800 }
                },
                pageNumber: { start: 1, formatType: NumberFormat.DECIMAL }
            },
            footers: {
                default: new Footer({
                    children: [
                        new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [new TextRun({ children: [PageNumber.CURRENT], font: "Times New Roman", size: 24 })]
                        })
                    ]
                })
            },
            children: [
                centeredBold("BAB I", 26),
                centeredBold("PENDAHULUAN", 26),
                emptyLine(),
                heading2("1.1. Latar Belakang"),
                new Paragraph({
                    alignment: AlignmentType.JUSTIFIED,
                    spacing: { before: 0, after: 120, line: 360, lineRule: "auto" },
                    indent: { firstLine: 720 },
                    children: [
                        new TextRun({ text: "Mata uang kripto (", font: "Times New Roman", size: 24 }),
                        new TextRun({ text: "cryptocurrency", font: "Times New Roman", size: 24, italics: true }),
                        new TextRun({ text: ") telah berkembang menjadi salah satu aset investasi digital yang sangat diminati namun memiliki tingkat volatilitas ekstrem. Dalam manajemen portofolio tradisional, model Mean-Variance dari Markowitz kerap digunakan untuk mengalokasikan aset demi mencapai kombinasi return dan risiko yang optimal. Sayangnya, model klasik ini sangat rentan terhadap ", font: "Times New Roman", size: 24 }),
                        new TextRun({ text: "noise", font: "Times New Roman", size: 24, italics: true }),
                        new TextRun({ text: " dan estimasi matriks korelasi yang tidak stabil, terutama pada saat gejolak pasar (", font: "Times New Roman", size: 24 }),
                        new TextRun({ text: "market crash", font: "Times New Roman", size: 24, italics: true }),
                        new TextRun({ text: ") seperti fenomena ", font: "Times New Roman", size: 24 }),
                        new TextRun({ text: "crypto winter", font: "Times New Roman", size: 24, italics: true }),
                        new TextRun({ text: " [1].", font: "Times New Roman", size: 24 }),
                    ]
                }),
                new Paragraph({
                    alignment: AlignmentType.JUSTIFIED,
                    spacing: { before: 0, after: 120, line: 360, lineRule: "auto" },
                    indent: { firstLine: 720 },
                    children: [
                        new TextRun({ text: "Ketika terjadi guncangan pasar, sebagian besar aset kripto cenderung jatuh secara bersamaan, merusak struktur korelasi normal dan menyebabkan portofolio standar mengalami kerugian parah (", font: "Times New Roman", size: 24 }),
                        new TextRun({ text: "drawdown", font: "Times New Roman", size: 24, italics: true }),
                        new TextRun({ text: " yang dalam). Untuk mengatasi tantangan tersebut, penelitian terkini mulai menggabungkan Teori Jaringan Kompleks (", font: "Times New Roman", size: 24 }),
                        new TextRun({ text: "Complex Network Theory", font: "Times New Roman", size: 24, italics: true }),
                        new TextRun({ text: ") ke dalam optimasi portofolio [2]. Penggunaan instrumen seperti ", font: "Times New Roman", size: 24 }),
                        new TextRun({ text: "Minimum Spanning Tree (MST)", font: "Times New Roman", size: 24, italics: true }),
                        new TextRun({ text: " dan ", font: "Times New Roman", size: 24 }),
                        new TextRun({ text: "Eigenvector Centrality", font: "Times New Roman", size: 24, italics: true }),
                        new TextRun({ text: " terbukti efisien dalam memetakan interaksi antar aset dan menghukum (", font: "Times New Roman", size: 24 }),
                        new TextRun({ text: "penalize", font: "Times New Roman", size: 24, italics: true }),
                        new TextRun({ text: ") aset-aset yang menjadi titik pusat kegagalan sistemik.", font: "Times New Roman", size: 24 }),
                    ]
                }),
                new Paragraph({
                    alignment: AlignmentType.JUSTIFIED,
                    spacing: { before: 0, after: 120, line: 360, lineRule: "auto" },
                    indent: { firstLine: 720 },
                    children: [
                        new TextRun({ text: "Kendati model Network Markowitz statis menunjukkan proteksi yang lebih relevan dibandingkan Classical Markowitz, penentuan faktor penalti sentralitas \u03b3 (\u03b3) yang kaku kerap menimbulkan masalah di fase pasar yang dinamis (misalnya fase ", font: "Times New Roman", size: 24 }),
                        new TextRun({ text: "recovery", font: "Times New Roman", size: 24, italics: true }),
                        new TextRun({ text: " atau ", font: "Times New Roman", size: 24 }),
                        new TextRun({ text: "bullish", font: "Times New Roman", size: 24, italics: true }),
                        new TextRun({ text: "). Oleh karena itu, diperlukan pendekatan adaptif yang mengintegrasikan teknik pembersihan sinyal berbasis ", font: "Times New Roman", size: 24 }),
                        new TextRun({ text: "Random Matrix Theory (RMT)", font: "Times New Roman", size: 24, italics: true }),
                        new TextRun({ text: " disandingkan dengan optimalisasi ", font: "Times New Roman", size: 24 }),
                        new TextRun({ text: "Grid Search", font: "Times New Roman", size: 24, italics: true }),
                        new TextRun({ text: " dinamis (", font: "Times New Roman", size: 24 }),
                        new TextRun({ text: "rolling window", font: "Times New Roman", size: 24, italics: true }),
                        new TextRun({ text: ") agar portofolio dapat membentengi aset di saat ", font: "Times New Roman", size: 24 }),
                        new TextRun({ text: "crypto winter", font: "Times New Roman", size: 24, italics: true }),
                        new TextRun({ text: " tanpa mengorbankan rasio ", font: "Times New Roman", size: 24 }),
                        new TextRun({ text: "upside gain", font: "Times New Roman", size: 24, italics: true }),
                        new TextRun({ text: " di saat pembalikan arah (", font: "Times New Roman", size: 24 }),
                        new TextRun({ text: "bullish", font: "Times New Roman", size: 24, italics: true }),
                        new TextRun({ text: ").", font: "Times New Roman", size: 24 }),
                    ]
                }),
                emptyLine(),
                heading2("1.2. Identifikasi Masalah"),
                new Paragraph({
                    alignment: AlignmentType.JUSTIFIED,
                    spacing: { before: 0, after: 120, line: 360, lineRule: "auto" },
                    indent: { firstLine: 720 },
                    children: [
                        new TextRun({ text: "Berdasarkan latar belakang di atas, dapat diidentifikasi masalah sebagai berikut:", font: "Times New Roman", size: 24 }),
                    ]
                }),
                letterItem("Model Classical Markowitz rentan terhadap spurious correlations pada aset kripto yang bervolatilitas sangat tinggi, terutama pada kondisi krisis ekstrem.", "letters"),
                new Paragraph({
                    numbering: { reference: "letters", level: 0 },
                    alignment: AlignmentType.JUSTIFIED,
                    spacing: { before: 0, after: 80, line: 360, lineRule: "auto" },
                    children: [
                        new TextRun({ text: "Implementasi Network Markowitz yang telah ada sering kali menggunakan hiperparameter penalti sentralitas (\u03b3) bernilai konstan (statis), yang berpotensi menjadi bumerang saat pasar memasuki rezim ", font: "Times New Roman", size: 24 }),
                        new TextRun({ text: "recovery", font: "Times New Roman", size: 24, italics: true }),
                        new TextRun({ text: " atau ", font: "Times New Roman", size: 24 }),
                        new TextRun({ text: "bullish", font: "Times New Roman", size: 24, italics: true }),
                        new TextRun({ text: ".", font: "Times New Roman", size: 24 }),
                    ]
                }),
                new Paragraph({
                    numbering: { reference: "letters", level: 0 },
                    alignment: AlignmentType.JUSTIFIED,
                    spacing: { before: 0, after: 80, line: 360, lineRule: "auto" },
                    children: [
                        new TextRun({ text: "Kurangnya model sistematis yang secara metodis menyesuaikan struktur jaringan portofolio dengan rezim sikluk guncangan harga terkini menggunakan teknik optimasi ", font: "Times New Roman", size: 24 }),
                        new TextRun({ text: "rolling window", font: "Times New Roman", size: 24, italics: true }),
                        new TextRun({ text: " untuk aset-aset kripto utama.", font: "Times New Roman", size: 24 }),
                    ]
                }),
                emptyLine(),
                heading2("1.3. Tujuan Penelitian"),
                body("Tujuan dari penelitian ini adalah:"),
                letterItem("Menganalisis keandalan metodologi Network Markowitz (dengan integrasi RMT filter dan Eigenvector Centrality) dalam menekan ekstrimitas downside risk dibandingkan pendekatan portofolio naif dan konvensional.", "letters1"),
                letterItem("Merancang dan menguji model Network Markowitz adaptif (Grid Search Optimization) yang mampu melakukan re-kalibrasi dinamis dengan menggunakan paradigma rolling window pada berbagai lanskap pasar (Bearish, Recovery, Stable).", "letters1"),
                letterItem("Membandingkan performa perlindungan risiko sistemik (VaR) dan asimetri imbal hasil (Rachev Ratio) antara pemodelan baru dengan metode-metode baseline pada reksadana aset kripto.", "letters1"),
                emptyLine(),
                heading2("1.4. Ruang Lingkup Penelitian"),
                body("Ruang lingkup penelitian ini dibatasi pada:"),
                letterItem("Objek penelitian terfokus pada data fluktuasi harga harian dari 10 (sepuluh) aset kripto utama dalam kerangka waktu historis termasuk masa resesi crypto winter (14 September 2017 hingga 17 Oktober 2019).", "letters2"),
                letterItem("Metode yang dibandingkan secara teknis mencakup Equally Weighted (EW), Classical Markowitz (CM), Glasso Markowitz (GM), Network Markowitz statis (\u03b3 = 0, 1.0, 2.0), serta Optimized Network Markowitz secara dinamis berbasis Grid Search.", "letters2"),
                letterItem("Pengujian (backtesting) dilakukan dalam out-of-sample rolling window (120 observasi ke belakang dengan frekuensi penyesuaian rebalance 7 hari) yang disimulasikan menggunakan transaction cost atau estimasi biaya bursa (0.1%).", "letters2"),
                emptyLine(),
                heading2("1.5. Sistematika Penulisan"),
                body("Sistematika penulisan proposal tesis ini disusun sebagai berikut:"),
                emptyLine(),
                new Paragraph({
                    spacing: { before: 0, after: 80 },
                    children: [new TextRun({ text: "BAB I PENDAHULUAN", font: "Times New Roman", size: 24, bold: true })]
                }),
                body("Bab ini membahas latar belakang penelitian, identifikasi masalah, tujuan penelitian, ruang lingkup penelitian, dan sistematika penulisan."),
                new Paragraph({
                    spacing: { before: 80, after: 80 },
                    children: [new TextRun({ text: "BAB II LANDASAN/KERANGKA PEMIKIRAN", font: "Times New Roman", size: 24, bold: true })]
                }),
                body("Bab ini membahas kerangka teori yang relevan mencakup Teori Portofolio Modern, Complex Network Theory, dan Manajemen Risiko, serta tinjauan pustaka terhadap penelitian terdahulu di bidang Robo-Advisory kripto."),
                new Paragraph({
                    spacing: { before: 80, after: 80 },
                    children: [new TextRun({ text: "BAB III METODOLOGI PENELITIAN", font: "Times New Roman", size: 24, bold: true })]
                }),
                body("Bab ini menjelaskan alur sistematis eksperimen yang digunakan untuk memproses data instrumen kripto, penyaringan RMT pada matriks korelasi historis, pembangunan MST, fungsi objektif Markowitz modifikasi, dan komputasi skema backtesting."),
            ],
            footers: {
                default: new Footer({
                    children: [
                        new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [new TextRun({ children: [PageNumber.CURRENT], font: "Times New Roman", size: 24 })]
                        })
                    ]
                })
            }
        },
        // ==================== BAB II ====================
        {
            properties: {
                page: {
                    size: { width: 11906, height: 16838 },
                    margin: { top: 1440, right: 1440, bottom: 1440, left: 1800 }
                }
                // Lanjutkan penomoran decimal dari section sebelumnya
            },
            children: [
                centeredBold("BAB II", 26),
                centeredBold("LANDASAN/KERANGKA PEMIKIRAN", 26),
                emptyLine(),
                heading2("2.1. Kerangka Teori"),
                heading3("2.1.1. Modern Portfolio Theory (Markowitz)"),
                mixedBody([
                    {text: "Teori Portofolio Modern, yang dipelopori oleh Harry Markowitz, berupaya memaksimalkan imbal hasil yang diharapkan ("},
                    {text: "expected return", italic: true},
                    {text: ") pada tingkat risiko ("},
                    {text: "variance", italic: true},
                    {text: ") tertentu, atau sebaliknya. Masalah mendasarnya adalah bahwa kovarians dari kumpulan aset finansial sangat sensitif terhadap nilai-nilai ekstrem historis, yang dikenal sebagai "},
                    {text: "Markowitz Curse", italic: true},
                    {text: " [4]."}
                ]),
                emptyLine(),
                heading3("2.1.2. Random Matrix Theory (RMT) dan Kompleksitas Jaringan"),
                mixedBody([
                    {text: "Teori "},
                    {text: "Random Matrix", italic: true},
                    {text: " memungkinkan disaringnya "},
                    {text: "noise", italic: true},
                    {text: " dari struktur korelasi dengan membandingkan nilai eigen ("},
                    {text: "eigenvalues", italic: true},
                    {text: ") struktur empiris terhadap nilai batas distribusi teoritis Marchenko-Pastur [3]. Ini merupakan "},
                    {text: "vital element", italic: true},
                    {text: " sebelum dilakukan visualisasi graf berupa "},
                    {text: "Minimum Spanning Tree (MST)", italic: true},
                    {text: "."}
                ]),
                emptyLine(),
                heading3("2.1.3. Network Markowitz"),
                mixedBody([
                    {text: "Diperkenalkan baru-baru ini untuk penanganan "},
                    {text: "robo-advisory", italic: true},
                    {text: " pada kripto, komponen sentralitas "},
                    {text: "eigenvector", italic: true},
                    {text: " ditambahkan sebagai instrumen penalti di dalam penyelesaian optimasi "},
                    {text: "Mean-Variance", italic: true},
                    {text: " [1]. Sentralitas mewakili kerentanan sebuah aset mentransmisikan "},
                    {text: "shock", italic: true},
                    {text: " pada seluruh jaringan koin di pasar."}
                ]),
                emptyLine(),
                heading3("2.1.4. Penelitian Terdahulu"),
                body("Beberapa penelitian terdahulu yang relevan dengan penelitian ini antara lain:"),
                emptyLine(),
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [new TextRun({ text: "Tabel II.1. Perbandingan Penelitian Terdahulu", font: "Times New Roman", size: 24, bold: true })]
                }),
                new Table({
                    width: { size: 9026, type: WidthType.DXA },
                    columnWidths: [700, 2500, 3000, 2826],
                    rows: [
                        new TableRow({
                            tableHeader: true,
                            children: [
                                new TableCell({
                                    borders, width: { size: 700, type: WidthType.DXA },
                                    shading: { fill: "D5E8F0", type: ShadingType.CLEAR },
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    verticalAlign: VerticalAlign.CENTER,
                                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "No", font: "Times New Roman", size: 22, bold: true })] })]
                                }),
                                new TableCell({
                                    borders, width: { size: 2500, type: WidthType.DXA },
                                    shading: { fill: "D5E8F0", type: ShadingType.CLEAR },
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Penulis / Tahun", font: "Times New Roman", size: 22, bold: true })] })]
                                }),
                                new TableCell({
                                    borders, width: { size: 3000, type: WidthType.DXA },
                                    shading: { fill: "D5E8F0", type: ShadingType.CLEAR },
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Judul", font: "Times New Roman", size: 22, bold: true })] })]
                                }),
                                new TableCell({
                                    borders, width: { size: 2826, type: WidthType.DXA },
                                    shading: { fill: "D5E8F0", type: ShadingType.CLEAR },
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Hasil", font: "Times New Roman", size: 22, bold: true })] })]
                                }),
                            ]
                        }),
                        new TableRow({
                            children: [
                                new TableCell({
                                    borders, width: { size: 700, type: WidthType.DXA },
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "1", font: "Times New Roman", size: 22 })] })]
                                }),
                                new TableCell({
                                    borders, width: { size: 2500, type: WidthType.DXA },
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    children: [new Paragraph({ children: [new TextRun({ text: "Giudici, et al. (2020)", font: "Times New Roman", size: 22 })] })]
                                }),
                                new TableCell({
                                    borders, width: { size: 3000, type: WidthType.DXA },
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    children: [new Paragraph({ children: [new TextRun({ text: "Network Models to Improve Automated Cryptocurrency Portfolio Management", font: "Times New Roman", size: 22 })] })]
                                }),
                                new TableCell({
                                    borders, width: { size: 2826, type: WidthType.DXA },
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    children: [new Paragraph({ children: [new TextRun({ text: "Mengusulkan Network Markowitz dan sukses mendemonstrasikan perbaikan struktur dibandingkan Markowitz biasa di era crypto winter.", font: "Times New Roman", size: 22 })] })]
                                }),
                            ]
                        }),
                        new TableRow({
                            children: [
                                new TableCell({
                                    borders, width: { size: 700, type: WidthType.DXA },
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "2", font: "Times New Roman", size: 22 })] })]
                                }),
                                new TableCell({
                                    borders, width: { size: 2500, type: WidthType.DXA },
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    children: [new Paragraph({ children: [new TextRun({ text: "Papenbrock (2011)", font: "Times New Roman", size: 22 })] })]
                                }),
                                new TableCell({
                                    borders, width: { size: 3000, type: WidthType.DXA },
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    children: [new Paragraph({ children: [new TextRun({ text: "Financial networks and risk management", font: "Times New Roman", size: 22 })] })]
                                }),
                                new TableCell({
                                    borders, width: { size: 2826, type: WidthType.DXA },
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    children: [new Paragraph({ children: [new TextRun({ text: "Bukti bahwa penggunaan analisis jaringan meningkatkan akurasi risk metrics.", font: "Times New Roman", size: 22 })] })]
                                }),
                            ]
                        }),
                        new TableRow({
                            children: [
                                new TableCell({
                                    borders, width: { size: 700, type: WidthType.DXA },
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "3", font: "Times New Roman", size: 22 })] })]
                                }),
                                new TableCell({
                                    borders, width: { size: 2500, type: WidthType.DXA },
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    children: [new Paragraph({ children: [new TextRun({ text: "Mantegna (1999)", font: "Times New Roman", size: 22 })] })]
                                }),
                                new TableCell({
                                    borders, width: { size: 3000, type: WidthType.DXA },
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    children: [new Paragraph({ children: [new TextRun({ text: "Hierarchical structure in financial markets", font: "Times New Roman", size: 22 })] })]
                                }),
                                new TableCell({
                                    borders, width: { size: 2826, type: WidthType.DXA },
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    children: [new Paragraph({ children: [new TextRun({ text: "MST sukses memotret topologi kedekatan dan konektivitas sektor finansial dari pergerakan deret waktu pasar saham.", font: "Times New Roman", size: 22 })] })]
                                }),
                            ]
                        }),
                    ]
                }),
            ],
            footers: {
                default: new Footer({
                    children: [
                        new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [new TextRun({ children: [PageNumber.CURRENT], font: "Times New Roman", size: 24 })]
                        })
                    ]
                })
            }
        },
        // ==================== BAB III ====================
        {
            properties: {
                page: {
                    size: { width: 11906, height: 16838 },
                    margin: { top: 1440, right: 1440, bottom: 1440, left: 1800 }
                }
                // Lanjutkan penomoran decimal dari section sebelumnya
            },
            children: [
                centeredBold("BAB III", 26),
                centeredBold("METODOLOGI PENELITIAN", 26),
                emptyLine(),
                heading2("3.1. Tahapan Penelitian"),
                body("Penelitian ini dilaksanakan melalui tahapan-tahapan sebagai berikut:"),
                numItem("Studi Literatur"),
                mixedBody([
                    {text: "Melakukan kajian komprehensif terhadap jurnal yang membahas "},
                    {text: "crypto portfolio optimization", italic: true},
                    {text: ", "},
                    {text: "graph theory", italic: true},
                    {text: ", dan "},
                    {text: "Network Markowitz", italic: true},
                    {text: "."}
                ]),
                numItem("Pengumpulan Data & Pengkondisian"),
                mixedBody([
                    {text: "Mengumpulkan seri harga "},
                    {text: "close", italic: true},
                    {text: " "},
                    {text: "cryptocurrency", italic: true},
                    {text: ". Harga lalu diubah menjadi format "},
                    {text: "log returns", italic: true},
                    {text: "."}
                ]),
                numItem("Pra-pemrosesan Data (RMT Filtering)"),
                mixedBody([
                    {text: "Data "},
                    {text: "returns", italic: true},
                    {text: " akan diubah menjadi Matriks Korelasi (Pearson). Selanjutnya struktur "},
                    {text: "noise", italic: true},
                    {text: " historis dipotong menggunakan mekanisme filtrasi nilai eigen ("},
                    {text: "eigenvalue clipping boundary", italic: true},
                    {text: " batas Marchenko-Pastur)."}
                ]),
                numItem("Kuantifikasi Jaringan Aset (MST & Centrality)"),
                mixedBody([
                    {text: "Pembangunan jarak konektivitas ("},
                    {text: "distance matrix", italic: true},
                    {text: ") dari hasil korelasi terfilter untuk diekstraksi ke bentuk graf pohon "},
                    {text: "Minimum Spanning Tree (MST)", italic: true},
                    {text: ". "},
                    {text: "Node importance", italic: true},
                    {text: " kemudian diukur melalui "},
                    {text: "Eigenvector Centrality", italic: true},
                    {text: "."}
                ]),
                numItem("Optimasi Jaringan Adaptif (Dynamic Grid-Search)"),
                mixedBody([
                    {text: "Merancang algoritma komputasi untuk menyesuaikan parameter "},
                    {text: "impact", italic: true},
                    {text: " korelasi (\u03b3) yang secara otomatis mengekang instrumen-instrumen bervolatilitas sistemik pada jendela uji terkalibrasi ke belakang ("},
                    {text: "backtrack validity", italic: true},
                    {text: ")."}
                ]),
                numItem("Eksekusi Backtesting Portofolio"),
                mixedBody([
                    {text: "Mensimulasikan pembelian pada titik waktu (t) dan meninjau portofolio secara periodik menggunakan sistem "},
                    {text: "Rolling Window", italic: true},
                    {text: ". Terdapat perlakuan pengenaan "},
                    {text: "slippage/transaction cost", italic: true},
                    {text: " pada "},
                    {text: "rebalancing", italic: true},
                    {text: " harian."}
                ]),
                numItem("Pengukuran Evaluasi Resiko (Performance Metrics)"),
                mixedBody([
                    {text: "Mengukur nilai profit kumulatif, asimetri VaR 95%, hingga penyesuaian "},
                    {text: "Sharpe Ratio", italic: true},
                    {text: " dan "},
                    {text: "Rachev Ratio", italic: true},
                    {text: " di sepanjang berbagai transisi fasa pasar ekstrem."}
                ]),
                emptyLine(),
                heading2("3.2. Alat dan Bahan Penelitian"),
                heading3("3.2.1. Perangkat Lunak"),
                bulletItem("Sistem Operasi: Windows 10/11"),
                bulletItem("Bahasa Pemrograman: Python 3.x (dengan ekosistem Anaconda)"),
                bulletItem("Framework/Library: Pandas, Numpy, Scipy (Optimization), Scikit-Learn, NetworkX (Graph Analytics)"),
                emptyLine(),
                heading2("3.3. Dataset"),
                mixedBody([
                    {text: "Data historis harga harian diperoleh dari "},
                    {text: "Yahoo Finance", italic: true},
                    {text: " menggunakan "},
                    {text: "library yfinance", italic: true},
                    {text: " pada Python. Data mencakup 10 aset "},
                    {text: "cryptocurrency", italic: true},
                    {text: " berkapitalisasi pasar tinggi, yaitu: Bitcoin (BTC), Ethereum (ETH), Ripple (XRP), Tether (USDT), Bitcoin Cash (BCH), Litecoin (LTC), Binance Coin (BNB), EOS (EOS), Stellar (XLM), dan Tron (TRX). Pemilihan aset ini selaras dengan referensi Giudici et al. (2020) yang menjadi landasan penelitian ini."}
                ]),
                mixedBody([
                    {text: "Periode pengambilan data adalah 14 September 2017 hingga 17 Oktober 2019, yang secara sengaja dipilih untuk mencakup tiga rezim pasar berbeda: era "},
                    {text: "speculative bubble", italic: true},
                    {text: " (akhir 2017), era "},
                    {text: "crypto winter", italic: true},
                    {text: " atau "},
                    {text: "bear market", italic: true},
                    {text: " berkepanjangan (2018), serta fase awal "},
                    {text: "recovery", italic: true},
                    {text: " dan stabilisasi (2019). Rentang ini menghasilkan total 762 hari perdagangan setelah proses pengkondisian data."}
                ]),
                mixedBody([
                    {text: "Data harga penutupan harian ("},
                    {text: "Close Price", italic: true},
                    {text: ") kemudian dikonversi menjadi "},
                    {text: "log returns", italic: true},
                    {text: " harian menggunakan formula r\u209C = ln(P\u209C / P\u209C\u208B\u2081). Transformasi ini dipilih karena menormalkan distribusi imbal hasil dan memiliki sifat aditif antar waktu yang lebih stabil secara statistik. Untuk menangani nilai yang hilang ("},
                    {text: "missing values", italic: true},
                    {text: ") akibat perbedaan hari perdagangan antar aset, diterapkan metode "},
                    {text: "forward fill", italic: true},
                    {text: " dilanjutkan "},
                    {text: "backward fill", italic: true},
                    {text: " secara berurutan."}
                ]),
                mixedBody([
                    {text: "Keluaran akhir data disimpan dalam format "},
                    {text: "spreadsheet", italic: true},
                    {text: " (.xlsx) yang mencakup lima tabel: (1) "},
                    {text: "Returns", italic: true},
                    {text: " — matriks "},
                    {text: "log returns", italic: true},
                    {text: " harian 10 aset; (2) "},
                    {text: "Prices", italic: true},
                    {text: " — harga ternormalisasi ke nilai awal 100; (3) "},
                    {text: "Statistics", italic: true},
                    {text: " — ringkasan statistik deskriptif (rata-rata, standar deviasi, "},
                    {text: "kurtosis", italic: true},
                    {text: ", "},
                    {text: "skewness", italic: true},
                    {text: ", min, maks); (4) "},
                    {text: "Correlation", italic: true},
                    {text: " — matriks korelasi Pearson antar aset; dan (5) "},
                    {text: "Metadata", italic: true},
                    {text: " — informasi parameter utama dataset."}
                ]),
                mixedBody([
                    {text: "Dalam proses pengunduhan data nyata, ditemukan beberapa tantangan kualitas data yang perlu ditangani secara eksplisit. Pertama, terdapat "},
                    {text: "missing values", italic: true},
                    {text: " yang signifikan pada beberapa aset, terutama Binance Coin (BNB), EOS, dan Tron (TRX) yang baru diluncurkan setelah periode dimulai (September 2017), sehingga data awal mereka bersifat kosong ("},
                    {text: "NaN", italic: true},
                    {text: "). Kondisi ini ditangani dengan strategi "},
                    {text: "forward fill", italic: true},
                    {text: " diikuti "},
                    {text: "backward fill", italic: true},
                    {text: " untuk memastikan tidak ada baris yang kosong sebelum kalkulasi "},
                    {text: "log returns", italic: true},
                    {text: " dilakukan."}
                ]),
                mixedBody([
                    {text: "Kedua, terdapat potensi inkonsistensi data harga yang disebabkan oleh perbedaan jadwal bursa kripto antar platform. Karena "},
                    {text: "Yahoo Finance", italic: true},
                    {text: " mengacu pada harga agregat dari berbagai sumber, nilai harga penutupan ("},
                    {text: "close price", italic: true},
                    {text: ") kadang mengalami lonjakan ekstrem sesaat ("},
                    {text: "outlier", italic: true},
                    {text: ") yang tidak mencerminkan kondisi pasar sesungguhnya. Nilai-nilai ini diidentifikasi melalui inspeksi visual terhadap distribusi "},
                    {text: "log returns", italic: true},
                    {text: " dan dibiarkan dalam dataset karena memang merupakan bagian dari volatilitas nyata aset kripto yang hendak dimodelkan."}
                ]),
                mixedBody([
                    {text: "Ketiga, Tether (USDT) sebagai "},
                    {text: "stablecoin", italic: true},
                    {text: " memiliki karakteristik distribusi yang berbeda dari aset lainnya — dengan volatilitas mendekati nol dan korelasi yang sangat rendah terhadap semua aset. Keberadaannya dalam pool tetap dipertahankan karena merupakan bagian dari 10 aset teratas berdasarkan kapitalisasi pasar pada periode tersebut, dan memberikan kontribusi diversifikasi yang unik dalam konstruksi jaringan portofolio."}
                ]),
                emptyLine(),
                heading2("3.4. Metode/Algoritma yang Digunakan"),
                mixedBody([
                    {text: "Optimasi yang diajukan akan merubah fungsi pencarian model klasik Markowitz menjadi kerangka berbasis "},
                    {text: "penalty function", italic: true},
                    {text: " dinamis. Formula "},
                    {text: "Network Markowitz", italic: true},
                    {text: ":"}
                ]),
                mixedBody([
                    {text: "\u03B3  = Parameter skalar reguler untuk tingkat penghukuman ("},
                    {text: "penalty level", italic: true},
                    {text: ") sentralitas graf."}
                ]),
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    spacing: { before: 120, after: 120 },
                    children: [new TextRun({ text: "min\u1D64  w\u1D40 \u00B7 Sf \u00B7 w + \u03B3 \u03A3(Ce \u00B7 w)    ... (III.1)", font: "Times New Roman", size: 24, italics: true })]
                }),
                body("Keterangan:"),
                bodyNoIndent("w  = Vektor alokasi bobot untuk setiap aset kripto (sumbangan = 1)."),
                bodyNoIndent("Sf = Matriks Kovarians terfilter RMT."),
                bodyNoIndent("\u03B3  = Parameter skalar reguler untuk tingkat penghukuman (penalty level) sentralitas graf."),
                bodyNoIndent("Ce = Vektor skor Eigenvector Centrality tiap-tiap entitas node."),
                emptyLine(),
                mixedBody([
                    {text: "Pada penelitian ini, bobot \u03b3 tidak akan dilakukan "},
                    {text: "hard-coded", italic: true},
                    {text: " statis, melainkan secara luwes dan "},
                    {text: "rolling", italic: true},
                    {text: " akan difungsikan optimasi "},
                    {text: "grid validation", italic: true},
                    {text: " berbasis metrik obyektif "},
                    {text: "Sharpe Ratio", italic: true},
                    {text: " periode belakang."}
                ]),
                emptyLine(),
                heading2("3.5. Rencana Jadwal Penelitian"),
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [new TextRun({ text: "Tabel III.1. Rencana Jadwal Penelitian", font: "Times New Roman", size: 24, bold: true })]
                }),
                new Table({
                    width: { size: 9026, type: WidthType.DXA },
                    columnWidths: [2900, 1021, 1021, 1021, 1021, 1021, 1021],
                    rows: [
                        new TableRow({
                            tableHeader: true,
                            children: [
                                new TableCell({
                                    borders, width: { size: 2900, type: WidthType.DXA },
                                    shading: { fill: "D5E8F0", type: ShadingType.CLEAR },
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Kegiatan", font: "Times New Roman", size: 22, bold: true })] })]
                                }),
                                ...["Bln 1", "Bln 2", "Bln 3", "Bln 4", "Bln 5", "Bln 6"].map(h => new TableCell({
                                    borders, width: { size: 1021, type: WidthType.DXA },
                                    shading: { fill: "D5E8F0", type: ShadingType.CLEAR },
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: h, font: "Times New Roman", size: 22, bold: true })] })]
                                }))
                            ]
                        }),
                        ...[
                            ["Studi Literatur", ["✓", "", "", "", "", ""]],
                            ["Pengumpulan Data", ["✓", "✓", "", "", "", ""]],
                            ["Pra-pemrosesan Data", ["", "✓", "✓", "", "", ""]],
                            ["Perancangan Model/Sistem", ["", "", "✓", "✓", "", ""]],
                            ["Implementasi", ["", "", "", "✓", "✓", ""]],
                            ["Pengujian dan Evaluasi", ["", "", "", "", "✓", "✓"]],
                            ["Penulisan Laporan", ["✓", "✓", "✓", "✓", "", ""]],
                        ].map(([activity, marks]) => new TableRow({
                            children: [
                                new TableCell({
                                    borders, width: { size: 2900, type: WidthType.DXA },
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    children: [new Paragraph({ children: [new TextRun({ text: activity, font: "Times New Roman", size: 22 })] })]
                                }),
                                ...marks.map(m => new TableCell({
                                    borders, width: { size: 1021, type: WidthType.DXA },
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: m, font: "Times New Roman", size: 22 })] })]
                                }))
                            ]
                        }))
                    ]
                }),
            ],
            footers: {
                default: new Footer({
                    children: [
                        new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [new TextRun({ children: [PageNumber.CURRENT], font: "Times New Roman", size: 24 })]
                        })
                    ]
                })
            }
        },
        // ==================== DAFTAR REFERENSI ====================
        {
            properties: {
                page: {
                    size: { width: 11906, height: 16838 },
                    margin: { top: 1440, right: 1440, bottom: 1440, left: 1800 }
                }
                // Lanjutkan penomoran decimal dari section sebelumnya
            },
            children: [
                centeredBold("DAFTAR REFERENSI", 26),
                emptyLine(),
                new Paragraph({
                    alignment: AlignmentType.JUSTIFIED,
                    spacing: { before: 0, after: 120, line: 360 },
                    indent: { left: 720, hanging: 720 },
                    children: [new TextRun({ text: "[1] P. Giudici, A. Sariev, and G. Toscani, \"Network Models to Improve Automated Cryptocurrency Portfolio Management,\" Risks, vol. 8, no. 3, p. 96, 2020. https://doi.org/10.3390/risks8030096.", font: "Times New Roman", size: 24 })]
                }),
                new Paragraph({
                    alignment: AlignmentType.JUSTIFIED,
                    spacing: { before: 0, after: 120, line: 360 },
                    indent: { left: 720, hanging: 720 },
                    children: [new TextRun({ text: "[2] S. S. Momeni dan R. E. Rostami, \"Portfolio selection using a hybrid methodology of network theory and optimization,\" Annals of Operations Research, vol. 308, pp. 317–345, 2021.", font: "Times New Roman", size: 24 })]
                }),
                new Paragraph({
                    alignment: AlignmentType.JUSTIFIED,
                    spacing: { before: 0, after: 120, line: 360 },
                    indent: { left: 720, hanging: 720 },
                    children: [new TextRun({ text: "[3] V. A. Marchenko and L. A. Pastur, \"Distribution of eigenvalues for some sets of random matrices,\" Mathematics of the USSR-Sbornik, vol. 1, no. 4, pp. 457–483, 1967.", font: "Times New Roman", size: 24 })]
                }),
                new Paragraph({
                    alignment: AlignmentType.JUSTIFIED,
                    spacing: { before: 0, after: 120, line: 360 },
                    indent: { left: 720, hanging: 720 },
                    children: [new TextRun({ text: "[4] H. Markowitz, \"Portfolio Selection,\" The Journal of Finance, vol. 7, no. 1, pp. 77–91, 1952.", font: "Times New Roman", size: 24 })]
                }),
                new Paragraph({
                    alignment: AlignmentType.JUSTIFIED,
                    spacing: { before: 0, after: 120, line: 360 },
                    indent: { left: 720, hanging: 720 },
                    children: [new TextRun({ text: "[5] R. N. Mantegna, \"Hierarchical structure in financial markets,\" The European Physical Journal B, vol. 11, no. 1, pp. 193–197, 1999.", font: "Times New Roman", size: 24 })]
                }),
            ],
            footers: {
                default: new Footer({
                    children: [
                        new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [new TextRun({ children: [PageNumber.CURRENT], font: "Times New Roman", size: 24 })]
                        })
                    ]
                })
            }
        },
    ]
});

Packer.toBuffer(doc).then(buffer => {
    fs.writeFileSync("proposal_tesis_ragil.docx", buffer);
    console.log("Done!");
});
