const docx = require('docx');
const {
    Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
    AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
    LevelFormat, PageNumber, PageBreak, TabStopType, TabStopPosition,
    VerticalAlign, ImageRun, Footer, NumberFormat, SectionType
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
                    margin: { top: 1701, right: 1417, bottom: 1417, left: 2268 }
                }
            },
            children: [
                emptyLine(),
                centeredBold("OPTIMASI DINAMIS PEMODELAN NETWORK MARKOWITZ", 28),
                centeredBold("UNTUK MANAJEMEN PORTOFOLIO MATA UANG KRIPTO", 28),
                emptyLine(),
                emptyLine(),
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                        new ImageRun({
                            data: fs.readFileSync("e:\\ProjectNodeJs\\temp_doc_build\\logo_unm.png"),
                            transformation: {
                                width: 200, // Reduced slightly to save space
                                height: 200,
                            },
                        }),
                    ],
                }),
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
                emptyLine(),
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
                    margin: { top: 1701, right: 1417, bottom: 1417, left: 2268 },
                    pageNumbers: { start: 2, formatType: NumberFormat.LOWER_ROMAN }
                },
                type: SectionType.NEXT_PAGE,
                pageNumbers: { start: 2, formatType: NumberFormat.LOWER_ROMAN }
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
                new Paragraph({
                    children: [new TextRun({ text: "Jakarta, 12 Maret 2026", font: "Times New Roman", size: 24 })]
                }),
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
                    margin: { top: 1701, right: 1417, bottom: 1417, left: 2268 },
                    pageNumbers: { formatType: NumberFormat.LOWER_ROMAN }
                },
                type: SectionType.NEXT_PAGE,
                pageNumbers: { formatType: NumberFormat.LOWER_ROMAN }
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
                    margin: { top: 1701, right: 1417, bottom: 1417, left: 2268 },
                    pageNumbers: { start: 1, formatType: NumberFormat.DECIMAL }
                },
                type: SectionType.NEXT_PAGE,
                pageNumbers: { start: 1, formatType: NumberFormat.DECIMAL }
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
                        new TextRun({ text: " [1], [10], [11].", font: "Times New Roman", size: 24 }),
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
                        new TextRun({ text: ") ke dalam optimasi portofolio [1], [2]. Penggunaan instrumen seperti ", font: "Times New Roman", size: 24 }),
                        new TextRun({ text: "Minimum Spanning Tree (MST)", font: "Times New Roman", size: 24, italics: true }),
                        new TextRun({ text: " dan ", font: "Times New Roman", size: 24 }),
                        new TextRun({ text: "Eigenvector Centrality", font: "Times New Roman", size: 24, italics: true }),
                        new TextRun({ text: " terbukti efisien dalam memetakan interaksi antar aset dan menghukum (", font: "Times New Roman", size: 24 }),
                        new TextRun({ text: "penalize", font: "Times New Roman", size: 24, italics: true }),
                        new TextRun({ text: ") aset-aset yang menjadi titik pusat kegagalan sistemik [5], [6], [12].", font: "Times New Roman", size: 24 }),
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
                        new TextRun({ text: " [3], [10] disandingkan dengan optimalisasi ", font: "Times New Roman", size: 24 }),
                        new TextRun({ text: "Grid Search", font: "Times New Roman", size: 24, italics: true }),
                        new TextRun({ text: " dinamis (", font: "Times New Roman", size: 24 }),
                        new TextRun({ text: "rolling window", font: "Times New Roman", size: 24, italics: true }),
                        new TextRun({ text: ") agar portofolio dapat membentengi aset di saat ", font: "Times New Roman", size: 24 }),
                        new TextRun({ text: "crypto winter", font: "Times New Roman", size: 24, italics: true }),
                        new TextRun({ text: " tanpa mengorbankan rasio ", font: "Times New Roman", size: 24 }),
                        new TextRun({ text: "upside gain", font: "Times New Roman", size: 24, italics: true }),
                        new TextRun({ text: " di saat pembalikan arah (", font: "Times New Roman", size: 24 }),
                        new TextRun({ text: "bullish", font: "Times New Roman", size: 24, italics: true }),
                        new TextRun({ text: ") [1], [13].", font: "Times New Roman", size: 24 }),
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
                    margin: { top: 1701, right: 1417, bottom: 1417, left: 2268 }
                },
                type: SectionType.NEXT_PAGE
            },
            // Lanjutkan penomoran decimal dari section sebelumnya
            children: [
                centeredBold("BAB II", 26),
                centeredBold("LANDASAN/KERANGKA PEMIKIRAN", 26),
                emptyLine(),
                heading2("2.1. Kerangka Teori"),
                heading3("2.1.1. Modern Portfolio Theory (Markowitz)"),
                mixedBody([
                    {text: "Teori Portofolio Modern (MPT), yang dipelopori oleh Harry Markowitz, berupaya memaksimalkan imbal hasil yang diharapkan ("},
                    {text: "expected return", italic: true},
                    {text: ") pada tingkat risiko ("},
                    {text: "variance", italic: true},
                    {text: ") tertentu melalui pemilihan aset yang terdiversifikasi dalam sebuah "},
                    {text: "Efficient Frontier", italic: true},
                    {text: " [4]. Namun, dalam praktiknya, MPT memiliki keterbatasan signifikan terkait stabilitas estimasi. Masalah mendasarnya adalah bahwa kovarians dari kumpulan aset finansial sangat sensitif terhadap nilai-nilai ekstrem historis, yang dikenal sebagai "},
                    {text: "Markowitz Curse", italic: true},
                    {text: " atau enigma optimasi Markowitz [14]. Fenomena ini terjadi karena algoritma optimasi cenderung memperkuat kesalahan estimasi ("},
                    {text: "error maximization", italic: true},
                    {text: "), sehingga perubahan kecil pada input data dapat menghasilkan perubahan drastis pada alokasi bobot portofolio."}
                ]),
                emptyLine(),
                heading3("2.1.2. Random Matrix Theory (RMT) dan Kompleksitas Jaringan"),
                mixedBody([
                    {text: "Teori "},
                    {text: "Random Matrix", italic: true},
                    {text: " memungkinkan disaringnya "},
                    {text: "noise", italic: true},
                    {text: " dari struktur korelasi dengan memisahkan nilai eigen ("},
                    {text: "eigenvalues", italic: true},
                    {text: ") yang membawa informasi sinyal pasar dari nilai eigen yang bersifat acak. Berdasarkan distribusi Marchenko-Pastur, nilai eigen yang jatuh dalam rentang "},
                    {text: "noise bulk", italic: true},
                    {text: " dianggap sebagai residu statistik, sementara nilai eigen yang berada di luar batas tersebut merepresentasikan korelasi ekonomi yang nyata [3], [10]. Filtrasi ini merupakan "},
                    {text: "vital element", italic: true},
                    {text: " untuk memastikan stabilitas topologi jaringan sebelum dilakukan visualisasi graf berupa "},
                    {text: "Minimum Spanning Tree (MST)", italic: true},
                    {text: " [15]. Tanpa pembersihan RMT, struktur pohon yang dihasilkan cenderung tidak stabil dan sensitif terhadap fluktuasi data jangka pendek."}
                ]),
                emptyLine(),
                heading3("2.1.3. Network Markowitz"),
                mixedBody([
                    {text: "Diferensiasi utama "},
                    {text: "Network Markowitz", italic: true},
                    {text: " dibanding model klasik terletak pada integrasi risiko sistemik ke dalam fungsi optimasi. Diperkenalkan baru-baru ini untuk penanganan "},
                    {text: "robo-advisory", italic: true},
                    {text: " pada kripto, komponen sentralitas "},
                    {text: "eigenvector", italic: true},
                    {text: " ditambahkan sebagai instrumen penalti di dalam penyelesaian optimasi "},
                    {text: "Mean-Variance", italic: true},
                    {text: " [1], [9]. Sentralitas mewakili kerentanan sebuah aset mentransmisikan "},
                    {text: "shock", italic: true},
                    {text: " pada seluruh jaringan koin di pasar. Dengan memberikan penalti pada aset yang memiliki sentralitas tinggi, model ini secara proaktif mengurangi paparan terhadap 'titik pusat kegagalan' sistemik ("},
                    {text: "systemic points of failure", italic: true},
                    {text: "), yang terbukti sangat efektif dalam menjaga stabilitas portofolio saat fenomena penularan pasar ("},
                    {text: "market contagion", italic: true},
                    {text: ") terjadi."}
                ]),
                emptyLine(),
                heading3("2.1.4. Teori Siklus dan Rezim Pasar Kripto"),
                mixedBody([
                    {text: "Pasar kripto dicirikan oleh volatilitas yang jauh lebih tinggi dibandingkan pasar aset tradisional, dengan siklus yang terbagi dalam rezim pasar yang kontras. Secara teoritis, siklus ini terdiri dari fase "},
                    {text: "Bullish", italic: true},
                    {text: " (pertumbuhan eksponensial), "},
                    {text: "Bearish/Crypto Winter", italic: true},
                    {text: " (penyusutan nilai secara sistemik), dan "},
                    {text: "Recovery", italic: true},
                    {text: " (stabilisasi ulang). Dinamika korelasi antar aset cenderung meningkat drastis (korelasi positif kuat) saat pasar mengalami gejolak ("},
                    {text: "market crash", italic: true},
                    {text: "), yang secara drastis mengurangi manfaat diversifikasi model statis [11]. Oleh karena itu, pengenalan rezim pasar melalui pendekatan adaptif menjadi krusial untuk menjaga performa portofolio."}
                ]),
                emptyLine(),
                heading3("2.1.5. Analisis Kebaruan (Gap Analysis)"),
                mixedBody([
                    {text: "Penelitian ini memiliki kebaruan signifikan dibandingkan model yang diusulkan oleh Giudici et al. (2020). Jika penelitian tersebut menggunakan parameter penghukuman jaringan (\u03b3) yang bernilai statis (konstan), penelitian ini mengusulkan "},
                    {text: "Optimized Dynamic Network Markowitz", italic: true},
                    {text: ". Kebaruan utama terletak pada penggunaan algoritma "},
                    {text: "Grid Search", italic: true},
                    {text: " yang dikombinasikan dengan jendela uji bergulir ("},
                    {text: "rolling window", italic: true},
                    {text: ") untuk menentukan nilai \u03b3 yang termutakhir berdasarkan kondisi pasar terkini. Dengan demikian, model tidak hanya menyaring "},
                    {text: "noise", italic: true},
                    {text: " melalui RMT, tetapi juga secara aktif menyesuaikan intensitas kontrol sentralitas terhadap perubahan struktur jaringan aset secara "},
                    {text: "real-time", italic: true},
                    {text: "."}
                ]),
                emptyLine(),
                heading3("2.1.6. Penelitian Terdahulu"),
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
                                    children: [new Paragraph({ alignment: AlignmentType.JUSTIFIED, children: [new TextRun({ text: "Mengusulkan Network Markowitz dan sukses mendemonstrasikan perbaikan struktur dibandingkan Markowitz biasa di era crypto winter.", font: "Times New Roman", size: 22 })] })]
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
                                    children: [new Paragraph({ children: [new TextRun({ text: "Jing & Rocha (2023)", font: "Times New Roman", size: 22 })] })]
                                }),
                                new TableCell({
                                    borders, width: { size: 3000, type: WidthType.DXA },
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    children: [new Paragraph({ children: [new TextRun({ text: "A network-based strategy of price correlations for optimal cryptocurrency portfolios", font: "Times New Roman", size: 22 })] })]
                                }),
                                new TableCell({
                                    borders, width: { size: 2826, type: WidthType.DXA },
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    children: [new Paragraph({ alignment: AlignmentType.JUSTIFIED, children: [new TextRun({ text: "Menggabungkan MST dan MPT untuk memilih 46 dari 157 kripto berdasarkan dekorelasi jaringan; portofolio MST mengungguli seluruh benchmark (BTC, TOP5, RAND); koin populer berkapitalisasi besar terbukti jarang optimal.", font: "Times New Roman", size: 22 })] })]
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
                                    children: [new Paragraph({ children: [new TextRun({ text: "Kitanovski, et al. (2024)", font: "Times New Roman", size: 22 })] })]
                                }),
                                new TableCell({
                                    borders, width: { size: 3000, type: WidthType.DXA },
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    children: [new Paragraph({ children: [new TextRun({ text: "Network-based diversification of stock and cryptocurrency portfolios", font: "Times New Roman", size: 22 })] })]
                                }),
                                new TableCell({
                                    borders, width: { size: 2826, type: WidthType.DXA },
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    children: [new Paragraph({ alignment: AlignmentType.JUSTIFIED, children: [new TextRun({ text: "Menggunakan algoritma komunitas (Louvain & Affinity Propagation) untuk diversifikasi; strategi jaringan secara konsisten mengungguli portofolio acak dan indeks pasar; menunjukkan keunikan kripto di mana aset perifer (sentralitas rendah) menghasilkan return lebih tinggi.", font: "Times New Roman", size: 22 })] })]
                                }),
                            ]
                        }),
                        new TableRow({
                            children: [
                                new TableCell({
                                    borders, width: { size: 700, type: WidthType.DXA },
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "4", font: "Times New Roman", size: 22 })] })]
                                }),
                                new TableCell({
                                    borders, width: { size: 2500, type: WidthType.DXA },
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    children: [new Paragraph({ children: [new TextRun({ text: "Kitanovski, et al. (2022)", font: "Times New Roman", size: 22 })] })]
                                }),
                                new TableCell({
                                    borders, width: { size: 3000, type: WidthType.DXA },
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    children: [new Paragraph({ children: [new TextRun({ text: "Cryptocurrency Portfolio Diversification Using Network Community Detection", font: "Times New Roman", size: 22 })] })]
                                }),
                                new TableCell({
                                    borders, width: { size: 2826, type: WidthType.DXA },
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    children: [new Paragraph({ alignment: AlignmentType.JUSTIFIED, children: [new TextRun({ text: "Memanfaatkan deteksi komunitas (Louvain & Affinity Propagation) pada jaringan korelasi kripto untuk diversifikasi; membantu mengurangi volatilitas dan mengoptimalkan return bagi investor.", font: "Times New Roman", size: 22 })] })]
                                }),
                            ]
                        }),
                        new TableRow({
                            children: [
                                new TableCell({
                                    borders, width: { size: 700, type: WidthType.DXA },
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "5", font: "Times New Roman", size: 22 })] })]
                                }),
                                new TableCell({
                                    borders, width: { size: 2500, type: WidthType.DXA },
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    children: [new Paragraph({ children: [new TextRun({ text: "Giudici, et al. (2021)", font: "Times New Roman", size: 22 })] })]
                                }),
                                new TableCell({
                                    borders, width: { size: 3000, type: WidthType.DXA },
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    children: [new Paragraph({ children: [new TextRun({ text: "Network models to improve robot advisory portfolios", font: "Times New Roman", size: 22 })] })]
                                }),
                                new TableCell({
                                    borders, width: { size: 2826, type: WidthType.DXA },
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    children: [new Paragraph({ alignment: AlignmentType.JUSTIFIED, children: [new TextRun({ text: "Menunjukkan bahwa model jaringan dapat meningkatkan performa portofolio robotik dengan memitigasi risiko sistemik melalui struktur keterhubungan pasar yang lebih akurat.", font: "Times New Roman", size: 22 })] })]
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
                    margin: { top: 1701, right: 1417, bottom: 1417, left: 2268 }
                },
                type: SectionType.NEXT_PAGE
            },
            // Lanjutkan penomoran decimal dari section sebelumnya
            children: [
                centeredBold("BAB III", 26),
                centeredBold("METODOLOGI PENELITIAN", 26),
                emptyLine(),
                heading2("3.1. Tahapan Penelitian"),
                body("Penelitian ini dilaksanakan melalui tahapan-tahapan sebagai berikut:"),
                emptyLine(),
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                        new ImageRun({
                            data: fs.readFileSync("e:\\ProjectNodeJs\\temp_doc_build\\framwrok.jpg"),
                            transformation: {
                                width: 550,
                                height: 350,
                            },
                        }),
                    ],
                }),
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [new TextRun({ text: "Gambar III.1. Kerangka Kerja Penelitian Network Markowitz dengan Grid Search", font: "Times New Roman", size: 22, bold: true })]
                }),
                emptyLine(),
                mixedBody([
                    {text: "Secara visual, alur penelitian pada Gambar III.1 dibagi menjadi lima tahapan utama: (1) "},
                    {text: "Data Acquisition", italic: true},
                    {text: " yaitu pengumpulan data historis kripto; (2) "},
                    {text: "RMT Filtering", italic: true},
                    {text: " untuk pembersihan sinyal korelasi; (3) "},
                    {text: "MST & Eigenvector Centrality Calculation", italic: true},
                    {text: " untuk ekstraksi struktur jaringan; (4) "},
                    {text: "Dynamic Grid Search", italic: true},
                    {text: " (\u03b3-tuning) untuk optimasi parameter adaptif; dan (5) "},
                    {text: "Portfolio Output & Evaluation", italic: true},
                    {text: " untuk pengujian performa akhir."}
                ]),
                emptyLine(),
                emptyLine(),
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
                numItem("Data Acquisition"),
                mixedBody([
                    {text: "Mengumpulkan seri harga "},
                    {text: "close", italic: true},
                    {text: " "},
                    {text: "cryptocurrency", italic: true},
                    {text: " dari Yahoo Finance. Harga lalu diubah menjadi format "},
                    {text: "log returns", italic: true},
                    {text: " harian untuk menormalisasi distribusi imbal hasil."}
                ]),
                numItem("RMT Filtering"),
                mixedBody([
                    {text: "Mentransformasi "},
                    {text: "returns", italic: true},
                    {text: " menjadi Matriks Korelasi Pearson, kemudian menerapkan filtrasi "},
                    {text: "Random Matrix Theory (RMT)", italic: true},
                    {text: " melalui mekanisme "},
                    {text: "eigenvalue clipping", italic: true},
                    {text: " untuk memisahkan sinyal korelasi pasar dari "},
                    {text: "noise", italic: true},
                    {text: " statistik."}
                ]),
                numItem("MST & Eigenvector Centrality Calculation"),
                mixedBody([
                    {text: "Membangun "},
                    {text: "distance matrix", italic: true},
                    {text: " dari korelasi terfilter untuk mengekstraksi struktur pohon "},
                    {text: "Minimum Spanning Tree (MST)", italic: true},
                    {text: ". Selanjutnya, dilakukan kuantifikasi "},
                    {text: "node importance", italic: true},
                    {text: " menggunakan "},
                    {text: "Eigenvector Centrality", italic: true},
                    {text: " sebagai ukuran risiko sistemik tiap aset."}
                ]),
                numItem("Dynamic Grid Search (\u03b3-tuning)"),
                mixedBody([
                    {text: "Melakukan proses kalibrasi parameter penalti jaringan (\u03b3) secara dinamis menggunakan algoritma "},
                    {text: "Grid Search", italic: true},
                    {text: ". Parameter ini dioptimalkan pada setiap jendela waktu bergulir ("},
                    {text: "rolling window", italic: true},
                    {text: ") untuk memastikan model tetap adaptif terhadap perubahan rezim pasar."}
                ]),
                numItem("Portfolio Output & Evaluation"),
                mixedBody([
                    {text: "Eksekusi alokasi bobot portofolio pada periode "},
                    {text: "out-of-sample", italic: true},
                    {text: " dan melakukan evaluasi performa menggunakan metrik "},
                    {text: "Sharpe Ratio", italic: true},
                    {text: ", "},
                    {text: "Value at Risk (VaR)", italic: true},
                    {text: ", dan "},
                    {text: "Rachev Ratio", italic: true},
                    {text: " [18]."}
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
                    {text: " berkapitalisasi pasar tinggi yang dirinci pada tabel berikut:"}
                ]),
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [new TextRun({ text: "Tabel III.1. Deskripsi Dataset Cryptocurrency", font: "Times New Roman", size: 24, bold: true })]
                }),
                new Table({
                    width: { size: 9026, type: WidthType.DXA },
                    columnWidths: [1000, 2500, 3500, 2026],
                    rows: [
                        new TableRow({
                            tableHeader: true,
                            children: [
                                new TableCell({ borders, shading: { fill: "D5E8F0" }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Ticker", font: "Times New Roman", size: 22, bold: true })] })] }),
                                new TableCell({ borders, shading: { fill: "D5E8F0" }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Nama Aset", font: "Times New Roman", size: 22, bold: true })] })] }),
                                new TableCell({ borders, shading: { fill: "D5E8F0" }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Kategori / Use Case", font: "Times New Roman", size: 22, bold: true })] })] }),
                                new TableCell({ borders, shading: { fill: "D5E8F0" }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Sumber", font: "Times New Roman", size: 22, bold: true })] })] }),
                            ]
                        }),
                        ...[
                            ["BTC", "Bitcoin", "Layer 1 / Story of Value", "Yahoo Finance"],
                            ["ETH", "Ethereum", "Layer 1 / Smart Contract", "Yahoo Finance"],
                            ["XRP", "Ripple", "Payment / Bridge Currency", "Yahoo Finance"],
                            ["USDT", "Tether", "Stablecoin / USD Pegged", "Yahoo Finance"],
                            ["BCH", "Bitcoin Cash", "Payment / Peer-to-Peer Cash", "Yahoo Finance"],
                            ["LTC", "Litecoin", "Payment / Digital Silver", "Yahoo Finance"],
                            ["BNB", "Binance Coin", "Exchange Token / Layer 1", "Yahoo Finance"],
                            ["EOS", "EOS", "Layer 1 / Smart Contract", "Yahoo Finance"],
                            ["XLM", "Stellar", "Payment / Bridge Currency", "Yahoo Finance"],
                            ["TRX", "Tron", "Layer 1 / Smart Contract", "Yahoo Finance"],
                        ].map(([t, n, k, s]) => new TableRow({
                            children: [
                                new TableCell({ borders, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: t, font: "Times New Roman", size: 22 })] })] }),
                                new TableCell({ borders, children: [new Paragraph({ children: [new TextRun({ text: n, font: "Times New Roman", size: 22 })] })] }),
                                new TableCell({ borders, children: [new Paragraph({ children: [new TextRun({ text: k, font: "Times New Roman", size: 22 })] })] }),
                                new TableCell({ borders, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: s, font: "Times New Roman", size: 22 })] })] }),
                            ]
                        }))
                    ]
                }),
                emptyLine(),
                mixedBody([
                    {text: "Pemilihan aset ini selaras dengan referensi Giudici et al. (2020) yang menjadi landasan penelitian ini."}
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

                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    spacing: { before: 240, after: 240 },
                    children: [
                        new TextRun({ text: "min", font: "Times New Roman", size: 24, italics: true }),
                        new TextRun({ text: "w", font: "Times New Roman", size: 18, italics: true, subScript: true }),
                        new TextRun({ text: " (", font: "Times New Roman", size: 26 }),
                        new TextRun({ text: "w", font: "Times New Roman", size: 24, italics: true }),
                        new TextRun({ text: "T", font: "Times New Roman", size: 18, italics: true, superScript: true }),
                        new TextRun({ text: " \u22C5 S", font: "Times New Roman", size: 24, italics: true }),
                        new TextRun({ text: "f", font: "Times New Roman", size: 18, italics: true, subScript: true }),
                        new TextRun({ text: " \u22C5 w + \u03B3 ", font: "Times New Roman", size: 24, italics: true }),
                        new TextRun({ text: "\u2211", font: "Times New Roman", size: 28 }),
                        new TextRun({ text: " (C", font: "Times New Roman", size: 24, italics: true }),
                        new TextRun({ text: "e", font: "Times New Roman", size: 18, italics: true, subScript: true }),
                        new TextRun({ text: " \u22C5 w)", font: "Times New Roman", size: 24, italics: true }),
                        new TextRun({ text: ")", font: "Times New Roman", size: 26 }),
                        new TextRun({ text: "             (III.1)", font: "Times New Roman", size: 24, bold: true }),
                    ]
                }),
                body("Keterangan:"),
                new Table({
                    width: { size: 100, type: WidthType.PERCENTAGE },
                    borders: {
                        top: { style: BorderStyle.NONE, size: 0 },
                        bottom: { style: BorderStyle.NONE, size: 0 },
                        left: { style: BorderStyle.NONE, size: 0 },
                        right: { style: BorderStyle.NONE, size: 0 },
                        insideHorizontal: { style: BorderStyle.NONE, size: 0 },
                        insideVertical: { style: BorderStyle.NONE, size: 0 },
                    },
                    rows: [
                        new TableRow({
                            children: [
                                new TableCell({ width: { size: 8, type: WidthType.PERCENTAGE }, children: [new Paragraph({ children: [new TextRun({ text: "w", font: "Times New Roman", size: 24, italics: true })] })] }),
                                new TableCell({ width: { size: 4, type: WidthType.PERCENTAGE }, children: [new Paragraph({ children: [new TextRun({ text: "=", font: "Times New Roman", size: 24 })] })] }),
                                new TableCell({ width: { size: 88, type: WidthType.PERCENTAGE }, children: [new Paragraph({ children: [new TextRun({ text: "Vektor alokasi bobot untuk setiap aset kripto (total = 1).", font: "Times New Roman", size: 24 })] })] }),
                            ]
                        }),
                        new TableRow({
                            children: [
                                new TableCell({ width: { size: 8, type: WidthType.PERCENTAGE }, children: [new Paragraph({ children: [new TextRun({ text: "S", font: "Times New Roman", size: 24, italics: true }), new TextRun({ text: "f", font: "Times New Roman", size: 18, italics: true, subScript: true })] })] }),
                                new TableCell({ width: { size: 4, type: WidthType.PERCENTAGE }, children: [new Paragraph({ children: [new TextRun({ text: "=", font: "Times New Roman", size: 24 })] })] }),
                                new TableCell({ width: { size: 88, type: WidthType.PERCENTAGE }, children: [new Paragraph({ children: [new TextRun({ text: "Matriks Kovarians terfilter RMT.", font: "Times New Roman", size: 24 })] })] }),
                            ]
                        }),
                        new TableRow({
                            children: [
                                new TableCell({ width: { size: 8, type: WidthType.PERCENTAGE }, children: [new Paragraph({ children: [new TextRun({ text: "\u03B3", font: "Times New Roman", size: 24, italics: true })] })] }),
                                new TableCell({ width: { size: 4, type: WidthType.PERCENTAGE }, children: [new Paragraph({ children: [new TextRun({ text: "=", font: "Times New Roman", size: 24 })] })] }),
                                new TableCell({ width: { size: 88, type: WidthType.PERCENTAGE }, children: [new Paragraph({ children: [new TextRun({ text: "Parameter skalar penghukuman sentralitas graf.", font: "Times New Roman", size: 24 })] })] }),
                            ]
                        }),
                        new TableRow({
                            children: [
                                new TableCell({ width: { size: 8, type: WidthType.PERCENTAGE }, children: [new Paragraph({ children: [new TextRun({ text: "C", font: "Times New Roman", size: 24, italics: true }), new TextRun({ text: "e", font: "Times New Roman", size: 18, italics: true, subScript: true })] })] }),
                                new TableCell({ width: { size: 4, type: WidthType.PERCENTAGE }, children: [new Paragraph({ children: [new TextRun({ text: "=", font: "Times New Roman", size: 24 })] })] }),
                                new TableCell({ width: { size: 88, type: WidthType.PERCENTAGE }, children: [new Paragraph({ children: [new TextRun({ text: "Vektor skor Eigenvector Centrality tiap node aset.", font: "Times New Roman", size: 24 })] })] }),
                            ]
                        }),
                    ],
                }),
                emptyLine(),
                heading3("3.4.1. Strategi Portofolio dan Benchmark"),
                body("Penelitian ini membandingkan empat strategi utama untuk mengevaluasi performa model yang diusulkan terhadap standar industri dan metodologi mutakhir:"),
                
                bulletItem("Equally Weighted (EW): Strategi alokasi 1/N yang memberikan bobot yang sama ke setiap aset tanpa mempertimbangkan parameter risiko atau imbal hasil. EW berfungsi sebagai 'benchmark naif' yang sangat tangguh karena tidak memiliki risiko estimasi (estimation risk)."),
                
                bulletItem("Classical Markowitz (CM): Model optimasi Mean-Variance standar yang berupaya meminimalkan variansi portofolio untuk tingkat imbal hasil tertentu. CM bertindak sebagai representasi teori portofolio tradisional yang sering kali menderita masalah ketidakstabilan numerik pada data historis yang berisik."),
                
                bulletItem("Graphical Lasso Markowitz (GM): Model yang menggunakan algoritma Lasso pada matriks presisi (invers kovarians) untuk memaksa elemen-elemen korelasi yang tidak signifikan menjadi nol (sparsity). GM dipilih sebagai benchmark karena kemampuannya menangani masalah sparsitas pada data berdimensi tinggi, yang merupakan tantangan utama dalam data kripto yang sangat terkorelasi secara palsu."),
                
                bulletItem("Network Markowitz (NW): Model jaringan original (Giudici et al., 2020) yang menggunakan parameter penalti sentralitas (\u03b3) statis. NW digunakan sebagai pembanding langsung untuk menunjukkan sejauh mana penambahan fitur 'Dynamic Grid Search' pada model yang diusulkan dapat meningkatkan performa portofolio dibandingkan model jaringan dasar."),
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
                heading2("3.5. Matriks Evaluasi Performa"),
                body("Untuk mengukur efektivitas model portofolio yang dikembangkan, digunakan beberapa metrik evaluasi sebagai berikut:"),
                
                heading3("3.5.1. Risk-Adjusted Return (Sharpe Ratio)"),
                mixedBody([
                    {text: "Metrik ini mengukur imbal hasil berlebih per unit risiko (standar deviasi). Formula: SR = (R\u209A - R\u209B) / \u03C3\u209A. Nilai Sharpe Ratio yang lebih tinggi menunjukkan efisiensi portofolio yang lebih baik dalam mengonversi risiko menjadi keuntungan [4], [13]."}
                ]),
                emptyLine(),
                
                heading3("3.5.2. Downside Risk (Value at Risk - VaR)"),
                mixedBody([
                    {text: "Value at Risk (VaR) pada tingkat kepercayaan 95% digunakan untuk mengestimasi potensi kerugian maksimal dalam satu periode perdagangan [16]. Ini sangat relevan untuk menguji ketahanan portofolio terhadap guncangan ("},
                    {text: "black swan event", italic: true},
                    {text: ") di pasar kripto [11]."}
                ]),
                emptyLine(),
                
                heading3("3.5.3. Tail Risk & Reward (Rachev Ratio)"),
                mixedBody([
                    {text: "Rachev Ratio digunakan untuk mengukur asimetri antara potensi imbal hasil ekstrem ("},
                    {text: "upper tail", italic: true},
                    {text: ") dan potensi kerugian ekstrem ("},
                    {text: "lower tail", italic: true},
                    {text: "). Metrik ini jauh lebih sensitif terhadap karakteristik "},
                    {text: "fat-tail", italic: true},
                    {text: " pada distribusi aset kripto dibandingkan Sharpe Ratio konvensional [17]."}
                ]),

                emptyLine(),
                heading2("3.6. Rencana Jadwal Penelitian"),
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [new TextRun({ text: "Tabel III.2. Rencana Jadwal Penelitian", font: "Times New Roman", size: 24, bold: true })]
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
                    margin: { top: 1701, right: 1417, bottom: 1417, left: 2268 }
                },
                type: SectionType.NEXT_PAGE
            },
            // Lanjutkan penomoran decimal dari section sebelumnya
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
                new Paragraph({
                    alignment: AlignmentType.JUSTIFIED,
                    spacing: { before: 0, after: 120, line: 360 },
                    indent: { left: 720, hanging: 720 },
                    children: [new TextRun({ text: "[6] Z. Jing and J. G. Rocha, \"A network-based strategy of price correlations for optimal cryptocurrency portfolios,\" Financial Innovation, vol. 9, no. 1, pp. 1–28, 2023.", font: "Times New Roman", size: 24 })]
                }),
                new Paragraph({
                    alignment: AlignmentType.JUSTIFIED,
                    spacing: { before: 0, after: 120, line: 360 },
                    indent: { left: 720, hanging: 720 },
                    children: [new TextRun({ text: "[7] I. Kitanovski et al., \"Network-based diversification of stock and cryptocurrency portfolios,\" Expert Systems with Applications, 2024.", font: "Times New Roman", size: 24 })]
                }),
                new Paragraph({
                    alignment: AlignmentType.JUSTIFIED,
                    spacing: { before: 0, after: 120, line: 360 },
                    indent: { left: 720, hanging: 720 },
                    children: [new TextRun({ text: "[8] I. Kitanovski et al., \"Cryptocurrency Portfolio Diversification Using Network Community Detection,\" Algorithms, 2022.", font: "Times New Roman", size: 24 })]
                }),
                new Paragraph({
                    alignment: AlignmentType.JUSTIFIED,
                    spacing: { before: 0, after: 120, line: 360 },
                    indent: { left: 720, hanging: 720 },
                    children: [new TextRun({ text: "[9] P. Giudici and G. Policardo, \"Network models to improve robot advisory portfolios,\" Statistica Applicata - Italian Journal of Applied Statistics, 2021.", font: "Times New Roman", size: 24 })]
                }),
                new Paragraph({
                    alignment: AlignmentType.JUSTIFIED,
                    spacing: { before: 0, after: 120, line: 360 },
                    indent: { left: 720, hanging: 720 },
                    children: [new TextRun({ text: "[10] L. Laloux, P. Cizeau, J. P. Bouchaud, and M. Potters, \"Noise dressing of financial correlation matrices,\" Physical Review Letters, vol. 83, no. 7, pp. 1467, 1999.", font: "Times New Roman", size: 24 })]
                }),
                new Paragraph({
                    alignment: AlignmentType.JUSTIFIED,
                    spacing: { before: 0, after: 120, line: 360 },
                    indent: { left: 720, hanging: 720 },
                    children: [new TextRun({ text: "[11] S. Corbet, B. Lucey, A. Urquhart, and L. Yarovaya, \"Cryptocurrencies as a financial asset: A systematic analysis,\" International Review of Financial Analysis, vol. 62, pp. 182-199, 2019.", font: "Times New Roman", size: 24 })]
                }),
                new Paragraph({
                    alignment: AlignmentType.JUSTIFIED,
                    spacing: { before: 0, after: 120, line: 360 },
                    indent: { left: 720, hanging: 720 },
                    children: [new TextRun({ text: "[12] G. Peralta and A. Zaresei, \"A network approach to portfolio selection,\" Journal of Network Theory in Finance, vol. 2, no. 4, pp. 1-20, 2016.", font: "Times New Roman", size: 24 })]
                }),
                new Paragraph({
                    alignment: AlignmentType.JUSTIFIED,
                    spacing: { before: 0, after: 120, line: 360 },
                    indent: { left: 720, hanging: 720 },
                    children: [new TextRun({ text: "[13] M. Lopez de Prado, \"Advances in Financial Machine Learning,\" John Wiley & Sons, 2018.", font: "Times New Roman", size: 24 })]
                }),
                new Paragraph({
                    alignment: AlignmentType.JUSTIFIED,
                    spacing: { before: 0, after: 120, line: 360 },
                    indent: { left: 720, hanging: 720 },
                    children: [new TextRun({ text: "[14] R. O. Michaud, \"The Markowitz Optimization Enigma: Is 'Optimized' Optimal?\" Financial Analysts Journal, vol. 45, no. 1, pp. 31-42, 1989.", font: "Times New Roman", size: 24 })]
                }),
                new Paragraph({
                    alignment: AlignmentType.JUSTIFIED,
                    spacing: { before: 0, after: 120, line: 360 },
                    indent: { left: 720, hanging: 720 },
                    children: [new TextRun({ text: "[15] C. Eom, G. Oh, S. Jung, H. Jeong, and S. Kim, \"Topological properties of stock networks based on minimal spanning tree and random matrix theory,\" Physica A: Statistical Mechanics and its Applications, vol. 388, no. 6, pp. 900-906, 2009.", font: "Times New Roman", size: 24 })]
                }),
                new Paragraph({
                    alignment: AlignmentType.JUSTIFIED,
                    spacing: { before: 0, after: 120, line: 360 },
                    indent: { left: 720, hanging: 720 },
                    children: [new TextRun({ text: "[16] P. Jorion, \"Value at Risk: The New Benchmark for Managing Financial Risk,\" McGraw-Hill, 2000.", font: "Times New Roman", size: 24 })]
                }),
                new Paragraph({
                    alignment: AlignmentType.JUSTIFIED,
                    spacing: { before: 0, after: 120, line: 360 },
                    indent: { left: 720, hanging: 720 },
                    children: [new TextRun({ text: "[17] A. Biglova, S. Ortobelli, S. T. Rachev, and S. V. Stoyanov, \"Different Approaches to Risk Estimation in Portfolio Theory,\" The Journal of Portfolio Management, vol. 31, no. 1, pp. 103-112, 2004.", font: "Times New Roman", size: 24 })]
                }),
                new Paragraph({
                    alignment: AlignmentType.JUSTIFIED,
                    spacing: { before: 0, after: 120, line: 360 },
                    indent: { left: 720, hanging: 720 },
                    children: [new TextRun({ text: "[18] A. Arratia, \"Computational Finance: An Introductory Course with R,\" Atlantis Press, 2014.", font: "Times New Roman", size: 24 })]
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
