const docx = require('docx');
const {
    Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
    AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
    LevelFormat, PageNumber, PageBreak, TabStopType, TabStopPosition,
    VerticalAlign, ImageRun, Header, Footer, NumberFormat, SectionType,
    TableOfContents,
    Math, MathRun, MathSubScript, MathSuperScript, MathFraction, MathSum, MathRoundBrackets, MathLimitLower
} = docx;

/**
 * Helper sederhana untuk mengonversi string LaTeX dasar ke komponen docx Math.
 * Mendukung: \min, \sum, \frac, simbol yunani (\Sigma, \gamma, \alpha, \dots), subskrip, dan superskrip.
 */
function rumus(latexStr) {
    // Note: Ini adalah parser sederhana khusus untuk kebutuhan proposal ini.
    // Jika ingin parser LaTeX full, biasanya diperlukan pustaka tambahan.
    
    // Penanganan khusus untuk rumus Network Markowitz (8)
    if (latexStr.includes("\\min_w") && latexStr.includes("\\Sigma^*")) {
        return [
            new MathLimitLower({
                children: [new MathRun("min")],
                limit: [new MathRun("w")],
            }),
            new MathRun(" "),
            new MathSuperScript({
                children: [new MathRun("w")],
                superScript: [new MathRun("T")],
            }),
            new MathSuperScript({
                children: [new MathRun("\u03A3")],
                superScript: [new MathRun("*")],
            }),
            new MathRun("w + \u03B3"),
            new MathSum({
                subScript: [new MathRun("i=1")],
                superScript: [new MathRun("n")],
                children: [
                    new MathSubScript({ children: [new MathRun("x")], subScript: [new MathRun("i")] }),
                    new MathSubScript({ children: [new MathRun("w")], subScript: [new MathRun("i")] })
                ]
            })
        ];
    }

    // Penanganan untuk Sharpe Ratio
    if (latexStr.includes("Sharpe\\ Ratio")) {
        return [
            new MathRun("Sharpe Ratio = "),
            new MathFraction({
                numerator: [
                    new MathSubScript({
                        children: [new MathRun("R")],
                        subScript: [new MathRun("p")]
                    }),
                    new MathRun(" - "),
                    new MathSubScript({
                        children: [new MathRun("R")],
                        subScript: [new MathRun("f")]
                    })
                ],
                denominator: [
                    new MathSubScript({
                        children: [new MathRun("\u03C3")], // Simbol Sigma
                        subScript: [new MathRun("p")]
                    })
                ]
            })
        ];
    }

    // Penanganan untuk VaR
    if (latexStr.includes("VaR")) {
        return [
            new MathSubScript({ children: [new MathRun("VaR")], subScript: [new MathRun("\u03B1")] }),
            new MathRoundBrackets({ children: [new MathRun("\u03B1")] }),
            new MathRun(" = -inf"),
            new MathRoundBrackets({ children: [new MathRun("x \u2208 \u211D : P(L > x) \u2264 1 - \u03B1")] })
        ];
    }

    // Penanganan untuk Rachev Ratio
    if (latexStr.includes("RR")) {
        return [
            new MathRun("RR = "),
            new MathFraction({
                numerator: [
                    new MathSubScript({ children: [new MathRun("ETR")], subScript: [new MathRun("\u03B1")] }),
                    new MathRoundBrackets({ children: [new MathRun("R")] })
                ],
                denominator: [
                    new MathSubScript({ children: [new MathRun("ES")], subScript: [new MathRun("\u03B2")] }),
                    new MathRoundBrackets({ children: [new MathRun("R")] })
                ]
            })
        ];
    }

    return [new MathRun(latexStr)];
}


// Fallback for TabLeader which might be named differently or missing in some versions
const TabLeader = docx.TabLeader || docx.TabStopLeader || { DOT: "dot" };
const fs = require('fs');

const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };
const noBorder = { style: BorderStyle.NONE, size: 0, color: "auto" };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder, insideHorizontal: noBorder, insideVertical: noBorder };

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
        children: [new TextRun({ text, font: "Times New Roman", size: 24, bold: true })]
    });
}

function heading2(text) {
    return new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun({ text, font: "Times New Roman", size: 24, bold: true })]
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

function chapterHeading(bab, title, size = 24) {
    return new Paragraph({
        heading: HeadingLevel.HEADING_1,
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 120 },
        children: [
            new TextRun({ text: bab, font: "Times New Roman", size, bold: true }),
            new TextRun({ break: 1 }),
            new TextRun({ text: title, font: "Times New Roman", size, bold: true })
        ]
    });
}

function sectionTitle(text, size = 24) {
    return new Paragraph({
        heading: HeadingLevel.HEADING_1,
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
                run: { size: 24, bold: true, font: "Times New Roman" },
                paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 0 }
            },
            {
                id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
                run: { size: 24, bold: true, font: "Times New Roman" },
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
                emptyLine(),
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                        new ImageRun({
                            data: fs.readFileSync("logo_unm.png"),
                            extension: "png",
                            transformation: {
                                width: 150,
                                height: 150,
                            },
                        })
                    ]
                }),
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
                type: SectionType.NEXT_PAGE, titlePage: true,
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
                sectionTitle("HALAMAN PENGESAHAN"),
                emptyLine(),
                centeredBold("PROPOSAL TESIS", 24),
                emptyLine(),
                emptyLine(),
                bodyNoIndent("Proposal tesis ini diajukan oleh:", { bold: true }),
                emptyLine(),
                ...[
                    ["Nama",          "Ragil Yulianto"],
                    ["NIM",           "14240007"],
                    ["Program Studi", "Ilmu Komputer"],
                    ["Fakultas",      "Teknologi Informasi"],
                    ["Jenjang",       "Strata Dua (S2)"],
                    ["Judul Tesis",   "OPTIMASI DINAMIS PEMODELAN NETWORK MARKOWITZ UNTUK MANAJEMEN PORTOFOLIO MATA UANG KRIPTO"],
                ].map(([label, value]) => new Paragraph({
                    indent: { left: 720 },
                    spacing: { line: 360 },
                    tabStops: [{ type: TabStopType.LEFT, position: 2500 }],
                    children: [
                        new TextRun({ text: label, font: "Times New Roman", size: 24 }),
                        new TextRun({ text: "\t: " + value, font: "Times New Roman", size: 24, bold: true }),
                    ]
                })),
                emptyLine(),
                new Paragraph({
                    alignment: AlignmentType.LEFT,
                    children: [new TextRun({ text: "telah diperiksa dan disetujui untuk diajukan sebagai rencana pelaksanaan penelitian", font: "Times New Roman", size: 24 })]
                }),
                emptyLine(),
                new Paragraph({
                    alignment: AlignmentType.CENTER,
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
                type: SectionType.NEXT_PAGE, titlePage: true,
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
                new TableOfContents("Daftar Isi", {
                    hyperlink: true,
                    headingStyleRange: "1-3",
                    caption: { text: "Daftar Isi" },
                }),
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
                type: SectionType.NEXT_PAGE, titlePage: true,
                pageNumbers: { start: 1, formatType: NumberFormat.DECIMAL }
            },
            headers: {
                default: new Header({
                    children: [
                        new Paragraph({
                            alignment: AlignmentType.RIGHT,
                            children: [new TextRun({ children: [PageNumber.CURRENT], font: "Times New Roman", size: 24 })]
                        })
                    ]
                })
            },
            footers: {
                first: new Footer({
                    children: [
                        new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [new TextRun({ children: [PageNumber.CURRENT], font: "Times New Roman", size: 24 })]
                        })
                    ]
                })
            },
            children: [
                chapterHeading("BAB I", "PENDAHULUAN"),
                emptyLine(),
                body("Bab ini membahas secara komprehensif latar belakang permasalahan, perumusan penelitian, tujuan, dan ruang lingkup yang menjadi batasan penelitian ini."),
                emptyLine(),
                heading2("1.1 Latar Belakang"),
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
                        new TextRun({ text: " [mendeley_cite:giudici2020network], [mendeley_cite:laloux1999noise], [mendeley_cite:corbet2019cryptocurrencies].", font: "Times New Roman", size: 24 }),
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
                        new TextRun({ text: ") ke dalam optimasi portofolio [mendeley_cite:giudici2020network], [mendeley_cite:momeni2021portfolio]. Penggunaan instrumen seperti ", font: "Times New Roman", size: 24 }),
                        new TextRun({ text: "Minimum Spanning Tree (MST)", font: "Times New Roman", size: 24, italics: true }),
                        new TextRun({ text: " dan ", font: "Times New Roman", size: 24 }),
                        new TextRun({ text: "Eigenvector Centrality", font: "Times New Roman", size: 24, italics: true }),
                        new TextRun({ text: " terbukti efisien dalam memetakan interaksi antar aset dan menghukum (", font: "Times New Roman", size: 24 }),
                        new TextRun({ text: "penalize", font: "Times New Roman", size: 24, italics: true }),
                        new TextRun({ text: ") aset-aset yang menjadi titik pusat kegagalan sistemik [mendeley_cite:mantegna1999hierarchical], [mendeley_cite:jing2023network], [mendeley_cite:peralta2016network].", font: "Times New Roman", size: 24 }),
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
                        new TextRun({ text: " [mendeley_cite:marchenko1967distribution], [mendeley_cite:laloux1999noise] disandingkan dengan optimalisasi ", font: "Times New Roman", size: 24 }),
                        new TextRun({ text: "Grid Search", font: "Times New Roman", size: 24, italics: true }),
                        new TextRun({ text: " dinamis (", font: "Times New Roman", size: 24 }),
                        new TextRun({ text: "rolling window", font: "Times New Roman", size: 24, italics: true }),
                        new TextRun({ text: ") agar portofolio dapat membentengi aset di saat ", font: "Times New Roman", size: 24 }),
                        new TextRun({ text: "crypto winter", font: "Times New Roman", size: 24, italics: true }),
                        new TextRun({ text: " tanpa mengorbankan rasio ", font: "Times New Roman", size: 24 }),
                        new TextRun({ text: "upside gain", font: "Times New Roman", size: 24, italics: true }),
                        new TextRun({ text: " di saat pembalikan arah (", font: "Times New Roman", size: 24 }),
                        new TextRun({ text: "bullish", font: "Times New Roman", size: 24, italics: true }),
                        new TextRun({ text: ") [mendeley_cite:giudici2020network], [mendeley_cite:lopezdeprado2018advances].", font: "Times New Roman", size: 24 }),
                    ]
                }),
                emptyLine(),
                heading2("1.2 Identifikasi Masalah"),
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
                heading2("1.3 Tujuan Penelitian"),
                body("Tujuan dari penelitian ini adalah:"),
                letterItem("Menganalisis keandalan metodologi Network Markowitz (dengan integrasi RMT filter dan Eigenvector Centrality) dalam menekan ekstrimitas downside risk dibandingkan pendekatan portofolio naif dan konvensional.", "letters1"),
                letterItem("Merancang dan menguji model Network Markowitz adaptif (Grid Search Optimization) yang mampu melakukan re-kalibrasi dinamis dengan menggunakan paradigma rolling window pada berbagai lanskap pasar (Bearish, Recovery, Stable).", "letters1"),
                letterItem("Membandingkan performa perlindungan risiko sistemik (VaR) dan asimetri imbal hasil (Rachev Ratio) antara pemodelan baru dengan metode-metode baseline pada reksadana aset kripto.", "letters1"),
                emptyLine(),
                heading2("1.4 Ruang Lingkup Penelitian"),
                body("Ruang lingkup penelitian ini dibatasi pada:"),
                letterItem("Objek penelitian terfokus pada data fluktuasi harga harian dari 10 (sepuluh) aset kripto utama dalam kerangka waktu historis termasuk masa resesi crypto winter (14 September 2017 hingga 17 Oktober 2019).", "letters2"),
                letterItem("Metode yang dibandingkan secara teknis mencakup Equally Weighted (EW), Classical Markowitz (CM), Glasso Markowitz (GM), Network Markowitz statis (\u03b3 = 0, 1.0, 2.0), serta Optimized Network Markowitz secara dinamis berbasis Grid Search.", "letters2"),
                letterItem("Pengujian (backtesting) dilakukan dalam out-of-sample rolling window (120 observasi ke belakang dengan frekuensi penyesuaian rebalance 7 hari) yang disimulasikan menggunakan transaction cost atau estimasi biaya bursa (0.1%).", "letters2"),
                emptyLine(),
                heading2("1.5 Sistematika Penulisan"),
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
            headers: {
                default: new Header({
                    children: [
                        new Paragraph({
                            alignment: AlignmentType.RIGHT,
                            children: [new TextRun({ children: [PageNumber.CURRENT], font: "Times New Roman", size: 24 })]
                        })
                    ]
                })
            },
            footers: {
                first: new Footer({
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
                type: SectionType.NEXT_PAGE, titlePage: true
            },
            // Lanjutkan penomoran decimal dari section sebelumnya
            children: [
                chapterHeading("BAB II", "LANDASAN/KERANGKA PEMIKIRAN"),
                emptyLine(),
                body("Bab ini menguraikan berbagai teori dasar, konsep, serta tinjauan pustaka dari penelitian terdahulu yang menjadi landasan pemikiran bagi pengembangan sistem portofolio dalam penelitian ini."),
                emptyLine(),
                heading2("2.1 Modern Portfolio Theory (Markowitz)"),
                mixedBody([
                    {text: "Teori Portofolio Modern (MPT), yang dipelopori oleh Harry Markowitz, berupaya memaksimalkan imbal hasil yang diharapkan ("},
                    {text: "expected return", italic: true},
                    {text: ") pada tingkat risiko ("},
                    {text: "variance", italic: true},
                    {text: ") tertentu melalui pemilihan aset yang terdiversifikasi dalam sebuah "},
                    {text: "Efficient Frontier", italic: true},
                    {text: " [mendeley_cite:markowitz1952portfolio]. Namun, dalam praktiknya, MPT memiliki keterbatasan signifikan terkait stabilitas estimasi. Masalah mendasarnya adalah bahwa kovarians dari kumpulan aset finansial sangat sensitif terhadap nilai-nilai ekstrem historis, yang dikenal sebagai "},
                    {text: "Markowitz Curse", italic: true},
                    {text: " atau enigma optimasi Markowitz [mendeley_cite:michaud1989markowitz]. Fenomena ini terjadi karena algoritma optimasi cenderung memperkuat kesalahan estimasi ("},
                    {text: "error maximization", italic: true},
                    {text: "), sehingga perubahan kecil pada input data dapat menghasilkan perubahan drastis pada alokasi bobot portofolio."}
                ]),
                emptyLine(),
                heading2("2.2 Random Matrix Theory dan Distribusi Marchenko-Pastur"),
                mixedBody([
                    {text: "Teori "},
                    {text: "Random Matrix", italic: true},
                    {text: " (RMT) digunakan untuk memisahkan korelasi yang mengandung informasi ekonomi sejati dari "},
                    {text: "noise", italic: true},
                    {text: " statistik pada matriks korelasi berdimensi tinggi. Inti dari filtrasi RMT terletak pada distribusi Marchenko-Pastur [mendeley_cite:marchenko1967distribution], yang mendefinisikan batas teoritis nilai eigen ("},
                    {text: "eigenvalues", italic: true},
                    {text: ") dari matriks korelasi acak sebagai \u03bb\u208A = \u03c3\u00b2(1 + \u221aq/N)\u00b2. Nilai eigen yang melampaui batas \u03bb\u208A merepresentasikan sinyal pasar kolektif, sementara nilai eigen di bawahnya dianggap sebagai residu yang harus dibersihkan agar stabilisasi topologi jaringan (MST) dapat tercapai secara konsisten [mendeley_cite:marchenko1967distribution], [mendeley_cite:eom2009topological]."}
                ]),
                emptyLine(),
                heading2("2.3 Teori Risiko Koheren (Coherent Risk Measures)"),
                mixedBody([
                    {text: "Penggunaan metrik risiko dalam optimasi portofolio harus memenuhi kriteria risiko koheren sebagaimana didefinisikan oleh Artzner et al. [mendeley_cite:artzner1999coherent]. Kriteria tersebut mencakup empat aksioma: "},
                    {text: "monotonicity, sub-additivity, homogeneity,", italic: true},
                    {text: " dan "},
                    {text: "translational invariance", italic: true},
                    {text: ". Berbeda dengan variansi pada model Markowitz klasik yang gagal memenuhi aksioma "},
                    {text: "sub-additivity", italic: true},
                    {text: " pada distribusi tidak normal, penggunaan metrik seperti "},
                    {text: "Expected Shortfall", italic: true},
                    {text: " atau proksi risiko ekor seperti "},
                    {text: "Rachev Ratio", italic: true},
                    {text: " dalam model ini memberikan perlindungan yang lebih kuat terhadap kejadian ekstrem ("},
                    {text: "fat-tail events", italic: true},
                    {text: ") di pasar kripto."}
                ]),
                bulletItem("Monotonicity: Jika portofolio X selalu tidak lebih buruk dari Y, maka risiko X harus lebih kecil atau sama dari Y (\u03c1(X) \u2264 \u03c1(Y) untuk X \u2265 Y)."),
                bulletItem("Sub-additivity: Risiko gabungan tidak boleh lebih dari jumlah risiko masing-masing (\u03c1(X + Y) \u2264 \u03c1(X) + \u03c1(Y)). Ini adalah kaidah efek diversifikasi."),
                bulletItem("Homogeneity: Menambah kelipatan ukuran posisi sejalan dengan mengalikan besaran risikonya (\u03c1(cX) = c \u03c1(X) untuk c > 0)."),
                bulletItem("Translational Invariance: Menambah sejumlah modal pasti bebas risiko ke portofolio akan mengurangi risiko sebesar persis nilai nominal tersebut."),
                emptyLine(),
                heading2("2.4 Topologi Jaringan Keuangan dan Risiko Penularan"),
                mixedBody([
                    {text: "Dalam perspektif teori jaringan, pasar kripto dapat dipetakan menjadi struktur topologi tertentu. Struktur "},
                    {text: "Star-like", italic: true},
                    {text: " yang didominasi oleh aset sentral (seperti Bitcoin) menunjukkan ketergantungan sistemik yang tinggi, di mana gejolak pada pusat jaringan akan dengan cepat menyebar melintasi jaringan ("},
                    {text: "financial contagion", italic: true},
                    {text: "). Fenomena ini seringkali memicu kegagalan beruntun ("},
                    {text: "cascading failures", italic: true},
                    {text: "), di mana likuidasi pada satu titik menyebar ke seluruh ekosistem akibat korelasi yang tinggi. Sebaliknya, struktur "},
                    {text: "Distributed", italic: true},
                    {text: " menawarkan manfaat diversifikasi yang lebih baik. Dengan menggunakan sentralitas sebagai penalti, model Network Markowitz secara efektif menggeser alokasi dari 'pusat penularan' ke 'periferi jaringan', sehingga memitigasi risiko kegagalan sistemik [mendeley_cite:mantegna1999hierarchical], [mendeley_cite:peralta2016network]."}
                ]),
                emptyLine(),
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                        new TextRun({
                            text: "[[%IMAGE_TOPOLOGY]]",
                            font: "Times New Roman",
                            size: 22,
                        })
                    ]
                }),
                emptyLine(),
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                        new TextRun({ text: "Gambar II.1. Ilustrasi Topologi Finansial Star-like vs Distributed", font: "Times New Roman", size: 22, bold: true })
                    ]
                }),
                emptyLine(),
                heading2("2.5 Network Markowitz"),
                mixedBody([
                    {text: "Integrasi variabel jaringan ke dalam fungsi objektif Markowitz memungkinkan model untuk 'menghukum' aset yang memiliki tingkat keterhubungan sistemik tinggi. Formula optimasi dimodifikasi dengan menambahkan faktor \u03b3 yang dikalikan dengan skor sentralitas vektor eigen [mendeley_cite:giudici2020network], [mendeley_cite:giudici2021network]. Hal ini memastikan bahwa portofolio tidak hanya efisien secara variansi-imbal hasil, tetapi juga tangguh terhadap dinamika struktur jaringan pasar kripto yang bervariasi mengikuti transisi fasa."}
                ]),
                emptyLine(),
                heading2("2.6 Teori Siklus dan Adaptive Market Hypothesis (AMH)"),
                mixedBody([
                    {text: "Kelemahan efisiensi pasar kripto dijelaskan melalui "},
                    {text: "Adaptive Market Hypothesis", italic: true},
                    {text: " (AMH) oleh Andrew Lo [mendeley_cite:lo2004adaptive]. AMH menyatakan bahwa efisiensi pasar bukanlah kondisi statis, melainkan hasil adaptasi pelaku pasar terhadap perubahan lingkungan. Hal ini memberikan landasan teoritis kuat bagi penggunaan metode "},
                    {text: "Rolling Window", italic: true},
                    {text: " dan "},
                    {text: "Grid Search", italic: true},
                    {text: " dalam penelitian ini; karena korelasi dan risiko aset kripto terus berevolusi melalui fase "},
                    {text: "Bearish, Recovery,", italic: true},
                    {text: " dan "},
                    {text: "Bullish", italic: true},
                    {text: ", maka parameter penalti portofolio (\u03b3) harus dikalibrasi secara dinamis untuk mencapai performa optimal."}
                ]),
                emptyLine(),
                heading2("2.7 Walk-forward Analysis dan Filosofi Rolling Window"),
                mixedBody([
                    {text: "Walk-forward Analysis merupakan teknik validasi utama dalam Machine Learning finansial untuk menghindari "},
                    {text: "look-ahead bias", italic: true},
                    {text: ". Berbeda dengan "},
                    {text: "cross-validation", italic: true},
                    {text: " tradisional yang mengabaikan urutan waktu, metode "},
                    {text: "rolling window", italic: true},
                    {text: " memastikan bahwa pengujian model dilakukan menggunakan data yang secara kronologis berada setelah data pelatihan. Pendekatan ini memberikan kepastian bahwa optimalisasi parameter \u03b3 pada setiap jendela waktu dilakukan dengan integritas data yang tinggi, sehingga hasil "},
                    {text: "backtesting", italic: true},
                    {text: " mencerminkan realitas perdagangan sesungguhnya di pasar yang sangat dinamis."}
                ]),
                emptyLine(),
                heading2("2.8 Analisis Kebaruan (Gap Analysis)"),
                mixedBody([
                    {text: "Penelitian ini memiliki kebaruan signifikan dibandingkan model yang diusulkan oleh Giudici et al. [mendeley_cite:giudici2020network]. Jika penelitian tersebut menggunakan parameter penghukuman jaringan (\u03b3) yang bernilai statis (konstan), penelitian ini mengusulkan "},
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
                heading2("2.9 Penelitian Terdahulu"),
                body("Beberapa penelitian terdahulu yang relevan dengan penelitian ini antara lain:"),
                mixedBody([
                    {text: "Penelitian yang dilakukan oleh Giudici dkk. (2020) mengusulkan kerangka kerja baru untuk manajemen portofolio kripto otomatis dengan memperluas model Markowitz tradisional melalui integrasi "},
                    {text: "Random Matrix Theory", italic: true},
                    {text: " (RMT) dan ukuran jaringan ("},
                    {text: "network measures", italic: true},
                    {text: "), seperti "},
                    {text: "centrality", italic: true},
                    {text: " dan "},
                    {text: "Minimal Spanning Tree", italic: true},
                    {text: " (MST). Metodologi ini bertujuan untuk meningkatkan profil risiko-imbal hasil pada instrumen keuangan yang sangat volatil seperti mata uang kripto. Hasil penelitian menunjukkan bahwa model yang menggabungkan RMT, MST, dan ukuran "},
                    {text: "centrality", italic: true},
                    {text: " secara konsisten mengungguli strategi alokasi aset lainnya, baik dalam kondisi pasar "},
                    {text: "bullish", italic: true},
                    {text: " maupun "},
                    {text: "bearish", italic: true},
                    {text: ". Model ini mampu memberikan perlindungan yang lebih baik terhadap kerugian signifikan selama penurunan pasar, sehingga menjadikannya strategi alokasi yang efisien dan adaptif untuk diintegrasikan dalam sistem "},
                    {text: "robo-advisory", italic: true},
                    {text: " guna mendukung konsultasi keuangan otomatis [mendeley_cite:giudici2020network]."}
                ]),
                mixedBody([
                    {text: "Penelitian Kitanovski et al. [mendeley_cite:kitanovski2022cryptocurrency] mengeksplorasi diversifikasi portofolio mata uang kripto dengan memanfaatkan metode deteksi komunitas jaringan, seperti "},
                    {text: "Louvain", italic: true},
                    {text: " dan "},
                    {text: "Affinity Propagation", italic: true},
                    {text: ". Dengan mengelompokkan aset kripto berdasarkan korelasi harga, portofolio dibentuk dengan memilih perwakilan dari masing-masing komunitas yang berbeda. Pendekatan tersebut secara signifikan terbukti membantu mengurangi volatilitas keseluruhan portofolio dan mengoptimalkan tingkat pengembalian ("},
                    {text: "return", italic: true},
                    {text: ")."}
                ]),
                mixedBody([
                    {text: "Pada penelitian lainnya, Jing dan Rocha [mendeley_cite:jing2023network] merancang strategi portofolio kripto optimal dengan menggabungkan "},
                    {text: "Minimum Spanning Tree", italic: true},
                    {text: " (MST) bersama pemodelan "},
                    {text: "Modern Portfolio Theory", italic: true},
                    {text: " (MPT). Pemilihan aset pada ekosistem kripto ini didasarkan pada prinsip maksimalisasi dekorelasi dalam keterhubungan jaringan. Studi tersebut membuktikan bahwa portofolio berbasis MST mampu secara utuh mengungguli seluruh tolok ukur ("},
                    {text: "benchmark", italic: true},
                    {text: ") investasi lainnya, baik itu performa Bitcoin tunggal (BTC), portofolio 5 kripto tertinggi (TOP5), maupun pemilihan portofolio secara acak (RAND)."}
                ]),
                mixedBody([
                    {text: "Terkait ketahanan portofolio menghadapi guncangan harga ekstrem, Kitanovski et al. [mendeley_cite:kitanovski2024network] kembali memperlihatkan keunggulan strategi diversifikasi berbasis topologi jaringan pada kombinasi dua keranjang aset berisiko tinggi; saham dan kripto. Riset ini menyimpulkan bahwa portofolio berbasis konektivitas metrik graf jauh lebih "},
                    {text: "resilient", italic: true},
                    {text: " meredam ancaman kerugian dibandingkan indeks alokasi tradisional, terutama sangat efektif memberikan proteksi perlindungan "},
                    {text: "drawdown", italic: true},
                    {text: " pada saat-saat terjadinya volatilitas parah akibat krisis global (fase pandemi dan perang)."}
                ]),
                mixedBody([
                    {text: "Lebih lanjut, Jing et al. [mendeley_cite:jing2025optimising] memperdalam integrasi antara analisis jaringan dan "},
                    {text: "Modern Portfolio Theory", italic: true},
                    {text: " (MPT) dengan memperkenalkan kerangka kerja teknis yang memanfaatkan "},
                    {text: "Louvain network community algorithm", italic: true},
                    {text: " dan "},
                    {text: "consensus clustering", italic: true},
                    {text: ". Pendekatan ini bertujuan untuk mendeteksi klaster mata uang kripto yang memiliki korelasi tinggi namun secara temporal stabil, dari mana pemilihan aset dilakukan. Penelitian ini juga mengintegrasikan model ARIMA untuk prediksi harga guna menjamin performa portofolio dalam cakrawala investasi jangka pendek (sampai 14 hari). Hasil analisis empiris selama periode 5 tahun menunjukkan bahwa pola harga tersembunyi dapat dimanfaatkan secara efektif melalui struktur jaringan untuk menghasilkan portofolio kripto yang menguntungkan secara konsisten meskipun di tengah volatilitas pasar yang ekstrem."}
                ]),
                emptyLine(),
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [new TextRun({ text: "Tabel II.1. Perbandingan Penelitian Terdahulu", font: "Times New Roman", size: 24, bold: true })]
                }),
                new Table({
                    alignment: AlignmentType.CENTER,
                    width: { size: 8200, type: WidthType.DXA },
                    columnWidths: [400, 1600, 2000, 4200],
                    rows: [
                        new TableRow({
                            tableHeader: true,
                            children: [
                                new TableCell({
                                    borders,
                                    shading: { fill: "D5E8F0", type: ShadingType.CLEAR },
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    verticalAlign: VerticalAlign.CENTER,
                                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "No", font: "Times New Roman", size: 22, bold: true })] })]
                                }),
                                new TableCell({
                                    borders,
                                    shading: { fill: "D5E8F0", type: ShadingType.CLEAR },
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Penulis / Tahun", font: "Times New Roman", size: 22, bold: true })] })]
                                }),
                                new TableCell({
                                    borders,
                                    shading: { fill: "D5E8F0", type: ShadingType.CLEAR },
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Judul", font: "Times New Roman", size: 22, bold: true })] })]
                                }),
                                new TableCell({
                                    borders,
                                    shading: { fill: "D5E8F0", type: ShadingType.CLEAR },
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Hasil", font: "Times New Roman", size: 22, bold: true })] })]
                                }),
                            ]
                        }),
                        new TableRow({
                            children: [
                                new TableCell({
                                    borders,
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    children: [new Paragraph({ alignment: AlignmentType.LEFT, children: [new TextRun({ text: "1", font: "Times New Roman", size: 22 })] })]
                                }),
                                new TableCell({
                                    borders,
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    children: [new Paragraph({ children: [new TextRun({ text: "Giudici, et al. [mendeley_cite:giudici2020network]", font: "Times New Roman", size: 22 })] })]
                                }),
                                new TableCell({
                                    borders,
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    children: [
                                        new Paragraph({ children: [new TextRun({ text: "Network Models to Enhance Automated Cryptocurrency Portfolio Management", font: "Times New Roman", size: 22 })] }),
                                        new Paragraph({ children: [new TextRun({ text: "DOI: 10.3389/frai.2020.00022", font: "Times New Roman", size: 20, italics: true })] })
                                    ]
                                }),
                                new TableCell({
                                    borders,
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    children: [new Paragraph({ alignment: AlignmentType.LEFT, children: [new TextRun({ text: "Mengusulkan Network Markowitz dan sukses mendemonstrasikan perbaikan struktur dibandingkan Markowitz biasa di era crypto winter.", font: "Times New Roman", size: 22 })] })]
                                }),
                            ]
                        }),
                        new TableRow({
                            children: [
                                new TableCell({
                                    borders,
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    children: [new Paragraph({ alignment: AlignmentType.LEFT, children: [new TextRun({ text: "2", font: "Times New Roman", size: 22 })] })]
                                }),
                                new TableCell({
                                    borders,
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    children: [new Paragraph({ children: [new TextRun({ text: "Kitanovski, et al. [mendeley_cite:kitanovski2022cryptocurrency]", font: "Times New Roman", size: 22 })] })]
                                }),
                                new TableCell({
                                    borders,
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    children: [
                                        new Paragraph({ children: [new TextRun({ text: "Cryptocurrency Portfolio Diversification Using Network Community Detection", font: "Times New Roman", size: 22 })] }),
                                        new Paragraph({ children: [new TextRun({ text: "DOI: 10.1109/TELFOR56187.2022.9983742", font: "Times New Roman", size: 20, italics: true })] })
                                    ]
                                }),
                                new TableCell({
                                    borders,
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    children: [new Paragraph({ alignment: AlignmentType.LEFT, children: [new TextRun({ text: "Memanfaatkan deteksi komunitas (Louvain & Affinity Propagation) pada jaringan korelasi kripto untuk diversifikasi; membantu mengurangi volatilitas dan mengoptimalkan return.", font: "Times New Roman", size: 22 })] })]
                                }),
                            ]
                        }),
                        new TableRow({
                            children: [
                                new TableCell({
                                    borders,
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    children: [new Paragraph({ alignment: AlignmentType.LEFT, children: [new TextRun({ text: "3", font: "Times New Roman", size: 22 })] })]
                                }),
                                new TableCell({
                                    borders,
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    children: [new Paragraph({ children: [new TextRun({ text: "Jing & Rocha [mendeley_cite:jing2023network]", font: "Times New Roman", size: 22 })] })]
                                }),
                                new TableCell({
                                    borders,
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    children: [
                                        new Paragraph({ children: [new TextRun({ text: "A network-based strategy of price correlations for optimal cryptocurrency portfolios", font: "Times New Roman", size: 22 })] }),
                                        new Paragraph({ children: [new TextRun({ text: "arXiv: 2304.02362", font: "Times New Roman", size: 20, italics: true })] }),
                                        new Paragraph({ children: [new TextRun({ text: "DOI: 10.1007/s40745-023-00473-7", font: "Times New Roman", size: 20, italics: true })] })
                                    ]
                                }),
                                new TableCell({
                                    borders,
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    children: [new Paragraph({ alignment: AlignmentType.LEFT, children: [new TextRun({ text: "Menggabungkan MST dan MPT untuk memilih kripto berdasarkan dekorelasi jaringan; portofolio MST mengungguli seluruh benchmark (BTC, TOP5, RAND).", font: "Times New Roman", size: 22 })] })]
                                }),
                            ]
                        }),
                        new TableRow({
                            children: [
                                new TableCell({
                                    borders,
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    children: [new Paragraph({ alignment: AlignmentType.LEFT, children: [new TextRun({ text: "4", font: "Times New Roman", size: 22 })] })]
                                }),
                                new TableCell({
                                    borders,
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    children: [new Paragraph({ children: [new TextRun({ text: "Kitanovski, et al. [mendeley_cite:kitanovski2024network]", font: "Times New Roman", size: 22 })] })]
                                }),
                                new TableCell({
                                    borders,
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    children: [
                                        new Paragraph({ children: [new TextRun({ text: "Network-based diversification of stock and cryptocurrency portfolios", font: "Times New Roman", size: 22 })] }),
                                        new Paragraph({ children: [new TextRun({ text: "arXiv: 2408.11739", font: "Times New Roman", size: 20, italics: true })] }),
                                        new Paragraph({ children: [new TextRun({ text: "DOI: 10.1007/s41109-025-00708-9", font: "Times New Roman", size: 20, italics: true })] })
                                    ]
                                }),
                                new TableCell({
                                    borders,
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    children: [new Paragraph({ alignment: AlignmentType.LEFT, children: [new TextRun({ text: "Menunjukkan keunggulan strategi jaringan yang lebih resilient selama periode guncangan pasar (Pandemi & Perang) pada aset saham dan kripto.", font: "Times New Roman", size: 22 })] })]
                                }),
                            ]
                        }),
                        new TableRow({
                            children: [
                                new TableCell({
                                    borders,
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    children: [new Paragraph({ alignment: AlignmentType.LEFT, children: [new TextRun({ text: "5", font: "Times New Roman", size: 22 })] })]
                                }),
                                new TableCell({
                                    borders,
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    children: [new Paragraph({ children: [new TextRun({ text: "Jing, et al. [mendeley_cite:jing2025optimising]", font: "Times New Roman", size: 22 })] })]
                                }),
                                new TableCell({
                                    borders,
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    children: [
                                        new Paragraph({ children: [new TextRun({ text: "Optimising cryptocurrency portfolios through stable clustering of price correlation networks", font: "Times New Roman", size: 22 })] }),
                                        new Paragraph({ children: [new TextRun({ text: "arXiv: 2505.24831", font: "Times New Roman", size: 20, italics: true })] })
                                    ]
                                }),
                                new TableCell({
                                    borders,
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    children: [new Paragraph({ alignment: AlignmentType.LEFT, children: [new TextRun({ text: "Mengintegrasikan stable clustering dengan MPT untuk optimasi portofolio di bawah ketidakpastian; memperkuat validitas integrasi Network-MPT hingga fase pasar terbaru.", font: "Times New Roman", size: 22 })] })]
                                }),
                            ]
                        }),
                    ]
                }),
            ],
            headers: {
                default: new Header({
                    children: [
                        new Paragraph({
                            alignment: AlignmentType.RIGHT,
                            children: [new TextRun({ children: [PageNumber.CURRENT], font: "Times New Roman", size: 24 })]
                        })
                    ]
                })
            },
            footers: {
                first: new Footer({
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
                type: SectionType.NEXT_PAGE, titlePage: true
            },
            // Lanjutkan penomoran decimal dari section sebelumnya
            children: [
                chapterHeading("BAB III", "METODOLOGI PENELITIAN"),
                emptyLine(),
                body("Bab ini menyajikan metodologi yang diaplikasikan dalam penelitian ini, mencakup rancangan tahapan logis penelitian, persiapan data, serta metrik evaluasi performa."),
                emptyLine(),
                heading2("3.1 Tahapan Penelitian"),
                body("Penelitian ini dilaksanakan melalui tahapan-tahapan sebagai berikut:"),
                emptyLine(),
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                        new TextRun({
                            text: "[[%IMAGE_FRAMEWORK]]",
                            font: "Times New Roman",
                            size: 22,
                        })
                    ]
                }),
                emptyLine(),
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
                    {text: "Melakukan kajian komprehensif terhadap berbagai literatur ilmiah, jurnal internasional, dan buku teks terkait "},
                    {text: "Modern Portfolio Theory", italic: true},
                    {text: ", "},
                    {text: "Random Matrix Theory", italic: true},
                    {text: ", "},
                    {text: "Graph Theory", italic: true},
                    {text: ", serta model "},
                    {text: "Network Markowitz", italic: true},
                    {text: ". Tahapan ini bertujuan untuk mengidentifikasi "},
                    {text: "research gap", italic: true},
                    {text: ", menentukan parameter dasar optimasi, serta memahami landasan matematis dari pendekatan "},
                    {text: "adaptive grid search", italic: true},
                    {text: " dalam konteks volatilitas pasar kripto."}
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
                    {text: " [mendeley_cite:arratia2014computational]."}
                ]),
                emptyLine(),
                heading2("3.2 Alat dan Bahan Penelitian"),
                heading3("3.2.1 Perangkat Lunak"),
                bulletItem("Sistem Operasi: Windows 10/11"),
                bulletItem("Bahasa Pemrograman: Python 3.x (dengan ekosistem Anaconda)"),
                bulletItem("Framework/Library: Pandas, Numpy, Scipy (Optimization), Scikit-Learn, NetworkX (Graph Analytics)"),
                emptyLine(),
                heading2("3.3 Dataset"),
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
                    width: { size: 8200, type: WidthType.DXA },
                    columnWidths: [800, 2000, 3400, 2000],
                    rows: [
                        new TableRow({
                            tableHeader: true,
                            children: [
                                new TableCell({ borders, children: [new Paragraph({ children: [new TextRun({ text: "Ticker", font: "Times New Roman", size: 22, bold: true })] })] }),
                                new TableCell({ borders, children: [new Paragraph({ children: [new TextRun({ text: "Nama Aset", font: "Times New Roman", size: 22, bold: true })] })] }),
                                new TableCell({ borders, children: [new Paragraph({ children: [new TextRun({ text: "Kategori / Use Case", font: "Times New Roman", size: 22, bold: true })] })] }),
                                new TableCell({ borders, children: [new Paragraph({ children: [new TextRun({ text: "Sumber", font: "Times New Roman", size: 22, bold: true })] })] }),
                            ]
                        }),
                        ...[
                            ["BTC", "Bitcoin", "Layer 1 / Store of Value", "Yahoo Finance"],
                            ["ETH", "Ethereum", "Layer 1 / Smart Contract", "Yahoo Finance"],
                            ["XRP", "Ripple", "Payment / Bridge Currency", "Yahoo Finance"],
                            ["USDT", "Tether", "Stablecoin / USD Pegged", "Yahoo Finance"],
                            ["BCH", "Bitcoin Cash", "Payment / Peer-to-Peer Cash", "Yahoo Finance"],
                            ["LTC", "Litecoin", "Payment / Digital Silver", "Yahoo Finance"],
                            ["BNB", "Binance Coin", "Layer 1 / Exchange Token", "Yahoo Finance"],
                            ["EOS", "EOS", "Layer 1 / Smart Contract", "Yahoo Finance"],
                            ["XLM", "Stellar", "Payment / Bridge Currency", "Yahoo Finance"],
                            ["TRX", "Tron", "Layer 1 / Smart Contract", "Yahoo Finance"],
                        ].map(([t, n, k, s]) => new TableRow({
                            children: [
                                new TableCell({ borders, children: [new Paragraph({ children: [new TextRun({ text: t, font: "Times New Roman", size: 22 })] })] }),
                                new TableCell({ borders, children: [new Paragraph({ children: [new TextRun({ text: n, font: "Times New Roman", size: 22 })] })] }),
                                new TableCell({ borders, children: [new Paragraph({ children: [new TextRun({ text: k, font: "Times New Roman", size: 22 })] })] }),
                                new TableCell({ borders, children: [new Paragraph({ children: [new TextRun({ text: s, font: "Times New Roman", size: 22 })] })] }),
                            ]
                        }))
                    ]
                }),
                emptyLine(),
                mixedBody([
                    {text: "Pemilihan aset ini selaras dengan referensi Giudici et al. [mendeley_cite:giudici2020network] yang menjadi landasan penelitian ini."}
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
                heading2("3.4 Metode/Algoritma yang Digunakan"),
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
                        new Math({
                            children: rumus("\\min_w w^T \\Sigma^* w + \\gamma \\sum_{i=1}^n x_i w_i")
                        }),
                        new TextRun({ text: "             (3.1)", font: "Times New Roman", size: 24, bold: true }),
                    ],
                }),
                body("Keterangan:"),
                new Table({
                    width: { size: 8200, type: WidthType.DXA },
                    columnWidths: [800, 400, 7000],
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
                                new TableCell({ children: [new Paragraph({ children: [new Math({ children: [new MathRun("w")] })] })] }),
                                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "=", font: "Times New Roman", size: 24 })] })] }),
                                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Vektor alokasi bobot untuk setiap aset kripto (total = 1).", font: "Times New Roman", size: 24 })] })] }),
                            ]
                        }),
                        new TableRow({
                            children: [
                                new TableCell({ children: [new Paragraph({ children: [new Math({ children: [new MathSuperScript({ children: [new MathRun("\u03A3")], superScript: [new MathRun("*")] })] })] })] }),
                                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "=", font: "Times New Roman", size: 24 })] })] }),
                                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Matriks Kovarians terfilter RMT.", font: "Times New Roman", size: 24 })] })] }),
                            ]
                        }),
                        new TableRow({
                            children: [
                                new TableCell({ children: [new Paragraph({ children: [new Math({ children: [new MathRun("\u03B3")] })] })] }),
                                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "=", font: "Times New Roman", size: 24 })] })] }),
                                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Parameter skalar penghukuman sentralitas graf.", font: "Times New Roman", size: 24 })] })] }),
                            ]
                        }),
                        new TableRow({
                            children: [
                                new TableCell({ children: [new Paragraph({ children: [new Math({ children: [new MathSubScript({ children: [new MathRun("x")], subScript: [new MathRun("i")] })] })] })] }),
                                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "=", font: "Times New Roman", size: 24 })] })] }),
                                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Vektor skor Eigenvector Centrality tiap node aset.", font: "Times New Roman", size: 24 })] })] }),
                            ]
                        }),
                    ],
                }),
                emptyLine(),
                heading3("3.4.1 Strategi Portofolio dan Benchmark"),
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
                heading2("3.5 Matriks Evaluasi Performa"),
                body("Untuk mengukur efektivitas model portofolio yang dikembangkan, digunakan beberapa metrik evaluasi sebagai berikut:"),
                
                heading3("3.5.1 Risk-Adjusted Return (Sharpe Ratio)"),
                mixedBody([
                    {text: "Sharpe Ratio", italic: true},
                    {text: " merupakan metrik standar industri yang diperkenalkan oleh William F. Sharpe (1966) untuk mengukur imbal hasil berlebih ("},
                    {text: "excess return", italic: true},
                    {text: ") per unit risiko total. Formula yang digunakan adalah:"}
                ]),
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    spacing: { before: 120, after: 120 },
                    children: [
                        new Math({
                            children: rumus("Sharpe\\ Ratio = \\frac{R_p - R_f}{\\sigma_p}")
                        }),
                        new TextRun({ text: "             (3.2)", font: "Times New Roman", size: 24, bold: true }),
                    ],
                }),
                body("Keterangan:"),
                new Table({
                    width: { size: 8200, type: WidthType.DXA },
                    columnWidths: [800, 400, 7000],
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
                                new TableCell({ children: [new Paragraph({ children: [new Math({ children: [new MathSubScript({ children: [new MathRun("R")], subScript: [new MathRun("p")] })] })] })] }),
                                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "=", font: "Times New Roman", size: 24 })] })] }),
                                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Imbal hasil (return) portofolio.", font: "Times New Roman", size: 24 })] })] }),
                            ]
                        }),
                        new TableRow({
                            children: [
                                new TableCell({ children: [new Paragraph({ children: [new Math({ children: [new MathSubScript({ children: [new MathRun("R")], subScript: [new MathRun("f")] })] })] })] }),
                                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "=", font: "Times New Roman", size: 24 })] })] }),
                                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Tingkat imbal hasil bebas risiko (risk-free rate).", font: "Times New Roman", size: 24 })] })] }),
                            ]
                        }),
                        new TableRow({
                            children: [
                                new TableCell({ children: [new Paragraph({ children: [new Math({ children: [new MathSubScript({ children: [new MathRun("\u03C3")], subScript: [new MathRun("p")] })] })] })] }),
                                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "=", font: "Times New Roman", size: 24 })] })] }),
                                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Standar deviasi dari imbal hasil berlebih portofolio.", font: "Times New Roman", size: 24 })] })] }),
                            ]
                        }),
                    ],
                }),
                emptyLine(),
                mixedBody([
                    {text: "Nilai Sharpe Ratio > 1 dianggap baik, > 2 sangat baik, dan > 3 luar biasa. Semakin besar nilai "},
                    {text: "Sharpe Ratio", italic: true},
                    {text: ", semakin baik kualitas portofolio karena menunjukkan imbal hasil yang lebih tinggi untuk setiap unit risiko yang diambil. Sebaliknya, semakin kecil nilai ini, semakin tidak efisien portofolio tersebut dalam menghasilkan imbal hasil terhadap risikonya [mendeley_cite:markowitz1952portfolio], [mendeley_cite:lopezdeprado2018advances]."}
                ]),
                emptyLine(),
                
                heading3("3.5.2 Downside Risk (Value at Risk - VaR)"),
                mixedBody([
                    {text: "Value at Risk", italic: true},
                    {text: " (VaR) mengkuantifikasi potensi kerugian maksimal yang mungkin dialami portofolio dalam satu periode perdagangan pada tingkat kepercayaan tertentu. Pada penelitian ini digunakan VaR "},
                    {text: "historical simulation", italic: true},
                    {text: " pada tingkat kepercayaan 95%, yang berarti terdapat probabilitas 5% bahwa kerugian aktual akan melebihi nilai VaR yang dihitung. Secara formal:"}
                ]),
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    spacing: { before: 120, after: 120 },
                    children: [
                        new Math({
                            children: rumus("VaR_\\alpha = -inf { x \u2208 \u211D : P(L > x) \u2264 1 - \u03B1 }")
                        }),
                        new TextRun({ text: "             (3.3)", font: "Times New Roman", size: 24, bold: true }),
                    ],
                }),
                body("Keterangan:"),
                new Table({
                    width: { size: 8200, type: WidthType.DXA },
                    columnWidths: [800, 400, 7000],
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
                                new TableCell({ children: [new Paragraph({ children: [new Math({ children: [new MathSubScript({ children: [new MathRun("VaR")], subScript: [new MathRun("\u03B1")] })] })] })] }),
                                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "=", font: "Times New Roman", size: 24 })] })] }),
                                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Value at Risk pada tingkat kepercayaan \u03B1.", font: "Times New Roman", size: 24 })] })] }),
                            ]
                        }),
                        new TableRow({
                            children: [
                                new TableCell({ children: [new Paragraph({ children: [new Math({ children: [new MathRun("\u03B1")] })] })] }),
                                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "=", font: "Times New Roman", size: 24 })] })] }),
                                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Tingkat kepercayaan (confidence level), misal 95%.", font: "Times New Roman", size: 24 })] })] }),
                            ]
                        }),
                        new TableRow({
                            children: [
                                new TableCell({ children: [new Paragraph({ children: [new Math({ children: [new MathRun("L")] })] })] }),
                                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "=", font: "Times New Roman", size: 24 })] })] }),
                                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Variabel acak kerugian portofolio (Loss).", font: "Times New Roman", size: 24 })] })] }),
                            ]
                        }),
                    ],
                }),
                mixedBody([
                    {text: "VaR dipilih karena relevansinya yang tinggi terhadap pasar kripto yang memiliki volatilitas ekstrem. Semakin kecil nilai "},
                    {text: "Value at Risk", italic: true},
                    {text: " (mendekati nol), semakin aman suatu portofolio dari potensi kerugian ekstrem. Sebaliknya, semakin besar nilai VaR, semakin tinggi risiko kerugian yang mungkin dihadapi investor dalam kondisi pasar yang buruk [mendeley_cite:corbet2019cryptocurrencies], [mendeley_cite:jorion2000value]."}
                ]),
                emptyLine(),
                
                heading3("3.5.3 Tail Risk & Reward (Rachev Ratio)"),
                mixedBody([
                    {text: "Rachev Ratio", italic: true},
                    {text: " merupakan ukuran performa yang secara eksplisit memperhitungkan distribusi ekor ("},
                    {text: "tail distribution", italic: true},
                    {text: ") dari "},
                    {text: "return", italic: true},
                    {text: " portofolio. Berbeda dengan Sharpe Ratio yang mengasumsikan distribusi normal, Rachev Ratio didefinisikan sebagai rasio antara "},
                    {text: "Expected Tail Return", italic: true},
                    {text: " (ETR) pada kuantil atas terhadap "},
                    {text: "Expected Shortfall", italic: true},
                    {text: " (ES) pada kuantil bawah:"}
                ]),
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    spacing: { before: 120, after: 120, line: 360 },
                    children: [
                        new Math({
                            children: rumus("RR = \\frac{ETR_\\alpha(R)}{ES_\\beta(R)}")
                        }),
                        new TextRun({ text: "             (3.4)", font: "Times New Roman", size: 24, bold: true }),
                    ]
                }),
                body("Keterangan:"),
                new Paragraph({
                    indent: { left: 720 },
                    spacing: { line: 360 },
                    tabStops: [{ type: TabStopType.LEFT, position: 2000 }],
                    children: [
                        new Math({ children: [new MathRun("RR")] }),
                        new TextRun({ text: "\t= Rachev Ratio.", font: "Times New Roman", size: 24 }),
                    ]
                }),
                new Paragraph({
                    indent: { left: 720 },
                    spacing: { line: 360 },
                    tabStops: [{ type: TabStopType.LEFT, position: 2000 }],
                    children: [
                        new Math({ children: [new MathSubScript({ children: [new MathRun("ETR")], subScript: [new MathRun("\u03B1")] })] }),
                        new TextRun({ text: "\t= Expected Tail Return pada tingkat kepercayaan \u03B1.", font: "Times New Roman", size: 24 }),
                    ]
                }),
                new Paragraph({
                    indent: { left: 720 },
                    spacing: { line: 360 },
                    tabStops: [{ type: TabStopType.LEFT, position: 2000 }],
                    children: [
                        new Math({ children: [new MathSubScript({ children: [new MathRun("ES")], subScript: [new MathRun("\u03B2")] })] }),
                        new TextRun({ text: "\t= Expected Shortfall pada tingkat kepercayaan \u03B2.", font: "Times New Roman", size: 24 }),
                    ]
                }),
                mixedBody([
                    {text: "Nilai Rachev Ratio > 1 mengindikasikan bahwa potensi keuntungan ekstrem melebihi potensi kerugian ekstrem. Semakin besar nilai "},
                    {text: "Rachev Ratio", italic: true},
                    {text: ", semakin baik profil risiko-imbalan suatu portofolio karena menunjukkan kemampuan portofolio untuk menangkap keuntungan di 'ekor kanan' distribusi melampaui risiko di 'ekor kiri'. Metrik ini sangat krusial untuk pasar kripto yang dikenal memiliki karakteristik "},
                    {text: "leptokurtic", italic: true},
                    {text: " [mendeley_cite:rachev2008advanced], [mendeley_cite:artzner1999coherent]."}
                ]),

                emptyLine(),
                heading2("3.6 Rencana Jadwal Penelitian"),
                body("Penelitian ini direncanakan akan dilaksanakan selama empat bulan. Rincian jadwal pelaksanaan setiap tahapan kegiatan disajikan pada Tabel III.2 berikut:"),
                emptyLine(),
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [new TextRun({ text: "Tabel III.2. Rencana Jadwal Penelitian", font: "Times New Roman", size: 24, bold: true })]
                }),
                new Table({
                    width: { size: 8200, type: WidthType.DXA },
                    columnWidths: [3400, 1200, 1200, 1200, 1200],
                    rows: [
                        new TableRow({
                            tableHeader: true,
                            children: [
                                new TableCell({
                                    borders,
                                    shading: { fill: "D5E8F0", type: ShadingType.CLEAR },
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Kegiatan", font: "Times New Roman", size: 22, bold: true })] })]
                                }),
                                ...["Bln 1", "Bln 2", "Bln 3", "Bln 4"].map(h => new TableCell({
                                    borders,
                                    shading: { fill: "D5E8F0", type: ShadingType.CLEAR },
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: h, font: "Times New Roman", size: 22, bold: true })] })]
                                }))
                            ]
                        }),
                        ...[
                            ["Studi Literatur", ["✓", "✓", "", ""]],
                            ["Pengumpulan Data", ["✓", "", "", ""]],
                            ["Pra-pemrosesan Data", ["✓", "✓", "", ""]],
                            ["Perancangan Model/Sistem", ["", "✓", "✓", ""]],
                            ["Implementasi", ["", "", "✓", "✓"]],
                            ["Pengujian dan Evaluasi", ["", "", "✓", "✓"]],
                            ["Penulisan Laporan", ["✓", "✓", "✓", "✓"]],
                        ].map(([activity, marks]) => new TableRow({
                            children: [
                                new TableCell({
                                    borders,
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    children: [new Paragraph({ children: [new TextRun({ text: activity, font: "Times New Roman", size: 22 })] })]
                                }),
                                ...marks.map(m => new TableCell({
                                    borders,
                                    margins: { top: 80, bottom: 80, left: 120, right: 120 },
                                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: m, font: "Times New Roman", size: 22 })] })]
                                }))
                            ]
                        }))
                    ]
                }),
            ],
            headers: {
                default: new Header({
                    children: [
                        new Paragraph({
                            alignment: AlignmentType.RIGHT,
                            children: [new TextRun({ children: [PageNumber.CURRENT], font: "Times New Roman", size: 24 })]
                        })
                    ]
                })
            },
            footers: {
                first: new Footer({
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
                type: SectionType.NEXT_PAGE, titlePage: true
            },
            // Lanjutkan penomoran decimal dari section sebelumnya
            children: [
                sectionTitle("DAFTAR REFERENSI"),
                emptyLine(),
                new Paragraph({
                    alignment: AlignmentType.JUSTIFIED,
                    spacing: { before: 0, after: 120, line: 360 },
                    children: [new TextRun({ text: "[mendeley_bibliography]", font: "Times New Roman", size: 24 })]
                }),
            ],
            headers: {
                default: new Header({
                    children: [
                        new Paragraph({
                            alignment: AlignmentType.RIGHT,
                            children: [new TextRun({ children: [PageNumber.CURRENT], font: "Times New Roman", size: 24 })]
                        })
                    ]
                })
            },
            footers: {
                first: new Footer({
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
