const docx = require('docx');
const {
    Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
    AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
    LevelFormat, PageNumber, PageBreak, TabStopType, TabStopPosition,
    VerticalAlign, ImageRun, Header, Footer, NumberFormat, SectionType,
    TableOfContents,
    Math, MathRun, MathSubScript, MathSuperScript, MathFraction, MathSum, MathRoundBrackets, MathLimitLower, MathRadical
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

    // Penanganan untuk Marchenko-Pastur
    if (latexStr.includes("\\lambda_{\\pm}")) {
        return [
            new MathSubScript({ children: [new MathRun("\u03BB")], subScript: [new MathRun("\u00B1")] }),
            new MathRun(" = \u03C3"),
            new MathSuperScript({ children: [new MathRun(" ")], superScript: [new MathRun("2")] }),
            new MathRoundBrackets({
                children: [
                    new MathRun("1 \u00B1 "),
                    new MathRadical({
                        children: [
                            new MathFraction({
                                numerator: [new MathRun("1")],
                                denominator: [new MathRun("Q")]
                            })
                        ]
                    })
                ]
            }),
            new MathSuperScript({ children: [new MathRun(" ")], superScript: [new MathRun("2")] })
        ];
    }

    // Penanganan untuk CVaR
    if (latexStr.includes("CVaR")) {
        return [
            new MathSubScript({ children: [new MathRun("CVaR")], subScript: [new MathRun("\u03B1")] }),
            new MathRun(" = E "),
            new MathRoundBrackets({
                children: [
                    new MathRun("L | L \u2265 "),
                    new MathSubScript({ children: [new MathRun("VaR")], subScript: [new MathRun("\u03B1")] })
                ]
            })
        ];
    }

    // Penanganan untuk Sortino
    if (latexStr.includes("Sortino")) {
        return [
            new MathRun("Sortino Ratio = "),
            new MathFraction({
                numerator: [
                    new MathSubScript({ children: [new MathRun("R")], subScript: [new MathRun("p")] }),
                    new MathRun(" - "),
                    new MathSubScript({ children: [new MathRun("R")], subScript: [new MathRun("f")] })
                ],
                denominator: [
                    new MathSubScript({ children: [new MathRun("\u03C3")], subScript: [new MathRun("d")] })
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

function heading4(text) {
    return new Paragraph({
        children: [new TextRun({ text, font: "Times New Roman", size: 24, bold: true })]
    });
}

const guidlineFooter = new Paragraph({
    alignment: AlignmentType.RIGHT,
    children: [
        new TextRun({
            text: "Program Studi Informatika (S2) FTI Universitas Nusa Mandiri",
            font: "Times New Roman",
            size: 20, // size 10 in Word
            bold: true,
        })
    ]
});

function formulaParagraph(mathChildren, label) {
    return new Paragraph({
        alignment: AlignmentType.LEFT,
        indent: { left: 850 }, // 1.5 cm dari kiri
        tabStops: [
            {
                type: TabStopType.RIGHT,
                position: 8221, // Batas kanan pengetikan (8221 DXA)
            },
        ],
        spacing: { before: 240, after: 240 },
        children: [
            new Math({ children: mathChildren }),
            new TextRun({ text: "\t(" + label + ")", font: "Times New Roman", size: 24, bold: true }),
        ],
    });
}

function formulaKeterangan(rowsData) {
    return new Table({
        width: { size: 8221, type: WidthType.DXA },
        columnWidths: [1000, 400, 6821],
        borders: {
            top: { style: BorderStyle.NONE, size: 0 },
            bottom: { style: BorderStyle.NONE, size: 0 },
            left: { style: BorderStyle.NONE, size: 0 },
            right: { style: BorderStyle.NONE, size: 0 },
            insideHorizontal: { style: BorderStyle.NONE, size: 0 },
            insideVertical: { style: BorderStyle.NONE, size: 0 },
        },
        rows: rowsData.map(([symbol, desc]) => new TableRow({
            children: [
                new TableCell({ children: [new Paragraph({ children: [Array.isArray(symbol) ? new Math({ children: symbol }) : new TextRun({ text: symbol, font: "Times New Roman", size: 24, italic: true })] })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "=", font: "Times New Roman", size: 24 })] })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: desc, font: "Times New Roman", size: 24 })] })] }),
            ]
        })),
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
            },
            {
                id: "TOC1", name: "toc 1", basedOn: "Normal", next: "Normal", quickFormat: true,
                run: { size: 24, bold: true, font: "Times New Roman" },
            }
        ]
    },
    sections: [
        // ==================== COVER PAGE ====================
        {
            properties: {
                page: {
                    size: { width: 11906, height: 16838 },
                    margin: { top: 1701, right: 1417, bottom: 1417, left: 2268, footer: 1200 },
                    pageNumbers: { start: 1, formatType: NumberFormat.LOWER_ROMAN }
                },
                pageNumbers: { start: 1, formatType: NumberFormat.LOWER_ROMAN }
            },
            footers: {
                default: new Footer({
                    children: []
                })
            },
            children: [
                new Paragraph({
                    heading: HeadingLevel.HEADING_1,
                    alignment: AlignmentType.CENTER,
                    children: [new TextRun({ text: "HALAMAN JUDUL", color: "FFFFFF", size: 2 })]
                }),
                emptyLine(),
                centeredBold("OPTIMALISASI DINAMIS PORTOFOLIO", 28),
                centeredBold("NETWORK MARKOWITZ", 28),
                centeredBold("BERBASIS DEEP REINFORCEMENT LEARNING", 28),
                centeredBold("YANG TERINTERPRETASI (EXPLAINABLE AI)", 28),
                emptyLine(),
                emptyLine(),
                emptyLine(),
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                        new TextRun({
                            text: "[[%LOGO_UNM]]",
                            font: "Times New Roman",
                            size: 22,
                        })
                    ]
                }),
                emptyLine(),
                centeredBold("PROPOSAL TESIS", 36),
                emptyLine(),
                centered("Diajukan sebagai salah satu syarat untuk memperoleh gelar"),
                centered("Magister Komputer (M.Kom)"),
                emptyLine(),
                emptyLine(),
                centeredBold("Ragil Yulianto", 28),
                centered("14240007"),
                emptyLine(),
                emptyLine(),
                emptyLine(),
                emptyLine(),
                emptyLine(),
                centered("Program Studi Informatika (S2)"),
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
                    margin: { top: 1701, right: 1417, bottom: 1417, left: 2268, footer: 1417 },
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
                            children: [new TextRun({ children: [PageNumber.CURRENT], font: "Times New Roman", size: 24 })]
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
                    ["Program Studi", "Informatika"],
                    ["Fakultas",      "Teknologi Informasi"],
                    ["Jenjang",       "Strata Dua (S2)"],
                    ["Judul Tesis",   "OPTIMALISASI DINAMIS PORTOFOLIO NETWORK MARKOWITZ BERBASIS DEEP REINFORCEMENT LEARNING YANG TERINTERPRETASI (EXPLAINABLE AI)"],
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
                    children: [new TextRun({ text: "Jakarta, 15 April 2026", font: "Times New Roman", size: 24 })]
                }),
                emptyLine(),
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [new TextRun({ text: "PEMBIMBING TESIS", font: "Times New Roman", size: 24, bold: true })]
                }),
                emptyLine(),
                emptyLine(),
                emptyLine(),
                new Paragraph({
                    indent: { left: 0 },
                    children: [
                        new TextRun({ text: "Pembimbing", font: "Times New Roman", size: 24 }),
                        new TextRun({ text: "\t:  Dr. Muhammad Haris, M. Eng", font: "Times New Roman", size: 24 }),
                        new TextRun({ text: "\t.............................", font: "Times New Roman", size: 24 }),
                    ],
                    tabStops: [
                        { type: TabStopType.LEFT, position: 2000 },
                        { type: TabStopType.RIGHT, position: 8000 }
                    ]
                }),
            ]
        },
        // ==================== PERNYATAAN ORISINALITAS ====================
        {
            properties: {
                page: {
                    size: { width: 11906, height: 16838 },
                    margin: { top: 1701, right: 1417, bottom: 1417, left: 2268, footer: 1417 },
                },
                type: SectionType.NEXT_PAGE,
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
                sectionTitle("SURAT PERNYATAAN ORISINALITAS DAN BEBAS PLAGIARISME"),
                emptyLine(),
                bodyNoIndent("Yang bertanda tangan di bawah ini:"),
                emptyLine(),
                ...[
                    ["Nama",          "Ragil Yulianto"],
                    ["NIM",           "14240007"],
                    ["Program Studi", "Informatika"],
                    ["Fakultas",      "Teknologi Informasi"],
                    ["Jenjang",       "Strata Dua (S2)"],
                    ["Peminatan",     "Data Science / Artificial Intelligence"],
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
                body("Dengan ini menyatakan bahwa tesis yang telah saya buat dengan judul:"),
                centeredBold("OPTIMALISASI DINAMIS PORTOFOLIO NETWORK MARKOWITZ BERBASIS DEEP REINFORCEMENT LEARNING YANG TERINTERPRETASI (EXPLAINABLE AI)", 24),
                body("adalah hasil karya sendiri, dan semua sumber baik yang dikutip maupun yang dirujuk telah saya nyatakan dengan benar, serta belum pernah diterbitkan atau dipublikasikan dimanapun dan dalam bentuk apapun."),
                body("Demikianlah surat pernyataan ini saya buat dengan sebenar-benarnya. Apabila dikemudian hari ternyata saya memberikan keterangan palsu dan atau ada pihak lain yang mengklaim bahwa tesis yang telah saya buat adalah hasil karya milik seseorang atau badan tertentu, saya bersedia diproses baik secara pidana maupun perdata dan kelulusan saya dari Program Studi Informatika (S2) Fakultas Teknologi Informasi Universitas Nusa Mandiri dicabut/dibatalkan."),
                emptyLine(),
                new Paragraph({
                    alignment: AlignmentType.RIGHT,
                    children: [new TextRun({ text: "Jakarta, 15 April 2026", font: "Times New Roman", size: 24 })]
                }),
                new Paragraph({
                    alignment: AlignmentType.RIGHT,
                    children: [new TextRun({ text: "Yang menyatakan,", font: "Times New Roman", size: 24 })]
                }),
                emptyLine(),
                emptyLine(),
                emptyLine(),
                new Paragraph({
                    alignment: AlignmentType.RIGHT,
                    children: [
                        new TextRun({ text: "Materai Rp 10.000", font: "Times New Roman", size: 18 }),
                    ]
                }),
                new Paragraph({
                    alignment: AlignmentType.RIGHT,
                    children: [
                        new TextRun({ text: "Ragil Yulianto", font: "Times New Roman", size: 24, bold: true }),
                    ]
                }),
            ]
        },
        // ==================== PERSETUJUAN PUBLIKASI ====================
        {
            properties: {
                page: {
                    size: { width: 11906, height: 16838 },
                    margin: { top: 1701, right: 1417, bottom: 1417, left: 2268, footer: 1417 },
                },
                type: SectionType.NEXT_PAGE,
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
                sectionTitle("SURAT PERNYATAAN PERSETUJUAN PUBLIKASI KARYA ILMIAH UNTUK KEPENTINGAN AKADEMIS"),
                emptyLine(),
                bodyNoIndent("Yang bertanda tangan di bawah ini:"),
                emptyLine(),
                ...[
                    ["Nama",          "Ragil Yulianto"],
                    ["NIM",           "14240007"],
                    ["Program Studi", "Informatika"],
                    ["Fakultas",      "Teknologi Informasi"],
                    ["Jenjang",       "Strata Dua (S2)"],
                    ["Jenis Karya",   "Tesis"],
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
                body("Demi pengembangan ilmu pengetahuan, dengan ini menyetujui untuk memberikan izin kepada pihak Program Studi Informatika (S2) Fakultas Teknologi Informasi Universitas Nusa Mandiri, Hak Bebas Royalti Non-Eksklusif (Non-exclusive Royalti-Free Right) atas karya ilmiah saya yang berjudul:"),
                centeredBold("OPTIMALISASI DINAMIS PORTOFOLIO NETWORK MARKOWITZ BERBASIS DEEP REINFORCEMENT LEARNING YANG TERINTERPRETASI (EXPLAINABLE AI)", 24),
                body("Dengan Hak Bebas Royalti Non-Eksklusif ini, pihak Universitas Nusa Mandiri berhak menyimpan, mengalih-media atau bentuk-kan, mengelolanya dalam pangkalan data (database), mendistribusikannya dan menampilkan atau mempublikasikannya di internet atau media lain untuk kepentingan akademis tanpa perlu meminta izin dari saya selama tetap mencantumkan nama saya sebagai penulis/pencipta karya ilmiah tersebut."),
                body("Saya bersedia untuk menanggung secara pribadi, tanpa melibatkan pihak Universitas Nusa Mandiri, segala bentuk tuntutan hukum yang timbul atas pelanggaran Hak Cipta dalam karya ilmiah saya ini."),
                emptyLine(),
                new Paragraph({
                    alignment: AlignmentType.RIGHT,
                    children: [new TextRun({ text: "Jakarta, 15 April 2026", font: "Times New Roman", size: 24 })]
                }),
                new Paragraph({
                    alignment: AlignmentType.RIGHT,
                    children: [new TextRun({ text: "Yang menyatakan,", font: "Times New Roman", size: 24 })]
                }),
                emptyLine(),
                emptyLine(),
                emptyLine(),
                new Paragraph({
                    alignment: AlignmentType.RIGHT,
                    children: [
                        new TextRun({ text: "Materai Rp 10.000", font: "Times New Roman", size: 18 }),
                    ]
                }),
                new Paragraph({
                    alignment: AlignmentType.RIGHT,
                    children: [
                        new TextRun({ text: "Ragil Yulianto", font: "Times New Roman", size: 24, bold: true }),
                    ]
                }),
            ]
        },
        // ==================== KATA PENGANTAR ====================
        {
            properties: {
                page: {
                    size: { width: 11906, height: 16838 },
                    margin: { top: 1701, right: 1417, bottom: 1417, left: 2268, footer: 1417 },
                },
                type: SectionType.NEXT_PAGE,
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
                sectionTitle("KATA PENGANTAR"),
                emptyLine(),
                body("Puji syukur penulis panjatkan ke hadirat Allah SWT atas segala rahmat dan hidayah-Nya, sehingga penulis dapat menyelesaikan proposal tesis ini dengan judul \"OPTIMALISASI DINAMIS PORTOFOLIO NETWORK MARKOWITZ BERBASIS DEEP REINFORCEMENT LEARNING YANG TERINTERPRETASI (EXPLAINABLE AI)\"."),
                body("Penulis menyadari bahwa keberhasilan penyusunan proposal ini tidak lepas dari bantuan, bimbingan, dan dukungan dari berbagai pihak. Oleh karena itu, penulis menyampaikan ucapan terima kasih kepada Bapak Dr. Muhammad Haris, M. Eng selaku dosen pembimbing yang telah memberikan arahan berharga dalam penelitian ini."),
                body("Penulis menyadari masih banyak kekurangan dalam proposal ini. Oleh karena itu, penulis mengharapkan kritik dan saran yang membangun demi penyempurnaan di masa yang akan datang."),
                emptyLine(),
                new Paragraph({
                    alignment: AlignmentType.RIGHT,
                    children: [new TextRun({ text: "Jakarta, 15 April 2026", font: "Times New Roman", size: 24 })]
                }),
                emptyLine(),
                new Paragraph({
                    alignment: AlignmentType.RIGHT,
                    children: [new TextRun({ text: "Penulis", font: "Times New Roman", size: 24, bold: true })]
                }),
            ]
        },
        // ==================== ABSTRAK ====================
        {
            properties: {
                page: {
                    size: { width: 11906, height: 16838 },
                    margin: { top: 1701, right: 1417, bottom: 1417, left: 2268, footer: 1417 },
                },
                type: SectionType.NEXT_PAGE,
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
                sectionTitle("ABSTRAK"),
                emptyLine(),
                bodyNoIndent("Tesis ini mengusulkan sebuah metodologi dinamis untuk optimasi portofolio mata uang kripto dengan mengintegrasikan Teori Jaringan Kompleks dan Deep Reinforcement Learning (DRL). Model yang diusulkan, disebut SAC-Net Markowitz, menggunakan agen Soft Actor-Critic (SAC) untuk secara adaptif menyesuaikan intensitas penalti sentralitas (gamma) berdasarkan kondisi jaringan pasar yang berubah. Selain itu, teknik Explainable AI (XAI) melalui metode SHAP diintegrasikan untuk memberikan interpretasi yang transparan terhadap logika pengambilan keputusan agen. Hasil eksperimen menunjukkan bahwa pendekatan ini mampu memberikan profil risiko-imbal hasil yang lebih unggul dibandingkan strategi benchmark konvensional."),
                emptyLine(),
                bodyNoIndent("Kata Kunci: Deep Reinforcement Learning, Soft Actor-Critic, Markowitz, Complex Network, Explainable AI, Cryptocurrency.", { bold: true }),
            ]
        },
        // ==================== ABSTRACT ====================
        {
            properties: {
                page: {
                    size: { width: 11906, height: 16838 },
                    margin: { top: 1701, right: 1417, bottom: 1417, left: 2268, footer: 1417 },
                },
                type: SectionType.NEXT_PAGE,
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
                sectionTitle("ABSTRACT"),
                emptyLine(),
                bodyNoIndent("This thesis proposes a dynamic methodology for cryptocurrency portfolio optimization by integrating Complex Network Theory and Deep Reinforcement Learning (DRL). The proposed model, named SAC-Net Markowitz, utilizes a Soft Actor-Critic (SAC) agent to adaptively adjust the centrality penalty intensity (gamma) in response to changing market network conditions. Furthermore, Explainable AI (XAI) techniques using the SHAP method are integrated to provide transparent interpretation of the agent's decision-making logic. Experimental results demonstrate that this approach delivers a superior risk-adjusted return profile compared to conventional benchmark strategies.", { italics: true }),
                emptyLine(),
                bodyNoIndent("Keywords: Deep Reinforcement Learning, Soft Actor-Critic, Markowitz, Complex Network, Explainable AI, Cryptocurrency.", { bold: true, italics: true }),
            ]
        },
        // ==================== DAFTAR ISI ====================
        {
            properties: {
                page: {
                    size: { width: 11906, height: 16838 },
                    margin: { top: 1701, right: 1417, bottom: 1417, left: 2268, footer: 1417 },
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
                            children: [new TextRun({ children: [PageNumber.CURRENT], font: "Times New Roman", size: 24 })]
                        })
                    ]
                })
            },
            children: [
                sectionTitle("DAFTAR ISI", 26),
                new TableOfContents("Daftar Isi", {
                    hyperlink: true,
                    headingStyleRange: "1-3",
                    caption: { text: "Daftar Isi" },
                }),
            ]
        },
        // ==================== BAB I ====================
        {
            properties: {
                page: {
                    size: { width: 11906, height: 16838 },
                    margin: { top: 1701, right: 1417, bottom: 1417, left: 2268, header: 850, footer: 1417 },
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
                }),
                first: new Header({ children: [new Paragraph({})] })
            },
            footers: {
                first: new Footer({
                    children: [
                        new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [new TextRun({ children: [PageNumber.CURRENT], font: "Times New Roman", size: 24 })]
                        }),
                        guidlineFooter
                    ]
                }),
                default: new Footer({
                    children: [guidlineFooter]
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
                        new TextRun({ text: "Industri keuangan global dalam beberapa tahun terakhir menyaksikan adanya lonjakan penggunaan konsultasi keuangan otomatis (", font: "Times New Roman", size: 24 }),
                        new TextRun({ text: "robo-advisors", font: "Times New Roman", size: 24, italics: true }),
                        new TextRun({ text: ") yang dirancang untuk membantu investor mengelola portofolio secara sistematis dan efisien [mendeley_cite:giudici2020network]. Fenomena ini beriringan dengan perkembangan mata uang kripto (", font: "Times New Roman", size: 24 }),
                        new TextRun({ text: "cryptocurrency", font: "Times New Roman", size: 24, italics: true }),
                        new TextRun({ text: ") yang telah menjadi salah satu aset investasi digital yang sangat diminati namun memiliki tingkat volatilitas ekstrem. Metode alokasi aset konvensional yang mengandalkan estimasi matriks korelasi historis sering kali gagal memberikan proteksi yang memadai karena kerentanannya terhadap ", font: "Times New Roman", size: 24 }),
                        new TextRun({ text: "noise", font: "Times New Roman", size: 24, italics: true }),
                        new TextRun({ text: " dan ketidakstabilan data, terutama pada saat gejolak pasar (", font: "Times New Roman", size: 24 }),
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
                        new TextRun({ text: "Kendati model Network Markowitz statis menunjukkan proteksi yang lebih relevan dibandingkan pendekatan tradisional, penentuan faktor penalti sentralitas \u03b3 (\u03b3) yang kaku kerap menimbulkan masalah di fase pasar yang dinamis (misalnya fase ", font: "Times New Roman", size: 24 }),
                        new TextRun({ text: "bullish", font: "Times New Roman", size: 24, italics: true }),
                        new TextRun({ text: "). Oleh karena itu, diperlukan pendekatan cerdas berbasis ", font: "Times New Roman", size: 24 }),
                        new TextRun({ text: "Deep Reinforcement Learning (DRL)", font: "Times New Roman", size: 24, italics: true }),
                        new TextRun({ text: ", khususnya algoritma ", font: "Times New Roman", size: 24 }),
                        new TextRun({ text: "Soft Actor-Critic (SAC)", font: "Times New Roman", size: 24, italics: true }),
                        new TextRun({ text: ", yang mampu bertindak sebagai pengontrol dinamis (", font: "Times New Roman", size: 24 }),
                        new TextRun({ text: "agent-based controller", font: "Times New Roman", size: 24, italics: true }),
                        new TextRun({ text: ") untuk menyesuaikan nilai \u03b3 secara real-time berdasarkan kondisi jaringan dan momentum pasar kripto [mendeley_cite:giudici2020network], [mendeley_cite:haarnoja2018soft].", font: "Times New Roman", size: 24 }),
                    ]
                }),
                new Paragraph({
                    alignment: AlignmentType.JUSTIFIED,
                    spacing: { before: 0, after: 120, line: 360, lineRule: "auto" },
                    indent: { firstLine: 720 },
                    children: [
                        new TextRun({ text: "Kendati demikian, integrasi model DRL yang kompleks sering kali memunculkan tantangan baru terkait transparansi keputusan aset. Tanpa adanya penjelasan yang memadai, strategi investasi yang dihasilkan oleh agen cerdas dapat dianggap sebagai 'kotak hitam' (", font: "Times New Roman", size: 24 }),
                        new TextRun({ text: "black-box", font: "Times New Roman", size: 24, italics: true }),
                        new TextRun({ text: ") yang sulit dipercaya oleh investor profesional. Oleh karena itu, penelitian ini mengusulkan penggunaan ", font: "Times New Roman", size: 24 }),
                        new TextRun({ text: "Explainable AI", font: "Times New Roman", size: 24, italics: true }),
                        new TextRun({ text: " (XAI) untuk memberikan interpretasi yang jelas terhadap logika agen SAC-Net dalam mengalokasikan bobot portofolio berdasarkan fitur-fitur jaringan pasar.", font: "Times New Roman", size: 24 }),
                    ]
                }),
                emptyLine(),                heading2("1.2 Identifikasi Masalah"),
                new Paragraph({
                    alignment: AlignmentType.JUSTIFIED,
                    spacing: { before: 0, after: 120, line: 360, lineRule: "auto" },
                    indent: { firstLine: 720 },
                    children: [
                        new TextRun({ text: "Berdasarkan latar belakang di atas, maka masalah dalam penelitian ini dapat diidentifikasi sebagai berikut:", font: "Times New Roman", size: 24 }),
                    ]
                }),
                new Paragraph({
                    numbering: { reference: "letters", level: 0 },
                    alignment: AlignmentType.JUSTIFIED,
                    spacing: { before: 0, after: 80, line: 360, lineRule: "auto" },
                    children: [
                        new TextRun({ text: "Model Network Markowitz konvensional dengan parameter penalti sentralitas (\u03b3) yang statis tidak mampu beradaptasi secara optimal terhadap perubahan fase pasar kripto yang sangat dinamis.", font: "Times New Roman", size: 24 }),
                    ]
                }),
                new Paragraph({
                    numbering: { reference: "letters", level: 0 },
                    alignment: AlignmentType.JUSTIFIED,
                    spacing: { before: 0, after: 80, line: 360, lineRule: "auto" },
                    children: [
                        new TextRun({ text: "Sifat ", font: "Times New Roman", size: 24 }),
                        new TextRun({ text: "black-box", font: "Times New Roman", size: 24, italics: true }),
                        new TextRun({ text: " pada algoritma ", font: "Times New Roman", size: 24 }),
                        new TextRun({ text: "Deep Reinforcement Learning", font: "Times New Roman", size: 24, italics: true }),
                        new TextRun({ text: " (DRL) menimbulkan kendala transparansi bagi investor profesional dalam memahami logika pengambilan keputusan alokasi bobot portofolio.", font: "Times New Roman", size: 24 }),
                    ]
                }),
                new Paragraph({
                    numbering: { reference: "letters", level: 0 },
                    alignment: AlignmentType.JUSTIFIED,
                    spacing: { before: 0, after: 80, line: 360, lineRule: "auto" },
                    children: [
                        new TextRun({ text: "Kebutuhan akan mekanisme ", font: "Times New Roman", size: 24 }),
                        new TextRun({ text: "incremental learning", font: "Times New Roman", size: 24, italics: true }),
                        new TextRun({ text: " untuk menjaga relevansi dan ketangguhan (", font: "Times New Roman", size: 24 }),
                        new TextRun({ text: "robustness", font: "Times New Roman", size: 24, italics: true }),
                        new TextRun({ text: ") model terhadap evolusi struktur jaringan pasar tanpa mengabaikan performa stabilitas jangka panjang.", font: "Times New Roman", size: 24 }),
                    ]
                }),
                emptyLine(),
                heading2("1.3 Tujuan Penelitian"),
                body("Tujuan dari penelitian ini adalah:"),
                letterItem("Merancang agen Soft Actor-Critic (SAC) sebagai Gamma Controller dinamis untuk mengoptimasi parameter penalti sentralitas pada model Network Markowitz sesuai kondisi pasar.", "letters1"),
                letterItem("Mengintegrasikan teknik Explainable AI (XAI) melalui metode SHAP untuk memberikan interpretasi transparan terhadap logika pengambilan keputusan investasi agen.", "letters1"),
                letterItem("Mengembangkan skema incremental learning untuk memastikan ketangguhan dan adaptabilitas model terhadap evolusi data jaringan pasar kripto secara berkelanjutan.", "letters1"),
                letterItem("Mengevaluasi dan membandingkan performa model yang diusulkan terhadap strategi benchmark konvensional menggunakan metrik risiko dan imbal hasil yang komprehensif.", "letters1"),
                emptyLine(),
                heading2("1.4 Ruang Lingkup Penelitian"),
                body("Ruang lingkup penelitian ini dibatasi pada:"),
                letterItem("Objek penelitian terfokus pada 9 (sembilan) aset kripto utama (BCH, BNB, BTC, EOS, ETH, LTC, TRX, XLM, XRP) dengan periode data latihan (training) tahun 2017-2019 dan data uji (testing) tahun 2024.", "letters2"),
                letterItem("Algoritma Reinforcement Learning yang digunakan adalah Soft Actor-Critic (SAC) dengan framework Stable Baselines3, menggunakan 9 state features (5 network metrics dan 4 market metrics).", "letters2"),
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
                    children: [new TextRun({ text: "BAB II TINJAUAN PUSTAKA", font: "Times New Roman", size: 24, bold: true })]
                }),
                body("Bab ini membahas kerangka teori yang relevan mencakup Teori Portofolio Modern, Complex Network Theory, dan Manajemen Risiko, serta tinjauan pustaka terhadap penelitian terdahulu di bidang Robo-Advisory kripto."),
                new Paragraph({
                    spacing: { before: 80, after: 80 },
                    children: [new TextRun({ text: "BAB III METODOLOGI PENELITIAN", font: "Times New Roman", size: 24, bold: true })]
                }),
                                body("Bab ini menjelaskan alur sistematis eksperimen yang digunakan untuk memproses data instrumen kripto, penyaringan RMT pada matriks korelasi historis, pembangunan MST, fungsi objektif Markowitz modifikasi, dan komputasi skema backtesting."),
                new Paragraph({
                    spacing: { before: 80, after: 80 },
                    children: [new TextRun({ text: "BAB IV HASIL PENELITIAN DAN PEMBAHASAN", font: "Times New Roman", size: 24, bold: true })]
                }),
                body("Bab ini menyajikan hasil-hasil yang diperoleh dari eksperimen dan pembahasan mendalam mengenai temuan penelitian."),
                new Paragraph({
                    spacing: { before: 80, after: 80 },
                    children: [new TextRun({ text: "BAB V PENUTUP", font: "Times New Roman", size: 24, bold: true })]
                }),
                body("Bab ini berisi kesimpulan dari seluruh rangkaian penelitian serta saran untuk pengembangan penelitian selanjutnya."),

            ],
            headers: {
                default: new Header({
                    children: [
                        new Paragraph({
                            alignment: AlignmentType.RIGHT,
                            children: [new TextRun({ children: [PageNumber.CURRENT], font: "Times New Roman", size: 24 })]
                        })
                    ]
                }),
                first: new Header({ children: [new Paragraph({})] })
            },
            footers: {
                first: new Footer({
                    children: [
                        new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [new TextRun({ children: [PageNumber.CURRENT], font: "Times New Roman", size: 24 })]
                        })
                    ]
                }),
                default: new Footer({ children: [new Paragraph({})] })
            }
        },
        // ==================== BAB II ====================
        {
            properties: {
                page: {
                    size: { width: 11906, height: 16838 },
                    margin: { top: 1701, right: 1417, bottom: 1417, left: 2268, header: 850, footer: 1417 },
                },
                type: SectionType.NEXT_PAGE, titlePage: true,
            },
            headers: {
                default: new Header({
                    children: [
                        new Paragraph({
                            alignment: AlignmentType.RIGHT,
                            children: [new TextRun({ children: [PageNumber.CURRENT], font: "Times New Roman", size: 24 })]
                        })
                    ]
                }),
                first: new Header({ children: [new Paragraph({})] })
            },
            footers: {
                first: new Footer({
                    children: [
                        new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [new TextRun({ children: [PageNumber.CURRENT], font: "Times New Roman", size: 24 })]
                        }),
                        guidlineFooter
                    ]
                }),
                default: new Footer({
                    children: [guidlineFooter]
                })
            },
            children: [
                chapterHeading("BAB II", "TINJAUAN PUSTAKA"),
                emptyLine(),
                body("Bab ini menyajikan ulasan komprehensif mengenai teori-teori dasar yang melandasi penelitian, tinjauan terhadap penelitian terdahulu yang relevan, serta kerangka pemikiran yang menghubungkan variabel-variabel penelitian."),
                emptyLine(),
                heading2("2.1 Tinjauan Teori"),
                body("Tinjauan teori mencakup landasan akademis utama yang digunakan dalam membangun model optimasi portofolio cerdas dalam penelitian ini."),
                emptyLine(),
                heading3("2.1.1 Modern Portfolio Theory (Markowitz)"),
                mixedBody([
                    {text: "Teori Portofolio Modern (MPT), yang dipelopori oleh Harry Markowitz, berupaya memaksimalkan imbal hasil yang diharapkan ("},
                    {text: "expected return", italic: true},
                    {text: ") pada tingkat risiko ("},
                    {text: "variance", italic: true},
                    {text: ") tertentu melalui pemilihan aset yang terdiversifikasi dalam sebuah "},
                    {text: "Efficient Frontier", italic: true},
                    {text: " [mendeley_cite:markowitz1952portfolio]. Namun, dalam praktiknya, MPT memiliki keterbatasan signifikan terkait stabilitas estimasi. Fenomena "},
                    {text: "Markowitz Curse", italic: true},
                    {text: " atau enigma optimasi Markowitz [mendeley_cite:michaud1989markowitz] terjadi karena algoritma optimasi cenderung memperkuat kesalahan estimasi ("},
                    {text: "error maximization", italic: true},
                    {text: "), sehingga perubahan kecil pada input data dapat menghasilkan perubahan drastis pada alokasi bobot portofolio."}
                ]),
                emptyLine(),
                heading3("2.1.2 Random Matrix Theory dan Distribusi Marchenko-Pastur"),
                mixedBody([
                    {text: "Teori "},
                    {text: "Random Matrix", italic: true},
                    {text: " (RMT) menyediakan kerangka matematis untuk membedakan antara korelasi yang signifikan secara ekonomi dengan korelasi yang muncul murni karena kebetulan ("},
                    {text: "noise", italic: true},
                    {text: ") pada matriks korelasi berdimensi tinggi. Dalam konteks portofolio kripto, pembersihan "},
                    {text: "noise", italic: true},
                    {text: " sangat krusial karena estimasi korelasi sering kali tidak stabil akibat fenomena "},
                    {text: "Markowitz Curse", italic: true},
                    {text: ". Distribusi Marchenko-Pastur mendefinisikan batas teoretis bagi nilai eigen (\u03BB) dari matriks korelasi acak murni sebagai berikut:"}
                ]),
                formulaParagraph(rumus("\\lambda_{\\pm}"), "2.1"),
                bodyNoIndent("Keterangan:"),
                formulaKeterangan([
                    [[new MathSubScript({ children: [new MathRun("\u03BB")], subScript: [new MathRun("\u00B1")] })], "Batas atas (+) dan bawah (-) nilai eigen dari noise."],
                    ["\u03C3\u00B2", "Variansi dari elemen matriks korelasi acak."],
                    ["Q", "Rasio antara jumlah observasi (T) terhadap jumlah aset (N)."]
                ]),
                emptyLine(),
                mixedBody([
                    {text: "Nilai eigen yang berada di atas batas "},
                    {text: "\u03BB+", italic: true},
                    {text: " dianggap sebagai "},
                    {text: "market mode", italic: true},
                    {text: " atau sinyal ekonomi sejati yang mengandung informasi struktural pasar. Sebaliknya, nilai eigen di bawah batas tersebut dikategorikan sebagai "},
                    {text: "noise", italic: true},
                    {text: " yang tidak memiliki signifikansi finansial. Dengan menerapkan filtrasi ini melalui metode "},
                    {text: "Eigenvalue Clipping", italic: true},
                    {text: ", topologi jaringan (MST) yang dihasilkan menjadi lebih stabil dan representatif terhadap struktur pasar sesungguhnya [mendeley_cite:marchenko1967distribution]."}
                ]),

                emptyLine(),
                heading3("2.1.3 Teori Risiko Koheren (Coherent Risk Measures)"),
                mixedBody([
                    {text: "Metrik risiko tradisional seperti variansi sering kali meremehkan risiko pada pasar kripto yang memiliki distribusi "},
                    {text: "fat-tail", italic: true},
                    {text: ". Oleh karena itu, penelitian ini mengacu pada "},
                    {text: "Teori Risiko Koheren", bold: true},
                    {text: " yang diperkenalkan oleh Artzner et al. (1999), yang menyatakan bahwa metrik risiko yang baik harus memenuhi empat aksioma: "},
                    {text: "sub-additivity, homogeneity, monotonicity,", italic: true},
                    {text: " dan "},
                    {text: "translation invariance", italic: true},
                    {text: "."}
                ]),
                emptyLine(),
                mixedBody([
                    {text: "Conditional Value at Risk (CVaR)", bold: true},
                    {text: ", atau "},
                    {text: "Expected Shortfall", italic: true},
                    {text: ", adalah metrik risiko koheren yang mengukur rata-rata kerugian pada ekor distribusi (kejadian ekstrem) yang melampaui ambang batas "},
                    {text: "Value at Risk", italic: true},
                    {text: " (VaR). Rumus CVaR didefinisikan sebagai:"}
                ]),
                formulaParagraph(rumus("CVaR"), "2.2"),
                bodyNoIndent("Keterangan:"),
                formulaKeterangan([
                    [[new MathSubScript({ children: [new MathRun("CVaR")], subScript: [new MathRun("\u03B1")] })], "Conditional Value at Risk pada tingkat kepercayaan \u03B1."],
                    ["E", "Operator ekspektasi atau nilai rata-rata."],
                    ["L", "Besaran kerugian portofolio."],
                    [[new MathSubScript({ children: [new MathRun("VaR")], subScript: [new MathRun("\u03B1")] })], "Value at Risk pada tingkat kepercayaan \u03B1."]
                ]),
                emptyLine(),
                emptyLine(),
                mixedBody([
                    {text: "Selain itu, untuk mengukur efisiensi imbal hasil terhadap risiko kerugian yang sesungguhnya (bukan sekadar volatilitas total), digunakan "},
                    {text: "Sortino Ratio", bold: true},
                    {text: ". Berbeda dengan Sharpe Ratio, Sortino Ratio hanya mempertimbangkan deviasi negatif ("},
                    {text: "downside deviation", italic: true},
                    {text: ") sebagai penyebut, sehingga memberikan gambaran yang lebih akurat mengenai performa portofolio dalam menghadapi risiko penurunan harga:"}
                ]),
                formulaParagraph(rumus("Sortino"), "2.3"),
                bodyNoIndent("Keterangan:"),
                formulaKeterangan([
                    ["Rp", "Imbal hasil (return) rata-rata portofolio."],
                    ["Rf", "Tingkat imbal hasil bebas risiko (risk-free rate)."],
                    ["\u03C3d", "Downside deviation (volatilitas pada imbal hasil negatif)."]
                ]),
                emptyLine(),
                emptyLine(),
                heading3("2.1.4 Topologi Jaringan Keuangan dan Risiko Sistemik"),
                mixedBody([
                    {text: "Teori jaringan kompleks memungkinkan pemodelan interaksi antar aset keuangan sebagai sistem yang dinamis dan saling terhubung. Dalam perspektif ini, pasar direpresentasikan sebagai graf dimana aset adalah simpul ("},
                    {text: "nodes", italic: true},
                    {text: ") dan korelasi harga adalah sisi ("},
                    {text: "edges", italic: true},
                    {text: "). Struktur "},
                    {text: "Star-like", italic: true},
                    {text: " yang didominasi oleh aset dengan sentralitas tinggi menunjukkan kerentanan sistemik yang besar, dimana kegagalan pada satu aset pusat dapat memicu penularan finansial ("},
                    {text: "financial contagion", italic: true},
                    {text: ") ke seluruh jaringan [mendeley_cite:giudici2020network]. Sebaliknya, topologi yang bersifat "},
                    {text: "Distributed", italic: true},
                    {text: " menawarkan ketahanan dan manfaat diversifikasi yang lebih stabil. Penggunaan algoritma "},
                    {text: "Minimum Spanning Tree (MST)", italic: true},
                    {text: " menjadi krusial untuk menyaring informasi korelasi yang paling signifikan dan mengidentifikasi 'tulang punggung' jaringan pasar, sehingga model Network Markowitz dapat secara efektif memberikan penalti terhadap aset-aset yang berada pada posisi kritis guna memitigasi risiko sistemik [mendeley_cite:mantegna1999hierarchical]."}
                ]),
                emptyLine(),
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [new TextRun({ text: "[[%IMAGE_TOPOLOGY]]", font: "Times New Roman", size: 22 })]
                }),
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [new TextRun({ text: "Gambar II.1. Ilustrasi Topologi Finansial Star-like vs Distributed", font: "Times New Roman", size: 22, bold: true })]
                }),
                emptyLine(),
                heading3("2.1.5 Network Markowitz"),
                mixedBody([
                    {text: "Integrasi variabel jaringan ke dalam fungsi objektif Markowitz memungkinkan model untuk menghukum aset dengan keterhubungan sistemik tinggi menggunakan faktor penalti \u03b3 [mendeley_cite:giudici2020network]. Hal ini memastikan portofolio tangguh terhadap dinamika struktur jaringan pasar kripto."}
                ]),
                emptyLine(),
                heading3("2.1.6 Adaptive Market Hypothesis (AMH)"),
                mixedBody([
                    {text: "AMH menyatakan bahwa efisiensi pasar bukanlah kondisi statis [mendeley_cite:lo2004adaptive]. Kondisi pasar kripto yang bervariasi memberikan landasan bagi penggunaan "},
                    {text: "Deep Reinforcement Learning (DRL)", italic: true},
                    {text: " khususnya algoritma "},
                    {text: "Soft Actor-Critic (SAC)", italic: true},
                    {text: " untuk melakukan kalibrasi parameter \u03b3 secara dinamis."}
                ]),
                emptyLine(),
                heading3("2.1.7 Walk-forward Analysis"),
                mixedBody([
                    {text: "Metode "},
                    {text: "rolling window", italic: true},
                    {text: " dengan skema rebalancing mingguan (7 hari) dan jendela observasi 30 hari digunakan untuk menghindari "},
                    {text: "look-ahead bias", italic: true},
                    {text: " dan memastikan hasil backtesting mencerminkan realitas pasar."}
                ]),
                emptyLine(),
                heading2("2.2 Tinjauan Penelitian Sebelumnya"),
                body("Beberapa penelitian terdahulu yang menjadi rujukan dalam penelitian ini dirangkum dalam tabel perbandingan berikut:"),
                emptyLine(),
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [new TextRun({ text: "Tabel II.1. Perbandingan Penelitian Terdahulu (Graph-Based)", font: "Times New Roman", size: 24, bold: true })]
                }),
                new Table({
                    alignment: AlignmentType.CENTER,
                    width: { size: 9500, type: WidthType.DXA },
                    columnWidths: [1700, 1300, 1300, 1300, 1100, 1200, 1600],
                    rows: [
                        new TableRow({
                            tableHeader: true,
                            children: [
                                new TableCell({ shading: { fill: "D5E8F0" }, borders, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Peneliti", font: "Times New Roman", size: 20, bold: true })] })] }),
                                new TableCell({ shading: { fill: "D5E8F0" }, borders, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Model learning", font: "Times New Roman", size: 20, bold: true })] })] }),
                                new TableCell({ shading: { fill: "D5E8F0" }, borders, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Graph method", font: "Times New Roman", size: 20, bold: true })] })] }),
                                new TableCell({ shading: { fill: "D5E8F0" }, borders, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Adaptivitas", font: "Times New Roman", size: 20, bold: true })] })] }),
                                new TableCell({ shading: { fill: "D5E8F0" }, borders, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Aset", font: "Times New Roman", size: 20, bold: true })] })] }),
                                new TableCell({ shading: { fill: "D5E8F0" }, borders, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "XAI", font: "Times New Roman", size: 20, bold: true })] })] }),
                                new TableCell({ shading: { fill: "D5E8F0" }, borders, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Metrik", font: "Times New Roman", size: 20, bold: true })] })] }),
                            ]
                        }),
                        ...[
                            ["Giudici (2020)", "Mean-Variance", "MST + Centrality", "Statis", "Kripto", "Tidak ada", "Sharpe, MDD"],
                            ["Korangi (2024)", "GAT (Supervised)", "TMFG + DistCorr", "Rolling", "Saham", "Tidak ada", "Sharpe, Return"],
                            ["Jing (2023)", "MPT", "MST", "Statis", "Kripto", "N/A", "Sharpe"],
                            ["Ioannidis (2023)", "Centrality Weights", "Transfer Entropy", "Multi-horizon", "Saham", "Centrality XAI", "Sharpe, Return"],
                            ["Wang (2023)", "Inverse Centrality", "SR-IFN Filter", "Bootstrap", "Saham", "Interpret.", "Sharpe"],
                            ["Takahashi (2025)", "WMIS Opt.", "Max Indep. Set", "Sim. Bifurc.", "Saham", "Set Structure", "Return, Risk"],
                            ["Choudhary (2025)", "Multi-Reward DRL", "N/A (Feature)", "Dyn. Reward", "Saham", "N/A", "Sharpe, MDD"],
                        ].map(([p, m, g, a, as, x, mt]) => new TableRow({
                            children: [
                                new TableCell({ borders, children: [new Paragraph({ children: [new TextRun({ text: p, font: "Times New Roman", size: 18 })] })] }),
                                new TableCell({ borders, children: [new Paragraph({ children: [new TextRun({ text: m, font: "Times New Roman", size: 18 })] })] }),
                                new TableCell({ borders, children: [new Paragraph({ children: [new TextRun({ text: g, font: "Times New Roman", size: 18 })] })] }),
                                new TableCell({ borders, children: [new Paragraph({ children: [new TextRun({ text: a, font: "Times New Roman", size: 18 })] })] }),
                                new TableCell({ borders, children: [new Paragraph({ children: [new TextRun({ text: as, font: "Times New Roman", size: 18 })] })] }),
                                new TableCell({ borders, children: [new Paragraph({ children: [new TextRun({ text: x, font: "Times New Roman", size: 18 })] })] }),
                                new TableCell({ borders, children: [new Paragraph({ children: [new TextRun({ text: mt, font: "Times New Roman", size: 18 })] })] }),
                            ]
                        })),
                        new TableRow({
                            children: [
                                new TableCell({ shading: { fill: "F0F0F0" }, borders, children: [new Paragraph({ children: [new TextRun({ text: "Penelitian Ini", font: "Times New Roman", size: 18, bold: true })] })] }),
                                new TableCell({ shading: { fill: "F0F0F0" }, borders, children: [new Paragraph({ children: [new TextRun({ text: "DRL (SAC)", font: "Times New Roman", size: 18, bold: true })] })] }),
                                new TableCell({ shading: { fill: "F0F0F0" }, borders, children: [new Paragraph({ children: [new TextRun({ text: "RMT+MST+Eigen", font: "Times New Roman", size: 18, bold: true })] })] }),
                                new TableCell({ shading: { fill: "F0F0F0" }, borders, children: [new Paragraph({ children: [new TextRun({ text: "Adaptive Gamma", font: "Times New Roman", size: 18, bold: true })] })] }),
                                new TableCell({ shading: { fill: "F0F0F0" }, borders, children: [new Paragraph({ children: [new TextRun({ text: "Kripto (9 aset)", font: "Times New Roman", size: 18, bold: true })] })] }),
                                new TableCell({ shading: { fill: "F0F0F0" }, borders, children: [new Paragraph({ children: [new TextRun({ text: "Post-hoc (SHAP)", font: "Times New Roman", size: 18, bold: true })] })] }),
                                new TableCell({ shading: { fill: "F0F0F0" }, borders, children: [new Paragraph({ children: [new TextRun({ text: "Sharpe, CVaR, Sortino", font: "Times New Roman", size: 18, bold: true })] })] }),
                            ]
                        }),
                    ]
                }),
                emptyLine(),
                heading2("2.3 Kerangka Konsep"),
                body("Kerangka konsep menggambarkan alur logika pemecahan masalah melalui integrasi metodologi DRL dan analisis jaringan pasar."),
                emptyLine(),
                heading3("2.3.1 Analisis Kebaruan (Gap Analysis)"),
                mixedBody([
                    {text: "Kebaruan penelitian ini terletak pada: (1) penggunaan algoritma "},
                    {text: "Soft Actor-Critic (SAC)", italic: true},
                    {text: " sebagai pengontrol dinamis nilai \u03b3; (2) penerapan "},
                    {text: "multi-seed validation", italic: true},
                    {text: " (seed 42, 123, 77); (3) perbandingan dengan baseline "},
                    {text: "Equal Risk Contribution", italic: true},
                    {text: " (ERC) dan "},
                    {text: "Classic Mean-Variance", italic: true},
                    {text: "; serta (4) integrasi "},
                    {text: "Explainable AI (SHAP)", italic: true},
                    {text: " untuk transparansi keputusan agen RL."}
                ]),
            ],
        },

        // ==================== BAB III ====================
        {
            properties: {
                page: {
                    size: { width: 11906, height: 16838 },
                    margin: { top: 1701, right: 1417, bottom: 1417, left: 2268, header: 850, footer: 1417 }
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
                    children: [new TextRun({ text: "Gambar III.1. Kerangka Kerja Penelitian SAC-Net Markowitz", font: "Times New Roman", size: 22, bold: true })]
                }),
                emptyLine(),
                mixedBody([
                    {text: "Secara visual, alur penelitian pada Gambar III.1 dibagi menjadi lima tahapan utama: (1) "},
                    {text: "Data Acquisition", italic: true},
                    {text: " yaitu pengumpulan data historis kripto; (2) "},
                    {text: "Feature Engineering", italic: true},
                    {text: " mencakup filtrasi RMT, pembangunan MST, dan ekstraksi 9 fitur indikator; (3) "},
                    {text: "SAC Agent Training", italic: true},
                    {text: " melatih pengontrol \u03b3 berbasis Deep Reinforcement Learning; (4) "},
                    {text: "Backtesting & Rebalancing", italic: true},
                    {text: " simulasi perdagangan mingguan; dan (5) "},
                    {text: "Performance Evaluation", italic: true},
                    {text: " menggunakan metrik Sharpe dan Calmar."}
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
                    {text: ", menentukan parameter dasar optimasi, serta memahami algoritma "},
                    {text: "Soft Actor-Critic (SAC)", italic: true},
                    {text: " dalam konteks manajemen portofolio dinamis."}
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
                numItem("Feature Engineering"),
                mixedBody([
                    {text: "Proses ekstraksi 9 fitur utama sebagai input (", italic: true},
                    {text: "state", italic: true},
                    {text: ") bagi agen RL. Fitur ini mencakup 5 indikator jaringan (varian sentralitas, densitas, dan MST distance) serta 4 indikator pasar (momentum, volatilitas jangka pendek, dan return rata-rata)."}
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
                numItem("SAC Agent Training"),
                mixedBody([
                    {text: "Melatih pengontrol \u03b3 dinamis menggunakan algoritma "},
                    {text: "Soft Actor-Critic (SAC)", italic: true},
                    {text: " dengan fungsi reward berbasis "},
                    {text: "Incremental Sharpe", italic: true},
                    {text: ". Agen dilatih untuk meminimalkan risiko penularan sistemik sekaligus memaksimalkan efisiensi imbal hasil portofolio."}
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
                bulletItem("Framework/Library: Pandas, Numpy, Scipy (Optimization), NetworkX (Graph Analytics), Stable Baselines3 (Deep Reinforcement Learning)"),
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
                    {text: " harian 9 aset; (2) "},
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
                    {text: " dikeluarkan dari dataset penelitian ini. Meskipun USDT merupakan aset teratas berdasarkan kapitalisasi pasar, sifatnya yang dipatok ke USD (dengan volatilitas mendekati nol) dapat menyebabkan bias dalam perhitungan matriks korelasi dan penalti sentralitas, sehingga penghapusannya memungkinkan model untuk berfokus sepenuhnya pada interaksi risiko antar aset volatil."}
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

                formulaParagraph(rumus("\\min_w w^T \\Sigma^* w + \\gamma \\sum_{i=1}^n x_i w_i"), "3.1"),
                bodyNoIndent("Keterangan:"),
                formulaKeterangan([
                    ["w", "Vektor alokasi bobot untuk setiap aset kripto."],
                    [[new MathSuperScript({ children: [new MathRun("\u03A3")], superScript: [new MathRun("*")] })], "Matriks Kovarians terfilter RMT."],
                    ["\u03B3", "Parameter skalar penghukuman sentralitas graf."],
                    [[new MathSubScript({ children: [new MathRun("x")], subScript: [new MathRun("i")] })], "Vektor skor Eigenvector Centrality tiap node aset."]
                ]),
                emptyLine(),
                heading3("3.4.1 Strategi Portofolio dan Benchmark"),
                body("Penelitian ini membandingkan empat strategi utama untuk mengevaluasi performa model yang diusulkan terhadap standar industri dan metodologi mutakhir:"),
                
                bulletItem("Equally Weighted (EW): Strategi alokasi 1/N yang memberikan bobot yang sama ke setiap aset tanpa mempertimbangkan parameter risiko atau imbal hasil. EW berfungsi sebagai 'benchmark naif' yang sangat tangguh karena tidak memiliki risiko estimasi (estimation risk)."),
                
                bulletItem("Classical Markowitz (CM): Model optimasi Mean-Variance standar yang berupaya meminimalkan variansi portofolio untuk tingkat imbal hasil tertentu. CM bertindak sebagai representasi teori portofolio tradisional yang sering kali menderita masalah ketidakstabilan numerik pada data historis yang berisik."),
                
                bulletItem("Network Markowitz (NW): Model jaringan original (Giudici et al., 2020) yang menggunakan parameter penalti sentralitas (\u03b3) statis. NW digunakan sebagai pembanding langsung untuk menunjukkan sejauh mana penambahan fitur 'Dynamic Grid Search' pada model yang diusulkan dapat meningkatkan performa portofolio dibandingkan model jaringan dasar."),
                emptyLine(),

                mixedBody([
                    {text: "Pada penelitian ini, intensitas penalti \u03b3 tidak akan dilakukan "},
                    {text: "hard-coded", italic: true},
                    {text: " statis, melainkan agen cerdas berbasis "},
                    {text: "Soft Actor-Critic (SAC)", italic: true},
                    {text: " akan mempelajari relasi antara 9 fitur input terhadap keputusan pergeseran parameter \u03b3 untuk memaksimalkan profil risiko-imbal hasil di masa depan."}
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
                formulaParagraph(rumus("Sharpe\\ Ratio = \\frac{R_p - R_f}{\\sigma_p}"), "3.2"),
                bodyNoIndent("Keterangan:"),
                formulaKeterangan([
                    [[new MathSubScript({ children: [new MathRun("R")], subScript: [new MathRun("p")] })], "Imbal hasil (return) portofolio."],
                    [[new MathSubScript({ children: [new MathRun("R")], subScript: [new MathRun("f")] })], "Tingkat imbal hasil bebas risiko (risk-free rate)."],
                    [[new MathSubScript({ children: [new MathRun("\u03C3")], subScript: [new MathRun("p")] })], "Standar deviasi dari imbal hasil berlebih portofolio."]
                ]),
                emptyLine(),
                mixedBody([
                    {text: "Nilai Sharpe Ratio > 1 dianggap baik, > 2 sangat baik, dan > 3 luar biasa. Semakin besar nilai "},
                    {text: "Sharpe Ratio", italic: true},
                    {text: ", semakin baik kualitas portofolio karena menunjukkan imbal hasil yang lebih tinggi untuk setiap unit risiko yang diambil. Sebaliknya, semakin kecil nilai ini, semakin tidak efisien portofolio tersebut dalam menghasilkan imbal hasil terhadap risikonya [mendeley_cite:markowitz1952portfolio], [mendeley_cite:lopezdeprado2018advances]."}
                ]),
                emptyLine(),
                
                
                heading3("3.5.4 Calmar Ratio"),
                mixedBody([
                    {text: "Calmar Ratio digunakan untuk mengevaluasi imbal hasil tahunan relatif terhadap penarikan maksimum (", font: "Times New Roman", size: 24 },
                    {text: "Maximum Drawdown", italic: true},
                    {text: "). Metrik ini krusial dalam dunia kripto untuk menguji ketahanan portofolio terhadap kejatuhan harga parah. Formula Calmar Ratio adalah:", font: "Times New Roman", size: 24 }
                ]),
                new Table({
                    width: { size: 8200, type: WidthType.DXA },
                    columnWidths: [4000, 4200],
                    borders: noBorders,
                    rows: [
                        new TableRow({
                            children: [
                                new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Calmar Ratio = Ann. Return / |Max Drawdown|", font: "Times New Roman", size: 24, bold: true })] })] }),
                            ]
                        }),
                    ]
                }),

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
                }),
                first: new Header({ children: [new Paragraph({})] })
            },
            footers: {
                first: new Footer({
                    children: [
                        new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [new TextRun({ children: [PageNumber.CURRENT], font: "Times New Roman", size: 24 })]
                        }),
                        guidlineFooter
                    ]
                }),
                default: new Footer({
                    children: [guidlineFooter]
                })
            }
        },
        // ==================== BAB IV ====================
        {
            properties: {
                page: {
                    size: { width: 11906, height: 16838 },
                    margin: { top: 1701, right: 1417, bottom: 1417, left: 2268, header: 850, footer: 1417 }
                },
                type: SectionType.NEXT_PAGE, titlePage: true
            },
            headers: {
                default: new Header({
                    children: [
                        new Paragraph({
                            alignment: AlignmentType.RIGHT,
                            children: [new TextRun({ children: [PageNumber.CURRENT], font: "Times New Roman", size: 24 })]
                        })
                    ]
                }),
                first: new Header({ children: [new Paragraph({})] })
            },
            footers: {
                first: new Footer({
                    children: [
                        new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [new TextRun({ children: [PageNumber.CURRENT], font: "Times New Roman", size: 24 })]
                        }),
                        guidlineFooter
                    ]
                }),
                default: new Footer({
                    children: [guidlineFooter]
                })
            },
            children: [
                chapterHeading("BAB IV", "HASIL PENELITIAN DAN PEMBAHASAN"),
                emptyLine(),
                body("Bab ini menyajikan hasil-hasil yang diperoleh dari eksperimen dan pembahasan mendalam mengenai temuan penelitian."),
                emptyLine(),
                heading2("4.1 Deskripsi Objek Penelitian"),
                body("[Outline: Penjelasan mengenai data yang digunakan dalam pengujian/backtesting]"),
                emptyLine(),
                heading2("4.2 Hasil Penelitian"),
                body("[Outline: Presentasi hasil metrik performa portofolio (Sharpe, CVaR, Sortino)]"),
                emptyLine(),
                heading2("4.3 Pembahasan"),
                body("[Outline: Analisis kelebihan dan kekurangan model dibandingkan benchmark]"),
            ]
        },
        // ==================== BAB V ====================
        {
            properties: {
                page: {
                    size: { width: 11906, height: 16838 },
                    margin: { top: 1701, right: 1417, bottom: 1417, left: 2268, header: 850, footer: 1417 }
                },
                type: SectionType.NEXT_PAGE, titlePage: true
            },
            headers: {
                default: new Header({
                    children: [
                        new Paragraph({
                            alignment: AlignmentType.RIGHT,
                            children: [new TextRun({ children: [PageNumber.CURRENT], font: "Times New Roman", size: 24 })]
                        })
                    ]
                }),
                first: new Header({ children: [new Paragraph({})] })
            },
            footers: {
                first: new Footer({
                    children: [
                        new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [new TextRun({ children: [PageNumber.CURRENT], font: "Times New Roman", size: 24 })]
                        }),
                        guidlineFooter
                    ]
                }),
                default: new Footer({
                    children: [guidlineFooter]
                })
            },
            children: [
                chapterHeading("BAB V", "PENUTUP"),
                emptyLine(),
                body("Bab ini berisi kesimpulan dari seluruh rangkaian penelitian serta saran untuk pengembangan penelitian selanjutnya."),
                emptyLine(),
                heading2("5.1 Kesimpulan"),
                body("[Outline: Rangkuman jawaban atas rumusan masalah]"),
                emptyLine(),
                heading2("5.2 Saran"),
                body("[Outline: Rekomendasi untuk penelitian di masa mendatang]"),
            ]
        },
        // ==================== DAFTAR REFERENSI ====================

        {
            properties: {
                page: {
                    size: { width: 11906, height: 16838 },
                    margin: { top: 1701, right: 1417, bottom: 1417, left: 2268, header: 850, footer: 1417 }
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
                }),
                first: new Header({ children: [new Paragraph({})] })
            },
            footers: {
                first: new Footer({
                    children: [
                        new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [new TextRun({ children: [PageNumber.CURRENT], font: "Times New Roman", size: 24 })]
                        }),
                        guidlineFooter
                    ]
                }),
                default: new Footer({
                    children: [guidlineFooter]
                })
            }
        },
    ]
});

Packer.toBuffer(doc).then(buffer => {
    fs.writeFileSync("proposal_tesis_ragil.docx", buffer);
    console.log("Done!");
});
