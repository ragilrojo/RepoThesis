import re
import os

filepath = r'e:\ProjectNodeJs\temp_doc_build\generate_proposal.js'
with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Bridging paragraphs
content = content.replace(
    'chapterHeading("BAB I", "PENDAHULUAN"),\n                emptyLine(),\n                heading2("1.1. Latar Belakang"),',
    'chapterHeading("BAB I", "PENDAHULUAN"),\n                emptyLine(),\n                body("Bab ini membahas secara komprehensif latar belakang permasalahan, perumusan penelitian, tujuan, dan ruang lingkup yang menjadi batasan penelitian ini."),\n                emptyLine(),\n                heading2("1.1. Latar Belakang"),'
)

content = content.replace(
    'chapterHeading("BAB II", "LANDASAN/KERANGKA PEMIKIRAN"),\n                emptyLine(),\n                heading2("2.1. Kerangka Teori"),\n                heading3("2.1.1. Modern Portfolio Theory (Markowitz)"),',
    'chapterHeading("BAB II", "LANDASAN/KERANGKA PEMIKIRAN"),\n                emptyLine(),\n                body("Bab ini menguraikan berbagai teori dasar, konsep, serta tinjauan pustaka dari penelitian terdahulu yang menjadi landasan pemikiran bagi pengembangan sistem portofolio dalam penelitian ini."),\n                emptyLine(),\n                heading2("2.1. Modern Portfolio Theory (Markowitz)"),'
)

content = content.replace(
    'chapterHeading("BAB III", "METODOLOGI PENELITIAN"),\n                emptyLine(),\n                heading2("3.1. Tahapan Penelitian"),',
    'chapterHeading("BAB III", "METODOLOGI PENELITIAN"),\n                emptyLine(),\n                body("Bab ini menyajikan metodologi yang diaplikasikan dalam penelitian ini, mencakup rancangan tahapan logis penelitian, persiapan data, serta metrik evaluasi performa."),\n                emptyLine(),\n                heading2("3.1. Tahapan Penelitian"),'
)

# 2. Change Chapter 2 2.1.x to 2.x
content = content.replace('heading3("2.1.2.', 'heading2("2.2.')
content = content.replace('heading3("2.1.3.', 'heading2("2.3.')
content = content.replace('heading3("2.1.4.', 'heading2("2.4.')
content = content.replace('heading3("2.1.5.', 'heading2("2.5.')
content = content.replace('heading3("2.1.6.', 'heading2("2.6.')
content = content.replace('heading3("2.1.7.', 'heading2("2.7.')
content = content.replace('heading3("2.1.8.', 'heading2("2.8.')
content = content.replace('heading3("2.1.9.', 'heading2("2.9.')

# 3. Add Axiom definitions (around 2.3)
to_find_axiom = 'fat-tail events", italic: true},\n                    {text: ") di pasar kripto."}\n                ]),\n                emptyLine(),\n                heading2("2.4. Topologi'
if to_find_axiom in content:
    replacement_axiom = 'fat-tail events", italic: true},\n                    {text: ") di pasar kripto."}\n                ]),\n                bulletItem("Monotonicity: Jika portofolio X selalu tidak lebih buruk dari Y, maka risiko X harus lebih kecil atau sama dari Y (\\u03c1(X) \\u2264 \\u03c1(Y) untuk X \\u2265 Y)."),\n                bulletItem("Sub-additivity: Risiko gabungan tidak boleh lebih dari jumlah risiko masing-masing (\\u03c1(X + Y) \\u2264 \\u03c1(X) + \\u03c1(Y)). Ini adalah kaidah efek diversifikasi."),\n                bulletItem("Homogeneity: Menambah kelipatan ukuran posisi sejalan dengan mengalikan besaran risikonya (\\u03c1(cX) = c \\u03c1(X) untuk c > 0)."),\n                bulletItem("Translational Invariance: Menambah sejumlah modal pasti bebas risiko ke portofolio akan mengurangi risiko sebesar persis nilai nominal tersebut."),\n                emptyLine(),\n                heading2("2.4. Topologi'
    content = content.replace(to_find_axiom, replacement_axiom)
else:
    print("Axiom not found")

# 4. Topology illustrations (in 2.4, before 2.5)
to_find_topo = "periferi jaringan', sehingga memitigasi risiko kegagalan sistemik [5], [12].\"}\n                ]),\n                emptyLine(),\n                heading2(\"2.5. Network Markowitz\"),"
if to_find_topo in content:
    replacement_topo = 'periferi jaringan\', sehingga memitigasi risiko kegagalan sistemik [5], [12]."}\n                ]),\n                emptyLine(),\n                new Paragraph({\n                    alignment: AlignmentType.CENTER,\n                    children: [\n                        new TextRun({\n                            text: "[[%IMAGE_TOPOLOGY]]",\n                            font: "Times New Roman",\n                            size: 22,\n                        })\n                    ]\n                }),\n                emptyLine(),\n                new Paragraph({\n                    alignment: AlignmentType.CENTER,\n                    children: [\n                        new TextRun({ text: "Gambar II.1. Ilustrasi Topologi Finansial Star-like vs Distributed", font: "Times New Roman", size: 22, bold: true })\n                    ]\n                }),\n                emptyLine(),\n                heading2("2.5. Network Markowitz"),'
    content = content.replace(to_find_topo, replacement_topo)
else:
    print("Topology not found")


# 5. Fix Reference numbers instead of years.
content = content.replace('Giudici et al. (2020)', 'Giudici et al. [1]')
content = content.replace('Kitanovski dkk. (2022)', 'Kitanovski et al. [8]')
content = content.replace('Jing dan Rocha (2023)', 'Jing dan Rocha [6]')
content = content.replace('Kitanovski dkk. (2024)', 'Kitanovski et al. [7]')
content = content.replace('Jing dkk. (2025)', 'Jing et al. [21]') 
content = content.replace('Artzner et al. (1999) [19]', 'Artzner et al. [19]')
content = content.replace('Andrew Lo (2004) [20]', 'Andrew Lo [20]')

content = content.replace('Jing, et al. (2025)', 'Jing, et al. [21]')
content = content.replace('Giudici, et al. (2020)', 'Giudici, et al. [1]')
content = content.replace('Kitanovski, et al. (2022)', 'Kitanovski, et al. [8]')
content = content.replace('Jing & Rocha (2023)', 'Jing & Rocha [6]')
content = content.replace('Kitanovski, et al. (2024)', 'Kitanovski, et al. [7]')

content = content.replace('distribusi Marchenko-Pastur [21]', 'distribusi Marchenko-Pastur [3]')

# Fix references list:
if '[18] A. Arratia' in content:
    content = content.replace(
        'children: [new TextRun({ text: "[18] A. Arratia,',
        'children: [new TextRun({ text: "[17] S. T. Rachev, S. V. Stoyanov, and F. J. Fabozzi, \\"Advanced Stochastic Models, Risk Assessment, and Portfolio Optimization,\\" John Wiley & Sons, 2008.", font: "Times New Roman", size: 24 })]\n                }),\n                new Paragraph({\n                    alignment: AlignmentType.JUSTIFIED,\n                    spacing: { before: 0, after: 120, line: 360 },\n                    indent: { left: 720, hanging: 720 },\n                    children: [new TextRun({ text: "[18] A. Arratia,'
    )
else:
    print("Ref 18 not found")

if '[21] V. A. Marchenko' in content:
    content = content.replace(
        '[21] V. A. Marchenko and L. A. Pastur, \\"Distribution of eigenvalues for some sets of random matrices,\\" Matematicheskii Sbornik, vol. 114, no. 4, pp. 507-536, 1967.',
        '[21] R. Jing, X. Zhao, and L. E. C. Rocha, \\"Optimising cryptocurrency portfolios through stable clustering of price correlation networks,\\" arXiv preprint arXiv:2505.24831, 2025.'
    )
else:
    print("Ref 21 not found")

with open(filepath, 'w', encoding='utf-8') as f:
    f.write(content)
print("Done")
