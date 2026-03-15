const docx = require('docx');
const { Document, Math, MathRun, Packer } = docx;
const fs = require('fs');

const doc = new Document({
    sections: [{
        children: [
            new docx.Paragraph({
                children: [
                    new Math({
                        children: [
                            new MathRun("\\min_w w^T \\Sigma^* w + \\gamma \\sum_{i=1}^n x_i w_i")
                        ]
                    })
                ]
            })
        ]
    }]
});

Packer.toBuffer(doc).then(buffer => {
    fs.writeFileSync("test_latex_string.docx", buffer);
    console.log("Done");
});
