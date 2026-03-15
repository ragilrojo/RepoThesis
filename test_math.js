const docx = require('docx');
const { Document, Math, MathRun, MathSubScript, Packer } = docx;
const fs = require('fs');

try {
    const doc = new Document({
        sections: [{
            children: [
                new docx.Paragraph({
                    children: [
                        new Math({
                            children: [
                                new MathSubScript({
                                    main: new MathRun("A"),
                                    subScript: new MathRun("b"),
                                })
                            ]
                        })
                    ]
                })
            ]
        }]
    });
    console.log("Constructor worked with single items");
} catch (e) {
    console.log("Failed with single items:", e.message);
}

try {
    const doc = new Document({
        sections: [{
            children: [
                new docx.Paragraph({
                    children: [
                        new Math({
                            children: [
                                new MathSubScript({
                                    main: [new MathRun("A")],
                                    subScript: [new MathRun("b")],
                                })
                            ]
                        })
                    ]
                })
            ]
        }]
    });
    console.log("Constructor worked with arrays");
} catch (e) {
    console.log("Failed with arrays:", e.message);
}
