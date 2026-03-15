const docx = require('docx');
const { Document, Math, MathRun, MathLimit, Packer } = docx;

try {
    const doc = new Document({
        sections: [{
            children: [
                new docx.Paragraph({
                    children: [
                        new Math({
                            children: [
                                new MathLimit({
                                    main: [new MathRun("min")],
                                    limit: [new MathRun("w")],
                                })
                            ]
                        })
                    ]
                })
            ]
        }]
    });
    console.log("MathLimit worked");
} catch (e) {
    console.log("MathLimit failed:", e.message);
}
