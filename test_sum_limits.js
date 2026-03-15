const docx = require('docx');
const { Math, MathRun, MathSum, MathSubScript, MathSuperScript, Document } = docx;

console.log("Testing MathSum properties...");
const sum = new MathSum({
    subScript: [new MathRun("i=1")],
    superScript: [new MathRun("n")],
    children: [new MathRun("x")]
});
console.log("MathSum constructed with sub/super properties");

const sum2 = new MathSum({
    children: [
        new MathRun("\u2211"), 
        new MathRun("x")
    ]
});

// Let's check if MathSubSuperScript is better for the sum symbol
// Or if MathSum has specific property names
console.log("Names in sum object:", Object.keys(sum));
