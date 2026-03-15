const docx = require('docx');
const { MathSum, MathRun } = docx;

try {
    new MathSum({
        children: [new MathRun("x")]
    });
    console.log("MathSum worked with children array");
} catch (e) {
    console.log("MathSum failed with children array:", e.message);
}
