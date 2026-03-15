const docx = require('docx');
const { MathRun } = docx;
const mr = new MathRun({ text: "w", bold: true });
console.log("MathRun with options:", mr);
