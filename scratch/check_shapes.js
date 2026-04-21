const pptxgen = require("pptxgenjs");
let pres = new pptxgen();
console.log("Shapes:", pres.ShapeType);
console.log("Library top level keys:", Object.keys(pptxgen));
