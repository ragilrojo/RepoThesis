const docx = require('docx');
const { MathLimit, MathLimitLower, MathRun } = docx;

console.log("MathLimit constructor properties requested by search:");
try {
    const ml = new MathLimit({ children: [] });
    console.log("MathLimit(children) ok");
} catch(e) { console.log("MathLimit(children) fail"); }

try {
    const mll = new MathLimitLower({ 
        main: [new MathRun("min")],
        limit: [new MathRun("w")]
    });
    console.log("MathLimitLower(main, limit) ok");
} catch(e) { console.log("MathLimitLower(main, limit) fail:", e.message); }

try {
    const mll = new MathLimitLower({ 
        children: [new MathRun("min"), new MathRun("w")]
    });
    console.log("MathLimitLower(children) ok");
} catch(e) { console.log("MathLimitLower(children) fail:", e.message); }
