const fs = require('fs');
const pdf = require('pdf-parse');

let dataBuffer = fs.readFileSync('e:\\\\ProjectNodeJs\\\\temp_doc_build\\\\PenelitianSebelumnya\\\\Pedoman Penyusunan Tesis Revisi 2 - Februari 2024.pdf');

pdf(dataBuffer).then(function(data) {
    console.log(data.text);
}).catch(err => {
    console.error(err);
});
