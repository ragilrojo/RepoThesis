const fs = require('fs');
const path = require('path');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const ImageModule = require('docxtemplater-image-module-free');

async function debugDocx() {
    const inputPath = path.resolve(__dirname, 'proposal_tesis_ragil.docx');
    if (!fs.existsSync(inputPath)) {
        console.error("File not found");
        return;
    }

    const content = fs.readFileSync(inputPath, 'binary');
    const zip = new PizZip(content);

    const imageOpts = {
        centered: true,
        getImage: function(tagValue, tagName) {
            console.log(`[DEBUG] getImage for tag: ${tagName}, value type: ${typeof tagValue}`);
            if (!tagValue) {
                console.error(`[DEBUG] No value for tag: ${tagName}`);
                return Buffer.from('');
            }
            if (Buffer.isBuffer(tagValue)) return tagValue;
            try {
                return fs.readFileSync(tagValue);
            } catch (e) {
                console.error(`[DEBUG] Error reading file for tag ${tagName}: ${e.message}`);
                return Buffer.from('');
            }
        },
        getSize: function(img, tagValue, tagName) {
            console.log(`[DEBUG] getSize for tag: ${tagName}`);
            if (tagName.includes("LOGO")) return [200, 200];
            if (tagName.includes("FRAMEWORK")) return [550, 350];
            if (tagName.includes("SYM_")) return [25, 25];
            if (tagName.includes("RUMUS")) return [400, 80];
            return [300, 300];
        }
    };

    const doc = new Docxtemplater(zip, {
        modules: [new ImageModule(imageOpts)],
        delimiters: { start: '[[', end: ']]' }
    });

    const renderData = {
        "LOGO_UNM": path.resolve(__dirname, 'logo_unm.png'),
        "IMAGE_FRAMEWORK": path.resolve(__dirname, 'framwrok.jpg'),
        "RUMUS_NETWORK_MARKOWITZ": Buffer.from('fake buffer'),
        "SYM_W": Buffer.from('fake buffer')
    };

    try {
        console.log("[DEBUG] Starting render...");
        doc.render(renderData);
        console.log("[DEBUG] Render successful");
    } catch (error) {
        console.error("[DEBUG] Render failed:", error.message);
        if (error.properties && error.properties.errors) {
            console.error(JSON.stringify(error.properties.errors, null, 2));
        }
    }
}

debugDocx();
