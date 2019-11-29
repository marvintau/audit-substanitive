const fs = require('fs');
const unzipper = require('unzipper');
const path = require('path');

const DOCX_DIR = './sheets/docx';

let list = fs
    .readdirSync('./sheets/docx')
    .filter(e => path.extname(e) === '.docx')
    .map(e => path.resolve(__dirname, DOCX_DIR, e));

for (let filePath of list) {

    let {dir, name} = path.parse(filePath),
        extractPath = path.resolve(dir, `${name}.xml`);

    console.log(dir, name);

    fs.createReadStream(filePath)
    .pipe(unzipper.Parse())
    .on('entry', entry => {
        if(entry.path == 'word/document.xml'){
            entry.pipe(fs.createWriteStream(extractPath));
        } else {
            entry.autodrain();
        }
    })
}