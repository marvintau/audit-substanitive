const fs = require('fs').promises;
const path = require('path');

const xml = require("xml2js");
const {stripPrefix} = xml.processors;

const parser = new xml.Parser({
    explicitChildren: true,
    preserveChildrenOrder: true,
    tagNameProcessors: [stripPrefix]
});

const XML_PATH = './sheets/docx';


const keypress = async () => {
    process.stdin.setRawMode(true)
    return new Promise(resolve => process.stdin.once('data', (input) => {
        let key = ([...input.values()])[0];
        process.stdin.setRawMode(false)
        if(key === 3){
            throw Error('exit');
        }
        resolve()
    }))
}

// expecting elem contains an 'p' property

function handlePrevSpace(e){
    return (typeof e === 'object' && e.$ !== undefined) ? ' ' : e;
}

function getParagraphStyle(elem) {
    if(elem['pPr'] && elem['pPr'].length > 0 && elem['pPr'][0]['pStyle'] && elem['pPr'][0]['pStyle'].length > 0){
        return (elem['pPr'][0]['pStyle'][0].$['w:val']);
    } else {
        return 'Normal'
    }
}

function handleParagraph(elem){

    let style = getParagraphStyle(elem);

    let content;
    if (elem['r'] !== undefined){
        content = elem['r']
            .map(e => e['t'] !== undefined ? e['t'] : '').flat()
            .map(handlePrevSpace)
            .join('');
    } else {
        content = '';
    }
    return {style, content};
}

function getGridSpan(elem){
    if('gridSpan' in elem['tcPr'][0]){
        return elem['tcPr'][0]['gridSpan'][0].$['val']
    } else {
        return 1;
    }
}

function getTableGrid(elem){
    let rows = [];

    for (let row of elem['tr']){
        let marshalledRow = row['tc'].map(e => {
            return {
                span: getGridSpan(e),
                cell: e['p'].map(e => handleParagraph(e)).join('\n')
            };
        });
        rows.push(marshalledRow);
    }
    return rows;
}

async function getDocTable(filePath){

    let parsed = [];

    let content = await fs.readFile(filePath);
    let result  = await parser.parseStringPromise(content);

    let docBody = result['document']['body'];

    for(let {$$} of docBody){
        for (let elem of $$){

            if(elem['#name'] === 'p') {
                parsed.push(handleParagraph(elem))
            }

            if (elem['#name'] === 'tbl'){
                parsed.push(getTableGrid(elem));
            }
        }
    }

    return parsed;
}

(async () => {

    let fileList;
    try{
        fileList = await fs.readdir(XML_PATH)
    } catch(e) {
        throw Error('reading dir error, stopped.')
    }

    fileList = fileList.filter(e => path.extname(e) === '.xml')
    .map(e => path.resolve(__dirname, XML_PATH, e));


    for(let filePath of fileList){
        console.log(await getDocTable(filePath));
    }
})()