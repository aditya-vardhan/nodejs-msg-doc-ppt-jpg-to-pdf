import express from 'express'
import {promises as fs} from 'fs';
import libre from 'libreoffice-convert';
import path from 'path';
import { jsPDF } from 'jspdf';
import MsgReader from '@kenjiuno/msgreader';

const app = express()
const port = 3000;

// libre.convertAsync = import('util').promisify(libre.convert);
(async () => {
    const utilD = await import('util')
    libre.convertAsync = utilD.promisify(libre.convert)
  })();

app.get('/', async (req, res) => {
    console.log('received request', req.query)
    const fileData = await getFile(req.query.file);
    // res.setHeader('Content-Length', stat.size);
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `attachment; filename=HAQAuditTrainReport_${new Date().toISOString()}.pdf`);
    res.write(fileData, 'binary');
    res.end();
    // res.status(200).send()
})

app.listen(port, () => {
    console.log(`Example app listening on port ${port}`)
})

const getFile = async (filename) => {
    const fileExtension = filename.split('.').pop()
    switch(fileExtension){
        case 'pdf':
            return handlePdf(filename);
        case 'doc':
        case 'docx':
        case 'ppt':
        case 'pptx':
            return handleDoc(filename);
        case 'msg':
            return handleMsg(filename);
    }
}

const handlePdf = async (filename) => {
    const fileData = await fs.readFile(filename);
    return fileData;
}

const handleDoc = async (filename) => {
    const ext = '.pdf'
    // Read file
    const docxBuf = await fs.readFile(filename);

    // Convert it to pdf format with undefined filter (see Libreoffice docs about filter)
    let pdfBuf = await libre.convertAsync(docxBuf, ext, undefined);
    
    // Here in done you have pdf file which you can save or transfer in another stream
    return pdfBuf;
}

const handleMsg = async (filename) => {
    const docxBuf = await fs.readFile(filename);
    const msgReader = new MsgReader.default(docxBuf);
    const fileData = msgReader.getFileData();
    //FileData will have body, subject etc
    const subject = fileData.subject;
    const getBody = fileData.body;
    const doc = new jsPDF();
    doc.text(`${subject}`, 14, 30);
    doc.text(`${getBody}`, 14, 45);
    return doc.output()
}
