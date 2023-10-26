const express = require('express');
const { PDFDocument } = require('pdf-lib');
const fileUpload = require('express-fileupload');
const ExcelJS = require('exceljs');
const archiver = require('archiver');
const stream = require('stream');
const Buffer = require('buffer').Buffer;

const interface = require("/index.html")

const port = process.env.PORT ?? 3333

const app = express();
app.use(fileUpload());

app.use("/form", interface);

app.get('/', (req, res) => {
  res.sendFile(__dirname + '/index.html');
});

app.post('/split', async (req, res) => {
  try {
    if (!req.files || !req.files.pdf || !req.files.excel) {
      return res.status(400).send('Nenhum arquivo enviado.');
    }

    const pdfBuffer = req.files.pdf.data;
    const pdfDoc = await PDFDocument.load(pdfBuffer);

    const excelBuffer = req.files.excel.data;
    const excelWorkbook = new ExcelJS.Workbook();
    await excelWorkbook.xlsx.load(excelBuffer);

    const worksheet = excelWorkbook.getWorksheet(1);

    const archive = archiver('zip', { zlib: { level: 9 } });
    const output = new stream.PassThrough();

    archive.pipe(output);

    for (let i = 0; i < pdfDoc.getPageCount(); i++) {
      const page = pdfDoc.getPages()[i];
      const newPdfDoc = await PDFDocument.create();
      const [newPage] = await newPdfDoc.copyPages(pdfDoc, [i]);

      newPdfDoc.addPage(newPage);

      const newPdfBytes = await newPdfDoc.save();

      if (i < worksheet.rowCount) {
        const filename = worksheet.getCell(i + 1, 2).value + '.pdf';
        // Converta os bytes em um buffer
        const pdfBuffer = Buffer.from(newPdfBytes);
        archive.append(pdfBuffer, { name: filename });
      } else {
        console.log(`Não há nome definido para a página ${i + 1}.`);
      }
    }

    archive.finalize();

    res.set('Content-Disposition', 'attachment; filename=pdfs.zip');
    res.set('Content-Type', 'application/zip');
    output.pipe(res);
  } catch (error) {
    console.error(error);
    res.status(500).send('Erro ao dividir e criar o arquivo ZIP.');
  }
});
app.listen(port, '0.0.0.0', () => {
  console.log(`HTTP Server running ${port}`);
});
// app.listen(3000, () => {
//   console.log('Servidor rodando na porta 3000');
// });