const express = require('express');
const xl = require('exceljs');

const app = express();
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

app.use((req, res, next) => {
  res.setHeader('Access-Control-Allow-Origin', 'localhost'); // <- Habilitado localhost para desarrollo
  res.setHeader('Access-Control-Allow-Headers', 'Origin, X-Requested-With, Content-Type, Accept')
  next();
});

app.get('/download', (req, res) => {
  // choose file name to send
  const fileName = 'myExcelFile.xlsx';

  // create Workbook
  const wb = new xl.Workbook();

  // some properties
  wb.creator = 'Mauricio Contreras';
  wb.created = new Date(2019,6,15);

  // adding sheet
  const ws = wb.addWorksheet('Hoja 1');
  console.log(ws);
  ws.getRow(1).getCell(1).value = 100;
  ws.getRow(2).getCell(1).value = 200;
  ws.getRow(3).getCell(1).value = ws.getRow(1).getCell(1).value + ws.getRow(2).getCell(1).value;

  // writing Workbook as Buffer
  wb.xlsx.writeBuffer()
    .then((buffer) => {
      console.log('Sending buffer');
      res.set({
        'Content-Type': 'application/octet-stream',
        'Content-Disposition': 'attachment; filename="' + fileName + '"',
        'X-Processed-FileName': fileName
      });

      return res.status(200).send(buffer);
    })
    .catch((err) => {
      console.log(err.message);
      return res.status(200).json({error: err.message})
    });
});

app.get('/', (req, res) => {
  let template = `
  <!DOCTYPE html>
  <html lang="en">
  <head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta http-equiv="X-UA-Compatible" content="ie=edge">
    <title>Get File</title>
  </head>
  <body>
    <button type="button" id="download">Descargar</button>
    <div id="headers"></div>
    <script src="https://unpkg.com/axios/dist/axios.min.js"></script>
    <script>
        const button = document.getElementById('download');
        button.addEventListener('click', (event) => {
          event.preventDefault();
          axios.get('http://localhost:3000/download')
          .then((res) => {
            const fileName = res.headers['x-processed-filename'];
            const url = window.URL.createObjectURL(new Blob([res]));
            const link = document.createElement('a');
            link.href = url;
            link.setAttribute('download', fileName);
            document.body.appendChild(link);
            link.click();
            link.remove();
          });
        });
    </script>
  </body>
  </html>
  `
  res.status(200).send(template);
});

app.listen(3000, () => {
  console.log('App started on port 3000');
})
