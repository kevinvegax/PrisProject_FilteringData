const express = require('express');
const app = express();
const exceljs = require('exceljs');
const multer = require('multer')
const path = require('path');
const { exec } = require('child_process');
const zipdir = require('zip-dir');
const fs = require('fs');



app.set('view engine', 'ejs');
app.use(express.urlencoded({ extended: false }));
app.use(express.static('public'));


const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    cb(null, 'public/uploads');
  },
    filename: function (req, file, cb) {
    cb(null, 'datos.xlsx');
  },
});
  
const upload = multer({ storage });

app.get('/', async  (req, res)=> {
    try {
        const workbook = new exceljs.Workbook();
        await workbook.xlsx.readFile('public/uploads/datos.xlsx');
        const worksheet = workbook.getWorksheet(1);
        const data = [];
        worksheet.eachRow((row, rowNumer) => {
            if( rowNumer != 0){
                data.push(row.values[1])
            }
        })
        res.render('index', { data });
        console.log('Data Reived')
    } catch (error) {
        console.error('Error al leer el archivo Excel:', error);
        res.status(500).send('Error al leer el archivo Excel');
    }

});

app.get('/data', async  (req, res)=> {
    try {
        const workbook = new exceljs.Workbook();
        await workbook.xlsx.readFile('public/uploads/datos.xlsx');
        const worksheet = workbook.getWorksheet(1);
        const data = [];
        worksheet.eachRow((row) => {
                data.push(row.values[1])
        })
        res.json({ data })
        console.log(data[6])
    } catch (error) {
        console.error('Error al leer el archivo Excel:', error);
        res.status(500).send('Error al leer el archivo Excel');
    }

});


app.post('/upload', upload.single('excelFile'), (req, res) => {
    res.redirect('/');
});
  

// Define una ruta para ejecutar WebdriverIO
app.post('/execute', (req, res) => {
  // Define la ruta al directorio de tu proyecto de WebdriverIO
  const wdioProjectPath = 'C:\\Users\\ingke\\OneDrive\\Desktop\\sandBox\\wdio';

  // Define el comando de WebdriverIO que deseas ejecutar (por ejemplo, 'npm run wdio')
  const wdioCommand = 'npm run wdio';

  // Utiliza el módulo child_process para ejecutar el comando de WebdriverIO
  const wdioProcess = exec(wdioCommand, { cwd: wdioProjectPath });

  wdioProcess.stdout.on('data', (data) => {
    console.log(`Salida de WebdriverIO: ${data}`);
  });

  wdioProcess.stderr.on('data', (data) => {
    console.error(`Error de WebdriverIO: ${data}`);
  });

  wdioProcess.on('close', (code) => {
    console.log(`Script ejecutado excitosamente ${code}`);
    res.redirect('/');
  });
});

app.post('/download', (req, res) => {
  // Ruta de la carpeta que deseas comprimir y descargar
  const carpetaPath = 'C:\\Users\\ingke\\OneDrive\\Desktop\\sandBox\\wdio\\data';

  // Ruta y nombre del archivo ZIP resultante
  const archivoZip = 'catosComprimidos.zip';

  // Comprimir la carpeta en un archivo ZIP
  zipdir(carpetaPath, { saveTo: archivoZip }, (err, buffer) => {
    if (err) {
      console.error('Error al comprimir la carpeta:', err);
      res.status(500).send('Error al comprimir la carpeta');
    } else {
      // Establece las cabeceras para la descarga
      res.setHeader('Content-Disposition', `attachment; filename=${archivoZip}`);
      res.setHeader('Content-Type', 'application/zip');

      // Envía el archivo ZIP como una respuesta de descarga
      res.download(archivoZip, (err) => {
        if (err) {
          console.error('Error al descargar el archivo ZIP:', err);
          res.status(500).send('Error al descargar el archivo ZIP');
        } else {
          // Después de la descarga, elimina los archivos originales de la carpeta
          fs.readdir(carpetaPath, (err, files) => {
            if (!err) {
              files.forEach((file) => {
                const filePath = path.join(carpetaPath, file);
                fs.unlink(filePath, (err) => {
                  if (err) {
                    console.error(`Error al eliminar el archivo ${file}:`, err);
                  }
                });
              });
            }
          });
        }
      });
    }
  });
});

app.listen(211, ()=> {
    console.log('App listening on port 3000')
});