import express from 'express';
import bodyParser from 'body-parser';
import fileUpload from 'express-fileupload';
import fs from 'fs';
import { PythonShell } from 'python-shell';


const app = express();


app.use(fileUpload()); // Don't forget this line!

let arList: Array<string> = [];
let ledgerList: Array<string>= [];

const deleteFiles = () => {
   console.log("deleting files?");
   arList.forEach((file) => {
      try {
         fs.unlinkSync(file);
         fs.unlinkSync(`${file.split('.')[0]}_fixed.xlsx`);
      } catch(error) {
         console.log('error deleting files', error);
      }
   });
   ledgerList.forEach((file) => {
      try {
         fs.unlinkSync(file);
         fs.unlinkSync(`${file.split('.')[0]}fixed.xlsx`);
      } catch(error) {
         console.log('error deleting files', error);
      }
   });
   fs.unlinkSync(`${__dirname}/../formatted.xlsx`);
   arList = [];
   ledgerList = [];
}

app.get('/resetfiles', (req, res) => {
   deleteFiles();
})


app.post('/uploadar', (req:any, res:any) => {
   console.log('uploading ar file');
   if (req.files) {
      fs.writeFile(`${__dirname}/pythonScripts/${req.files.file.name}`, req.files.file.data, (err) => {
         if(err) {
            res.status(400);
            res.send(`Error ${err.message}`)
         } else {
            arList.push(`${__dirname}/pythonScripts/${req.files.file.name}`);
            res.send("successful uploaded");
         }
      })
   }
});

app.post('/uploadledger', (req:any, res:any) => {
   console.log('uploading ledger file');
   if (req.files) {
      fs.writeFile(`${__dirname}/pythonScripts/${req.files.file.name}`, req.files.file.data, (err) => {
         if(err) {
            res.status(400);
            res.send(`Error ${err.message}`)
         } else {
            ledgerList.push(`${__dirname}/pythonScripts/${req.files.file.name}`);
            res.send("successful uploaded");
         }
      })
   }
});

app.get('/format', (req:any, res: any) => {
   console.log('formating documents', arList, ledgerList);
   const options: any = {
      mode: 'text',
      pythonPath: `${__dirname}/pythonScripts/bin/python3`,
      scriptPath: `${__dirname}/pythonScripts/`,
      args: [...arList, 'ardone', ...ledgerList],
   }
   PythonShell.run('myExcelEditor.py', options, (error, results) => {
      if (error) return console.log('error pyshell', error);
      console.log('format successful', results);
      console.log('success\n', results);
      res.setHeader('Content-Type', 'application/vnd.openxmlformats');
      res.download(`${__dirname}/../formatted.xlsx`, (err:Error) => {
         if(err) {
            res.status(400).send("error sending formatted file")
            console.log('failure in pythong shell= ', error);
         } 
         
         console.log('successfuly sent formatted file');
         deleteFiles();
      });
   });

});




app.use(express.static(`${__dirname}/../client/dist`));

export default app;