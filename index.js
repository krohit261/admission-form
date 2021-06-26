const express = require('express')
const fs = require('fs')
const app = express()
const {spawn} = require('child_process')
const port = 3000
var toPdf = require("office-to-pdf")

app.get('/', (req, res) => {
  res.sendFile(__dirname + '/index.html')
})

app.get('/upload' , (req, res) => {
  console.log(JSON.stringify(req.query));
  const python = spawn('python', ['auto.py' , `${JSON.stringify(req.query)}`]);
  python.stdout.on('data', function (data) {
    console.log(`${data}`);
   });
   python.stderr.on('data', (data) => {
    console.error(`child stderr:\n${data}`);
    res.send(`${data}`)
  });
  python.on('close',async (code, signal) => {
    console.log(`child process close all stdio with code code ${code} and signal ${signal}`);
    var wordBuffer = fs.readFileSync(__dirname + "/Admission_form.pptx")
    // send data to browser
    toPdf(wordBuffer).then(
      (pdfBuffer) => {
        console.log(pdfBuffer)
        console.log('sending pdf');
        res.contentType("application/pdf");
        res.send(pdfBuffer);
      }, (err) => {
        console.log(err)
        res.send(`${err}`)
      }
    )
  });
  
})

app.get('/excel' , (req , res) => {
  res.sendFile(__dirname + '/data.xlsx' , (err) => {if(err) console.log(err)});
})

app.listen(port, () => {
  console.log(`app listening at http://localhost:${port}`)
})