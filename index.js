const express = require("express")
const app = express()
const path = require("path")
const fileUpload = require("express-fileupload")
const Excel = require("exceljs")
const fs = require("fs")
app.use(fileUpload())

app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname + "/index.html"))
})

app.post("/upload", async function (req, res) {
  if (!req.files || Object.keys(req.files).length === 0) {
    return res.status(400).send("No files were uploaded.")
  }

  let sampleFile = req.files.sampleFile

  sampleFile.mv(path.join(__dirname, "input.xlsx"), function (err) {
    if (err) return res.status(500).send(err)
  })

  var workbook = new Excel.Workbook()
  await workbook.xlsx.readFile(path.join(__dirname, "./input.xlsx"))
  var worsheet = workbook.getWorksheet(1)
  var data = []
  for (let i = 2; i < worsheet.rowCount - 1; i++) {
    data.push({
      agno: worsheet.getCell("A" + i.toString()).value,
      mes: worsheet.getCell("B" + i.toString()).value,
      dia: worsheet.getCell("C" + i.toString()).value,
      valor: worsheet.getCell("D" + i.toString()).value,
    })
  }

  var lluvias = []

  var lluvia = {}
  var contador = 0

  data.forEach((dia) => {
    if (dia.valor.toString() != "0") {
      if (!lluvia.inicio) {
        lluvia.inicio = {
          agno: Number(dia.agno),
          mes: Number(dia.mes),
          dia: Number(dia.dia),
        }
        lluvia.pp = dia.valor
      } else {
        lluvia.pp = lluvia.pp + dia.valor
      }
      contador += 1
    } else {
      if (lluvia.inicio) {
        lluvia.duracion = contador
        lluvias.push(Object.assign({}, lluvia))
        lluvia = {}
        contador = 0
      }
    }
  })

  var clasificacion = {}
  lluvias.forEach((element) => {
    if (!clasificacion[element.duracion.toString()]) {
      clasificacion[element.duracion.toString()] = {
        cantidad: 1,
      }
    } else {
      clasificacion[element.duracion.toString()].cantidad =
        clasificacion[element.duracion.toString()].cantidad + 1
    }
  })
  console.log(clasificacion)

  const workbookout = new Excel.Workbook()
  const worksheetAF = workbookout.addWorksheet("AF")

  worksheetAF.columns = [
    { header: "Duracion", key: "duration", width: 20 },
    { header: "Cantidad", key: "quantity", width: 20 },
  ]
  Object.keys(clasificacion).forEach((duracion) => {
    worksheetAF.addRow({
      duration: duracion,
      quantity: clasificacion[duracion].cantidad,
    })
  })

  const worksheetTormentas = workbookout.addWorksheet("Tormentas")

  worksheetTormentas.columns = [
    { header: "Año", key: "agno", width: 20 },
    { header: "Mes", key: "mes", width: 20 },
    { header: "Dia", key: "dia", width: 20 },
    { header: "PP (mm)", key: "pp", width: 20 },
    { header: "Duracion", key: "duration", width: 20 },
  ]
  lluvias.forEach((storm) => {
    worksheetTormentas.addRow({
      agno: storm.inicio.agno,
      mes: storm.inicio.mes,
      dia: storm.inicio.dia,
      pp: storm.pp,
      duration: storm.duracion,
    })
  })

  await workbookout.xlsx.writeFile(path.join(__dirname, "output.xlsx"))
  res.sendFile(path.join(__dirname, "output.xlsx"), () => {
    fs.unlink(path.join(__dirname, "output.xlsx"), () => {})
  })
})

const port = process.env.PORT || 3000

app.listen(port, () => {
  console.log("server running on " + port.toString())
})
