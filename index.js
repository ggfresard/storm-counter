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
  var añoInicio = data[0].agno
  var añoFinal = data[data.length - 1].agno

  var lluvias = []
  var tormentas2 = {}
  var tormentas3 = {}
  var tormentas4 = {}
  var tormentas1 = {}

  for (let año = añoInicio; año < añoFinal + 1; año++) {
    var dias = data.filter((dia) => dia.agno == año)
    for (let i = 0; i < dias.length; i++) {
      if (dias[i].valor > 0) {
        if (!tormentas1[año.toString()]) {
          tormentas1[año.toString()] = [
            {
              inicio: {
                agno: dias[i].agno,
                mes: dias[i].mes,
                dia: dias[i].dia,
              },
              valor: dias[i].valor,
            },
          ]
        } else {
          tormentas1[año.toString()].push({
            inicio: {
              agno: dias[i].agno,
              mes: dias[i].mes,
              dia: dias[i].dia,
            },
            valor: dias[i].valor,
          })
        }
      }
    }
    for (let i = 0; i < dias.length - 1; i++) {
      var valor = 0
      for (let j = 0; j < 2; j++) {
        valor += dias[i + j].valor
      }
      if (valor > 0) {
        if (!tormentas2[año.toString()]) {
          tormentas2[año.toString()] = [
            {
              inicio: {
                agno: dias[i].agno,
                mes: dias[i].mes,
                dia: dias[i].dia,
              },
              valor: valor,
            },
          ]
        } else {
          tormentas2[año.toString()].push({
            inicio: { agno: dias[i].agno, mes: dias[i].mes, dia: dias[i].dia },
            valor: valor,
          })
        }
      }
      if (i < dias.length - 2) {
        valor = 0
        for (let j = 0; j < 3; j++) {
          valor += dias[i + j].valor
        }
        if (valor > 0) {
          if (!tormentas3[año.toString()]) {
            tormentas3[año.toString()] = [
              {
                inicio: {
                  agno: dias[i].agno,
                  mes: dias[i].mes,
                  dia: dias[i].dia,
                },
                valor: valor,
              },
            ]
          } else {
            tormentas3[año.toString()].push({
              inicio: {
                agno: dias[i].agno,
                mes: dias[i].mes,
                dia: dias[i].dia,
              },
              valor: valor,
            })
          }
        }
      }
      if (i < dias.length - 3) {
        valor = 0
        for (let j = 0; j < 4; j++) {
          valor += dias[i + j].valor
        }
        if (valor > 0) {
          if (!tormentas4[año.toString()]) {
            tormentas4[año.toString()] = [
              {
                inicio: {
                  agno: dias[i].agno,
                  mes: dias[i].mes,
                  dia: dias[i].dia,
                },
                valor: valor,
              },
            ]
          } else {
            tormentas4[año.toString()].push({
              inicio: {
                agno: dias[i].agno,
                mes: dias[i].mes,
                dia: dias[i].dia,
              },
              valor: valor,
            })
          }
        }
      }
    }
  }

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
  var maxAnual = {}
  lluvias.forEach((element) => {
    if (!maxAnual[element.inicio.agno]) {
      maxAnual[element.inicio.agno] = element
    } else {
      if (maxAnual[element.inicio.agno].pp < element.pp) {
        maxAnual[element.inicio.agno] = element
      }
    }
    if (!clasificacion[element.duracion.toString()]) {
      clasificacion[element.duracion.toString()] = {
        cantidad: 1,
      }
    } else {
      clasificacion[element.duracion.toString()].cantidad =
        clasificacion[element.duracion.toString()].cantidad + 1
    }
  })

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
  const worksheetMaxAnual = workbookout.addWorksheet("Maxima tormenta anual")

  worksheetMaxAnual.columns = [
    { header: "Año", key: "agno", width: 20 },
    { header: "Mes", key: "mes", width: 20 },
    { header: "Dia", key: "dia", width: 20 },
    { header: "PP (mm)", key: "pp", width: 20 },
    { header: "Duracion", key: "duration", width: 20 },
  ]
  Object.keys(maxAnual).forEach((year) => {
    var storm = maxAnual[year]
    worksheetMaxAnual.addRow({
      agno: storm.inicio.agno,
      mes: storm.inicio.mes,
      dia: storm.inicio.dia,
      pp: storm.pp,
      duration: storm.duracion,
    })
  })

  const worksheetTormentas1 = workbookout.addWorksheet("Tormentas 1 dia")

  worksheetTormentas1.columns = [
    { header: "Año", key: "agno", width: 20 },
    { header: "Mes", key: "mes", width: 20 },
    { header: "Dia", key: "dia", width: 20 },
    { header: "PP (mm)", key: "pp", width: 20 },
  ]
  Object.keys(tormentas1).forEach((year) => {
    var storms = tormentas1[year]
    var max = storms[0]
    for (var i = 1; i < storms.length; i++) {
      if (max.valor < storms[i].valor) {
        max = storms[i]
      }
    }

    worksheetTormentas1.addRow({
      agno: max.inicio.agno,
      mes: max.inicio.mes,
      dia: max.inicio.dia,
      pp: max.valor,
    })
  })

  const worksheetTormentas2 = workbookout.addWorksheet("Tormentas 2 dias")

  worksheetTormentas2.columns = [
    { header: "Año", key: "agno", width: 20 },
    { header: "Mes", key: "mes", width: 20 },
    { header: "Dia", key: "dia", width: 20 },
    { header: "PP (mm)", key: "pp", width: 20 },
  ]
  Object.keys(tormentas2).forEach((year) => {
    var storms = tormentas2[year]
    var max = storms[0]
    for (var i = 1; i < storms.length; i++) {
      if (max.valor < storms[i].valor) {
        max = storms[i]
      }
    }

    worksheetTormentas2.addRow({
      agno: max.inicio.agno,
      mes: max.inicio.mes,
      dia: max.inicio.dia,
      pp: max.valor,
    })
  })
  const worksheetTormentas3 = workbookout.addWorksheet("Tormentas 3 dias")

  worksheetTormentas3.columns = [
    { header: "Año", key: "agno", width: 20 },
    { header: "Mes", key: "mes", width: 20 },
    { header: "Dia", key: "dia", width: 20 },
    { header: "PP (mm)", key: "pp", width: 20 },
  ]
  Object.keys(tormentas3).forEach((year) => {
    var storms = tormentas3[year]
    var max = storms[0]
    for (var i = 1; i < storms.length; i++) {
      if (max.valor < storms[i].valor) {
        max = storms[i]
      }
    }

    worksheetTormentas3.addRow({
      agno: max.inicio.agno,
      mes: max.inicio.mes,
      dia: max.inicio.dia,
      pp: max.valor,
    })
  })
  const worksheetTormentas4 = workbookout.addWorksheet("Tormentas 4 dias")

  worksheetTormentas4.columns = [
    { header: "Año", key: "agno", width: 20 },
    { header: "Mes", key: "mes", width: 20 },
    { header: "Dia", key: "dia", width: 20 },
    { header: "PP (mm)", key: "pp", width: 20 },
  ]
  Object.keys(tormentas4).forEach((year) => {
    var storms = tormentas4[year]
    var max = storms[0]
    for (var i = 1; i < storms.length; i++) {
      if (max.valor < storms[i].valor) {
        max = storms[i]
      }
    }

    worksheetTormentas4.addRow({
      agno: max.inicio.agno,
      mes: max.inicio.mes,
      dia: max.inicio.dia,
      pp: max.valor,
    })
  })
  const worksheetOrdenadas = workbookout.addWorksheet("Ordenadas para informe")

  worksheetOrdenadas.columns = [
    { header: "Año", key: "agno", width: 20 },
    { header: "Mes", key: "mes", width: 20 },
    { header: "Dia", key: "dia", width: 20 },
    { header: "PP (mm)", key: "pp", width: 20 },
    { header: "Año", key: "agno1", width: 20 },
    { header: "Mes", key: "mes1", width: 20 },
    { header: "Dia", key: "dia1", width: 20 },
    { header: "PP (mm)", key: "pp1", width: 20 },
    { header: "Año", key: "agno2", width: 20 },
    { header: "Mes", key: "mes2", width: 20 },
    { header: "Dia", key: "dia2", width: 20 },
    { header: "PP (mm)", key: "pp2", width: 20 },
  ]
  var grupos = (data.length - (data.length % 90)) / 90
  for (let i = 0; i < grupos; i++) {
    for (let j = 0; j < 30; j++) {
      worksheetOrdenadas.addRow({
        agno: data[i * 90 + j].agno,
        mes: data[i * 90 + j].mes,
        dia: data[i * 90 + j].dia,
        pp: data[i * 90 + j].valor,
        agno1: data[i * 90 + j + 30].agno,
        mes1: data[i * 90 + j + 30].mes,
        dia1: data[i * 90 + j + 30].dia,
        pp1: data[i * 90 + j + 30].valor,
        agno2: data[i * 90 + j + 60].agno,
        mes2: data[i * 90 + j + 60].mes,
        dia2: data[i * 90 + j + 60].dia,
        pp2: data[i * 90 + j + 60].valor,
      })
    }
    worksheetOrdenadas.addRow({})
  }
  console.log("ay")
  var resto = data.length % 90
  if (resto > 60) {
    var dif = resto - 60
    let contador = 0
    for (let j = 0; j < 30; j++) {
      if (contador < dif) {
        worksheetOrdenadas.addRow({
          agno: data[grupos * 90 + j].agno,
          mes: data[grupos * 90 + j].mes,
          dia: data[grupos * 90 + j].dia,
          pp: data[grupos * 90 + j].valor,
          agno1: data[grupos * 90 + j + 30].agno,
          mes1: data[grupos * 90 + j + 30].mes,
          dia1: data[grupos * 90 + j + 30].dia,
          pp1: data[grupos * 90 + j + 30].valor,
          agno2: data[grupos * 90 + j + 60].agno,
          mes2: data[grupos * 90 + j + 60].mes,
          dia2: data[grupos * 90 + j + 60].dia,
          pp2: data[grupos * 90 + j + 60].valor,
        })
      } else {
        worksheetOrdenadas.addRow({
          agno: data[grupos * 90 + j].agno,
          mes: data[grupos * 90 + j].mes,
          dia: data[grupos * 90 + j].dia,
          pp: data[grupos * 90 + j].valor,
          agno1: data[grupos * 90 + j + 30].agno,
          mes1: data[grupos * 90 + j + 30].mes,
          dia1: data[grupos * 90 + j + 30].dia,
          pp1: data[grupos * 90 + j + 30].valor,
        })
      }
      contador += 1
    }
  } else if (resto > 30) {
    var dif = resto - 30
    let contador = 0
    for (let j = 0; j < 30; j++) {
      if (contador < dif) {
        worksheetOrdenadas.addRow({
          agno: data[grupos * 90 + j].agno,
          mes: data[grupos * 90 + j].mes,
          dia: data[grupos * 90 + j].dia,
          pp: data[grupos * 90 + j].valor,
          agno1: data[grupos * 90 + j + 30].agno,
          mes1: data[grupos * 90 + j + 30].mes,
          dia1: data[grupos * 90 + j + 30].dia,
          pp1: data[grupos * 90 + j + 30].valor,
        })
      } else {
        worksheetOrdenadas.addRow({
          agno: data[grupos * 90 + j].agno,
          mes: data[grupos * 90 + j].mes,
          dia: data[grupos * 90 + j].dia,
          pp: data[grupos * 90 + j].valor,
        })
      }
      contador += 1
    }
  } else if (resto > 0) {
    var dif = resto - 30
    for (let j = 0; j < resto; j++) {
      worksheetOrdenadas.addRow({
        agno: data[grupos * 90 + j].agno,
        mes: data[grupos * 90 + j].mes,
        dia: data[grupos * 90 + j].dia,
        pp: data[grupos * 90 + j].valor,
      })
    }
  }

  await workbookout.xlsx.writeFile(path.join(__dirname, "output.xlsx"))
  res.sendFile(path.join(__dirname, "output.xlsx"), () => {
    fs.unlink(path.join(__dirname, "output.xlsx"), () => {})
  })
})

const port = process.env.PORT || 3000

app.listen(port, () => {
  console.log("server running on " + port.toString())
})
