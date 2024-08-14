const express = require('express')
const serverless = require('serverless-http')
const AWS = require('aws-sdk')
const xlsx = require('xlsx')
const ExcelJS = require('exceljs')

const isAWSLambda = !!process.env.AWS_LAMBDA_FUNCTION_NAME
if (!isAWSLambda) {
  require('dotenv').config()
}
const AUTH_TOKEN = process.env.AUTH_TOKEN

const app = express()
app.use(express.json({ limit: '50mb' }))

const router = express.Router()
router.use((req, res, next) => {
  const token = req.get('x-auth-token')
  if (!!token && token === AUTH_TOKEN) {
    next()
  } else {
    res.status(401).json({ message: 'Invalid auth token' })
  }
})

router.post('/lambda/json-to-excel/styled', async (req, res) => {
  console.log('styled working')
  const jsonData = req.body.excel
  const defaultStyle = req.body.config.default_style
  const excelData = await convertJsonToStyledExcel(jsonData, defaultStyle)
  const url = await uploadToAWS(req.body.config, excelData)
  return res.json({ url })
})

router.post('/styled', async (req, res) => {
  console.log('styled not working form infra side')
  return res.json({ message: "jsonToStyledExcel" })
})

router.post('/api/jsonToExcel', async (req, res) => {
  console.log('old path')
  const jsonData = req.body.excel
  const excelData = await convertJsonToExcel(jsonData)
  const link = await uploadToAWS(req.body.config, excelData)

  return res.json(link)
})

router.post('*', async (req, res) => {
  console.log('default path')
  return res.json({ message: "default path" })
})

const uploadToAWS = async (config, excelData) => {
  const { s3FilePublic, s3Region, s3Bucket, s3KeyId, s3SecretKey, s3Path } = config

  AWS.config.update({ accessKeyId: s3KeyId, secretAccessKey: s3SecretKey, region: s3Region, signatureVersion: 'v4' })
  const s3 = new AWS.S3()
  const dataset = {
    Bucket: s3Bucket,
    Key: s3Path,
    Body: excelData,
    ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    ACL: !!s3FilePublic ? 'public-read' : 'private',
  }
  const response = await s3.upload(dataset).promise()
  return response.Location
}

const convertJsonToExcel = (jsonData) => {
  const workbook = xlsx.utils.book_new()

  Object.keys(jsonData).forEach((sheetName) => {
    const worksheet = xlsx.utils.json_to_sheet(jsonData[sheetName])
    xlsx.utils.book_append_sheet(workbook, worksheet, sheetName)
  })

  return xlsx.write(workbook, { type: 'buffer', bookType: 'xlsx' })
}

async function convertJsonToStyledExcel(jsonData, defaultStyle) {
  const workbook = new ExcelJS.Workbook()
  Object.keys(jsonData).forEach((sheetName) => {
    const worksheet = workbook.addWorksheet(sheetName)
    Object.entries(jsonData[sheetName].headers).forEach(([headerText, style]) => {
      const headerRow = worksheet.addRow([headerText])
      const headerStyle = style
      headerRow.getCell(1).style = headerStyle
    })

    // Get column headers from the first row of data
    const columnHeaders = Object.values(jsonData[sheetName].keysValue).map((keyValueObj) => keyValueObj.value)
    const headerRow = worksheet.addRow(columnHeaders)

    // Apply styles from "keysValue" to the header row
    Object.values(jsonData[sheetName].keysValue).forEach((keyValueObj, index) => {
      const cell = headerRow.getCell(index + 1)
      if (keyValueObj.style) {
        cell.style = keyValueObj.style
      }
    })

    const headers = Object.keys(jsonData[sheetName].data[0])
    // Iterate over rows from the JSON data
    jsonData[sheetName].data.forEach((row, rowIndex) => {
      const excelRow = worksheet.addRow(headers.map((key) => row[key].value))

      // Apply styles if they exist
      headers.forEach((key, colIndex) => {
        const cell = excelRow.getCell(colIndex + 1)
        if (row[key].style) {
          cell.style = {
            ...row[key].style,
          }
        }
      })
    })
    worksheet.columns.forEach((column) => {
      column.width = defaultStyle.width || 10
    })
  })

  const buffer = await workbook.xlsx.writeBuffer()
  return buffer
}

app.use('/', router)

const startServer = async () => {
  app.listen(3000, () => {
    console.log('listening on port 3000!')
  })
}
if (!isAWSLambda) {
  startServer()
}

//lambda handling
const handler = serverless(app)

exports.handler = async (event, context, callback) => {
  // console.log("entryyyyyy---->", event)
  const response = await handler(event, context, callback)
  // console.log('request closed as response', response)
  return response
}
