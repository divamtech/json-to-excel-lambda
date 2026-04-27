const express = require('express')
const serverless = require('serverless-http')
const AWS = require('aws-sdk')
const xlsx = require('xlsx')
const ExcelJS = require('exceljs')
const { PassThrough } = require('stream')

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

router.post('/lambda/json-to-excel/from-link', async (req, res) => {
  console.log('from link working')
  try {
    const { jsonUrl, type } = req.body
    let jsonData = await fetch(jsonUrl)
    jsonData = await jsonData.json()
    const defaultStyle = jsonData.config.default_style
    const excelFunc = getExcelFunc(type)
    const excelData = await excelFunc(jsonData.excel, defaultStyle)
    const url = await uploadToAWS(jsonData.config, excelData)
    return res.json({ url })  
  } catch (error) {
    console.log('error', error)
    res.status(400).json({ message: 'error in your request payload', error: error.message, rawError: error })
  }
})

router.post('/lambda/stream/json-to-excel/from-link', async (req, res) => {
  try {
    const { jsonUrl, type } = req.body
    let jsonData = await fetch(jsonUrl)
    jsonData = await jsonData.json()
    const excelFunc = getExcelFunc(type)
    const passThrough = new PassThrough()
    // start upload
    console.log('Process start')
    const startTime = Date.now()
    const uploadPromise = uploadToAWS(jsonData.config, passThrough, true)
    const streamOptions = { isStream: true, stream: passThrough, defaultStyle: jsonData.config?.default_style }
    console.log('Function start processing')
    await excelFunc(jsonData.excel, streamOptions)
    const endTime = Date.now()
    console.log('::::::::Time taken convert to excel -', (endTime - startTime)/1000,' seconds')
    const response = await uploadPromise
    return res.json({ url: response.Location })
  } catch (error) {
    console.log('error', error)
    res.status(400).json({
      message: 'error in your request payload',
      error: error.message,
    })
  }
})

router.post('/lambda/json-to-excel/styled', async (req, res) => {
  console.log('styled working')
  try {
    const jsonData = req.body.excel
    const defaultStyle = req.body.config.default_style
    const excelData = await convertJsonToStyledExcel(jsonData, defaultStyle)
    const url = await uploadToAWS(req.body.config, excelData)
    return res.json({ url })  
  } catch (error) {
    console.log('error', error)
    res.status(400).json({ message: 'error in your request payload', error: error.message, rawError: error })
  }
})

router.post('/lambda/json-to-excel/common-styled', async (req, res) => {
  console.log('styled working')
  try {
    const jsonData = req.body.excel
    const excelData = await convertJsonToCommonStyledExcel(jsonData)
    const url = await uploadToAWS(req.body.config, excelData)
    return res.json({ url })  
  } catch (error) {
    console.log('error', error)
    res.status(400).json({ message: 'error in your request payload', error: error.message, rawError: error })
  }
})

router.post('/lambda/json-to-excel/client-styled', async (req, res) => {
  console.log('client styled working')
  try {
    const jsonData = req.body.excel
    const excelData = await convertJsonToStyledExcel(jsonData)
    let finalBuffer = excelData
    if (jsonData.Lookups.data?.length) {
      const wb=new ExcelJS.Workbook()
       await wb.xlsx.load(excelData)
       const config=req.body.lookupConfig || {}
      await injectFormulasIntoSheet(wb, jsonData.Lookups.data,config)
     finalBuffer = await wb.xlsx.writeBuffer()
     } else {
      console.log('No countries found, skipping injection')
     }
    const url = await uploadToAWS(req.body.config, finalBuffer)
    return res.json({ url })  
  } catch (error) {
    console.log('error', error)
    res.status(400).json({ message: 'error in your request payload', error: error.message, rawError: error })
  }
})

router.post('/api/jsonToExcel', async (req, res) => {
  try {
    console.log('old path')
    const jsonData = req.body.excel
    const excelData = await convertJsonToExcel(jsonData)
    const link = await uploadToAWS(req.body.config, excelData)
    return res.json(link)
  } catch (error) {
    console.log('error', error)
    res.status(400).json({ message: 'error in your request payload', error: error.message, rawError: error })
  }
})

router.post('*', async (req, res) => {
  console.log('nothing to do path')
  return res.json({ message: "wrong method or path" })
})

const uploadToAWS = async (config, excelData , isStream = false) => {
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
  if (isStream) {
    return s3.upload(dataset).promise()
  } else {
    const response = await s3.upload(dataset).promise()
    return response.Location
  }
}

const getExcelFunc = (type) => {
  const typeObject = {
    'styled': convertJsonToStyledExcel,
    'common-styled': convertJsonToCommonStyledExcel,
    'simple': convertJsonToExcel,
  }
  const fn = typeObject[type]
  if (!fn) throw new Error(`Invalid type: ${type}`)
  return fn
}

async function injectFormulasIntoSheet(workbook, data, config) {
  const lookupSheet = config.lookupSheet
  const hiddenSheet = workbook.getWorksheet(lookupSheet) || workbook.addWorksheet(lookupSheet)
  hiddenSheet.state = 'veryHidden'

  const headerKeys = [
    ...new Set(
      data.flatMap(row => Object.keys(row))
    )
  ];

  // Column-wise data insertion into hidden sheet
  headerKeys.forEach((key, idx) => {
    const values = data.flatMap(row => row[key] ?? []);
    values.forEach((val, rowIndex) => {
      hiddenSheet.getRow(rowIndex + 2).getCell(idx + 1).value = val ?? "";
    });
    hiddenSheet.getRow(1).getCell(idx + 1).value = key; 
  });
  const findCol = (sheet, aliases) => {
    const headerRow = sheet.getRow(1)
    const lookupNames = Array.isArray(aliases) ? aliases : [aliases]

    for (let col = 1; col <= sheet.columnCount; col++) {
      const val = headerRow.getCell(col)?.value
      let text = '';

      if (typeof val === 'object') {
        if (Array.isArray(val.richText)) {
          text = val.richText.map(rt => rt.text).join('');
        } else if (val.result) {
          text = val.result;
        }
      } else if (val !== null && val !== undefined) {
        text = val;
      }

      const normalized = String(text || '').trim().toLowerCase();

      for (const alias of lookupNames) {
        if (!alias) continue;
        const normAlias = String(alias).trim().toLowerCase()
        if (normalized === normAlias) return col
        if (normalized.replace(/\s+/g, "_") === normAlias) return col
        if (normalized.replace(/[^a-z0-9]/gi, "") === normAlias.replace(/[^a-z0-9]/gi, "")) return col
      }
    }
    return null
  }
  for (const sheetConfig of config.dependentSheet) {
    const sheet = workbook.getWorksheet(sheetConfig.name);
    if (!sheet) throw new Error(`Sheet ${sheetConfig.name} not found`);

    let lastRow = 0;

    //Primary key dropdown (if exists)
    let primaryCol = null;

    if (sheetConfig.primaryKey) {
      primaryCol = findCol(sheet, sheetConfig.primaryKey);
      if (!primaryCol) throw new Error(`Primary key not found in ${sheetConfig.name}`);

      const lookupCol = findCol(hiddenSheet, sheetConfig.primaryKey);
      if (!lookupCol) throw new Error(`Primary key not found in Lookups sheet`);
      const lookupColLetter = String.fromCharCode(64 + lookupCol);
      lastRow=data[0][sheetConfig.primaryKey[0]].length + 1
      const primaryCell = sheet.getRow(2).getCell(primaryCol);
      primaryCell.dataValidation = {
        type: "list",
        allowBlank: true,
        formulae: [`${hiddenSheet.name}!$${lookupColLetter}$2:$${lookupColLetter}$${lastRow}`],
      };
    }
    //Dependent dropdowns
    for (const depGroup of sheetConfig.dependentKeys || []) {
      const depCol = findCol(sheet, depGroup); 
      if (!depCol) continue;

      const lookupCol = findCol(hiddenSheet, depGroup); 
      if (!lookupCol) continue;
      const lookupColLetter = String.fromCharCode(64 + lookupCol);
      lastRow=data[0][depGroup[0]].length + 1
      const depCell = sheet.getRow(2).getCell(depCol);
      depCell.dataValidation = {
        type: "list",
        allowBlank: true,
        formulae: [`${hiddenSheet.name}!$${lookupColLetter}$2:$${lookupColLetter}$${lastRow}`],
      };
    }

    //Lookup autofill (only if primary key exists)
    if (primaryCol && sheetConfig.lookupKeys) {
      for (const lookupGroup of sheetConfig.lookupKeys) {
        const lookupCol = findCol(sheet, lookupGroup); 
        if (!lookupCol) continue;
        const lookupKey = Array.isArray(lookupGroup) ? lookupGroup[0] : lookupGroup;
        const lookupColIndex = headerKeys.indexOf(lookupKey) + 1;
        lastRow= hiddenSheet.lastRow.number
        const lookupCell = sheet.getRow(2).getCell(lookupCol);
        const primaryAddress = sheet.getRow(2).getCell(primaryCol).address
        lookupCell.value = {
          formula: `IF(${primaryAddress}="","",VLOOKUP(${primaryAddress},${hiddenSheet.name}!$A$2:$${String.fromCharCode(
            64 + headerKeys.length
          )}$${lastRow},${lookupColIndex},FALSE))`,
        };
      }
    }
  }
}

function resolveStreamOptions(second) {
  if (second == null) return {}
  if (typeof second !== 'object') return {}
  if (second.isStream === true || second.stream != null) return second
  return {}
}

function resolveStyledOptions(second) {
  if (second == null) return {}

  if (typeof second !== 'object') {
    return { defaultStyle: second }
  }

  return {
    isStream: second.isStream === true,
    stream: second.stream ?? null,
    defaultStyle: second.defaultStyle ?? second,
  }
}

async function convertJsonToExcel(jsonData, second = {}) {
  const options = resolveStreamOptions(second)
  const { isStream = false, stream = null } = options

  if (!isStream) {
    const workbook = xlsx.utils.book_new()

    Object.keys(jsonData).forEach((sheetName) => {
      const rows = jsonData[sheetName] || []
      const worksheet = xlsx.utils.json_to_sheet(rows)
      xlsx.utils.book_append_sheet(workbook, worksheet, sheetName)
    })

    return xlsx.write(workbook, { type: 'buffer', bookType: 'xlsx' })
  }

  const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
    stream,
    useStyles: true,
    useSharedStrings: true,
  })

  for (const sheetName of Object.keys(jsonData)) {
    const rows = jsonData[sheetName] || []
    const worksheet = workbook.addWorksheet(sheetName)
    if (rows.length === 0) {
      worksheet.commit()
      continue
    }
    const keys = Object.keys(rows[0])
    const headerRow = worksheet.addRow(keys)
    headerRow.commit()
    for (const row of rows) {
      const excelRow = worksheet.addRow(keys.map((k) => row[k]))
      excelRow.commit()
    }
    worksheet.commit()
  }

  await workbook.commit()
  return true
}

async function convertJsonToStyledExcel(jsonData, second = {}) {
  const options = resolveStyledOptions(second)
  const { isStream = false, stream = null, defaultStyle } = options

  const workbook = isStream
  ? new ExcelJS.stream.xlsx.WorkbookWriter({
      stream,
      useStyles: true,
      useSharedStrings: true,
    })
  : new ExcelJS.Workbook()

for (const sheetName of Object.keys(jsonData)) {
  const sheet = jsonData[sheetName]
  const worksheet = workbook.addWorksheet(sheetName)

  // Headers
  if (sheet.headers) {
    for (const [headerText, style] of Object.entries(sheet.headers)) {
      const row = worksheet.addRow([headerText])
      if (style) row.getCell(1).style = style
      if (isStream) row.commit()
    }
  }

  // SubHeaders
  if (sheet.subHeaders) {
    const values = Object.values(sheet.subHeaders).map(s => s.value)
    const row = worksheet.addRow(values)

    Object.values(sheet.subHeaders).forEach((s, i) => {
      if (s.style) row.getCell(i + 1).style = s.style
    })

    if (isStream) row.commit()
  }

  // SubHeadersData
  if (sheet.subHeadersData) {
    const obj = sheet.subHeadersData[0]
    const values = Object.keys(obj).map(k => obj[k].value)

    const row = worksheet.addRow(values)

    Object.keys(obj).forEach((key, i) => {
      if (obj[key].style) {
        row.getCell(i + 1).style = obj[key].style
      }
    })

    if (isStream) row.commit()
  }

  // Column headers
  const columnHeaders = sheet.keysValue
    ? Object.values(sheet.keysValue).map(k => k.value)
    : Object.keys(sheet.data[0])

  const headerRow = worksheet.addRow(columnHeaders)

  if (sheet.keysValue) {
    Object.values(sheet.keysValue).forEach((k, i) => {
      if (k.style) headerRow.getCell(i + 1).style = k.style
    })
  }

  if (isStream) headerRow.commit()

  const headers = Object.keys(sheet.data[0])

  // Track data start row
  const dataStartRow = worksheet.rowCount + 1

  // Rows
  for (const row of sheet.data) {
    const excelRow = worksheet.addRow(headers.map(k => row[k].value))

    headers.forEach((key, colIndex) => {
      const cell = excelRow.getCell(colIndex + 1)

      if (row[key].style) {
        cell.style = row[key].style
      }
    })

    if (isStream) excelRow.commit()
  }

  // ✅ Validation ONLY in non-stream mode
  if (!isStream) {
    let validationColPointer = headers.length + 2

    headers.forEach((key, colIndex) => {
      const firstRow = sheet.data[0]?.[key]

      if (firstRow?.input?.options?.length) {
        const options = firstRow.input.options
        const col = worksheet.getColumn(validationColPointer)
        const colLetter = col.letter

        // Fill hidden column
        options.forEach((opt, i) => {
          worksheet.getCell(`${colLetter}${i + 1}`).value = opt
        })

        const range = `$${colLetter}$1:$${colLetter}$${options.length}`

        for (let r = dataStartRow; r < dataStartRow + sheet.data.length; r++) {
          worksheet.getRow(r).getCell(colIndex + 1).dataValidation = {
            type: 'list',
            allowBlank: true,
            formulae: [range],
            showErrorMessage: true,
            errorTitle: 'Invalid Selection',
            error: 'Please select a value from the dropdown',
          }
        }

        col.hidden = true
        validationColPointer++
      }
    })
  }

  // Column width
  worksheet.columns.forEach(col => {
    col.width = defaultStyle?.width || 20
  })

  if (isStream) worksheet.commit()
}

if (isStream) {
  await workbook.commit()
  return true
} else {
  return await workbook.xlsx.writeBuffer()
}
}

async function convertJsonToCommonStyledExcel(data, second = {}) {
  const options = resolveStreamOptions(second)
  const { isStream = false, stream = null } = options

  const workbook = isStream
    ? new ExcelJS.stream.xlsx.WorkbookWriter({
        stream,
        useStyles: true,
        useSharedStrings: true,
      })
    : new ExcelJS.Workbook()

  for (const [sheetName, sheetRows] of Object.entries(data)) {
    const worksheet = workbook.addWorksheet(sheetName)
    sheetRows.forEach((row, rowIndex) => {
      let colIndex = 1
      row.forEach((cell) => {
        const currentCell = worksheet.getCell(rowIndex + 1, colIndex)
        currentCell.value = cell.value
        if (cell.style) {
          Object.assign(currentCell.style, cell.style)
        }
        if (cell.colspan || cell.rowspan) {
          const startRow = rowIndex + 1
          const startCol = colIndex
          const endRow = startRow + (cell.rowspan || 1) - 1
          const endCol = startCol + (cell.colspan || 1) - 1

          worksheet.mergeCells(startRow, startCol, endRow, endCol)
          colIndex += cell.colspan || 1
        } else if (cell.dropdown) {
          const formulae = [`"${cell.dropdown.join(',')}"`]
          worksheet.getCell(rowIndex + 1, colIndex).dataValidation = {
            type: 'list',
            allowBlank: false,
            formulae,
            showErrorMessage: true,
            errorTitle: 'Invalid Selection',
            error: 'Please select a value from the dropdown',
          }
          colIndex++
        } else {
          colIndex++
        }
      })
    })
    worksheet.columns.forEach((column) => {
      column.width = column.width ? column.width : 20
    })
    if (isStream) worksheet.commit()
  }

  if (isStream) {
    await workbook.commit()
    return true
  }
  return workbook.xlsx.writeBuffer()
}


// const servicePrefix = process.env.SERVICE_PREFIX || '/'
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
  const response = await handler(event, context, callback)
  return response
}
