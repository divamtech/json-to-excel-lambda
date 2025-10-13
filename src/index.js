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

router.post('/lambda/json-to-excel/from-link', async (req, res) => {
  console.log('from link working')
  try {
    const { jsonUrl, type } = req.body
    let jsonData = await fetch(jsonUrl)
    jsonData = await jsonData.json()
    const defaultStyle = jsonData.config.default_style
    let excelFunc = null
    switch (type) {
      case 'styled':
        excelFunc = convertJsonToStyledExcel
        break
      case 'common-styled':
        excelFunc = convertJsonToCommonStyledExcel
        break
      case 'simple':
        excelFunc = convertJsonToExcel
        break
      default:
        throw new Error('Invalid type')
    }
    const excelData = await excelFunc(jsonData.excel, defaultStyle)
    const url = await uploadToAWS(jsonData.config, excelData)
    return res.json({ url })  
  } catch (error) {
    console.log('error', error)
    res.status(400).json({ message: 'error in your request payload', error: error.message, rawError: error })
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
    let primaryCell = null;

    if (sheetConfig.primaryKey) {
      primaryCol = findCol(sheet, sheetConfig.primaryKey);
      if (!primaryCol) throw new Error(`Primary key not found in ${sheetConfig.name}`);
      primaryCell = sheet.getRow(2).getCell(primaryCol);

      const lookupCol = findCol(hiddenSheet, sheetConfig.primaryKey);
      if (!lookupCol) throw new Error(`Primary key not found in Lookups sheet`);
      const lookupColLetter = String.fromCharCode(64 + lookupCol);
      lastRow=data[0][sheetConfig.primaryKey[0]].length + 1
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

      const depCell = sheet.getRow(2).getCell(depCol);

      const lookupCol = findCol(hiddenSheet, depGroup); 
      if (!lookupCol) continue;
      const lookupColLetter = String.fromCharCode(64 + lookupCol);
      lastRow=data[0][depGroup[0]].length + 1
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
        const lookupCell = sheet.getRow(2).getCell(lookupCol);
        lastRow= hiddenSheet.lastRow.number
        lookupCell.value = {
          formula: `IF(${primaryCell.address}="","",VLOOKUP(${primaryCell.address},${hiddenSheet.name}!$A$2:$${String.fromCharCode(
            64 + headerKeys.length
          )}$${lastRow},${lookupColIndex},FALSE))`,
        };
      }
    }
  }
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
    if (jsonData[sheetName].headers) {
      Object.entries(jsonData[sheetName].headers).forEach(([headerText, style]) => {
        const headerRow = worksheet.addRow([headerText])
        const headerStyle = style
        headerRow.getCell(1).style = headerStyle
      })
    }

    // Add subHeaders and corresponding subHeadersData after headers, if they exist
    if (jsonData[sheetName].subHeaders) {
      const subHeaderValues = Object.values(jsonData[sheetName].subHeaders).map(subHeaderObj => subHeaderObj.value);

      // Add subHeaders in the next row
      worksheet.addRow(subHeaderValues);

      // Apply styles for subHeaders
      Object.values(jsonData[sheetName].subHeaders).forEach((subHeaderObj, index) => {
        const subHeaderCell = worksheet.getRow(worksheet.rowCount).getCell(index + 1); // Row number dynamically
        if (subHeaderObj.style) {
          subHeaderCell.style = subHeaderObj.style;
        }
      });

      // Now add the subHeadersData corresponding to each subHeader, if available
      if (jsonData[sheetName].subHeadersData) {
        const subHeaderDataObject = jsonData[sheetName].subHeadersData[0]; // Access the first object in the array
        const subHeaderDataValues = Object.keys(subHeaderDataObject).map(key => {
          return subHeaderDataObject[key].value;
        });
        worksheet.addRow(subHeaderDataValues);

        Object.keys(jsonData[sheetName].subHeadersData).forEach((key, index) => {
          const subHeaderDataCell = worksheet.getRow(worksheet.rowCount).getCell(index + 1);
          if (jsonData[sheetName].subHeadersData[key].style) {
            subHeaderDataCell.style = jsonData[sheetName].subHeadersData[key].style;
          }
        });
      }
    }

    // Now add the KeyValue headers from 'keysValue' (if available)
    const columnHeaders = jsonData[sheetName].keysValue
      ? Object.values(jsonData[sheetName].keysValue).map((keyValueObj) => keyValueObj.value)
      : Object.keys(jsonData[sheetName].data[0]);

    const headerRow = worksheet.addRow(columnHeaders)

    // Apply styles from "keysValue" to the header row
    if (jsonData[sheetName].keysValue) {
      Object.values(jsonData[sheetName].keysValue).forEach((keyValueObj, index) => {
        const cell = headerRow.getCell(index + 1)
        if (keyValueObj.style) {
          cell.style = keyValueObj.style
        }
      })
    }

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
        if (row[key].input) {
          const formulae = row[key].input.options ? [`"${row[key].input.options.join(',')}"`] : [];
          const validationType = row[key].input.type || ''; 

          const dropdownColumnIndex = colIndex + 1;
          // Apply data validation for all rows in this column (from the second row onwards)
          worksheet.getColumn(dropdownColumnIndex).eachCell({ includeEmpty: true }, (cell, rowNumber) => {
            if (rowNumber > 1) { // Skip the header row
              cell.dataValidation = {
                type: validationType,
                allowBlank: false,
                formulae: formulae,
                showErrorMessage: true,
                errorTitle: 'Invalid Selection',
                error: 'Please select a value from the dropdown',
              };
            }
          });
        }

      })
    })

    worksheet.columns.forEach((column) => {
      column.width = defaultStyle ? defaultStyle.width : 10
    })
  })

  const buffer = await workbook.xlsx.writeBuffer()
  return buffer
}

async function convertJsonToCommonStyledExcel(data) {
  const workbook = new ExcelJS.Workbook();
  for (const [sheetName, sheetRows] of Object.entries(data)) {
      const worksheet = workbook.addWorksheet(sheetName);
      sheetRows.forEach((row, rowIndex) => {
          let colIndex = 1;
          row.forEach(cell => {
              const currentCell = worksheet.getCell(rowIndex + 1, colIndex);
              currentCell.value = cell.value;
              if (cell.style) {
                  Object.assign(currentCell.style, cell.style);
              }
              if (cell.colspan || cell.rowspan) {
                  const startRow = rowIndex + 1;
                  const startCol = colIndex;
                  const endRow = startRow + (cell.rowspan || 1) - 1;
                  const endCol = startCol + (cell.colspan || 1) - 1;

                  worksheet.mergeCells(startRow, startCol, endRow, endCol);
                  colIndex += (cell.colspan || 1);
              } 
              else if (cell.dropdown) {
                  const formulae = [`"${cell.dropdown.join(",")}"`];
                  worksheet.getCell(rowIndex + 1, colIndex).dataValidation = {
                      type: 'list',
                      allowBlank: false,
                      formulae: formulae,
                      showErrorMessage: true,
                      errorTitle: 'Invalid Selection',
                      error: 'Please select a value from the dropdown',
                  };
                  colIndex++;
              } else {
                  colIndex++;
              }
          });
      });
      worksheet.columns.forEach(column => {
          column.width = column.width ? column.width : 20; // Set a default width
      });
  }
  const buffer = await workbook.xlsx.writeBuffer();
  return buffer
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
