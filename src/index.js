const express = require('express')
const serverless = require('serverless-http')
const AWS = require('aws-sdk')
const xlsx = require('xlsx')
require('dotenv').config()
const AUTH_TOKEN = process.env.AUTH_TOKEN

const app = express()
app.use(express.json())

const router = express.Router()
router.use((req, res, next) => {
  const token = req.get('x-auth-token')
  if (!!token && token === AUTH_TOKEN) {
    next()
  } else {
    res.status(401).json({ message: 'Invalid auth token' })
  }
})

router.post('/jsonToExcel', async (req, res) => {
  const { s3FilePublic, s3Region, s3Bucket, s3KeyId, s3SecretKey, s3Path } = req.body.config
  const jsonData = req.body.excel
  const excelData = convertJsonToExcel(jsonData)
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
  return res.json(response.Location)
})

function convertJsonToExcel(jsonData) {
  const workbook = xlsx.utils.book_new();
  
  Object.keys(jsonData).forEach((sheetName) => {
    const worksheet = xlsx.utils.json_to_sheet(jsonData[sheetName]);
    xlsx.utils.book_append_sheet(workbook, worksheet, sheetName);
  });

  return xlsx.write(workbook, { type: 'buffer', bookType: 'xlsx' });
}

app.use('/api', router)

const startServer = async () => {
  app.listen(3000, () => {
    console.log('listening on port 3000!')
  })
}
startServer()

//lambda handling
const handler = serverless(app)

exports.handler = async (event, context, callback) => {
  const response = handler(event, context, callback)
  return response
}

