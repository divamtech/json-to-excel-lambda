const AWS = require('aws-sdk');
const xlsx = require('xlsx');
require('dotenv').config();
const AUTH_TOKEN = process.env.AUTH_TOKEN;

const convertJsonToExcel = (jsonData) => {
  const workbook = xlsx.utils.book_new();
  
  Object.keys(jsonData).forEach((sheetName) => {
    const worksheet = xlsx.utils.json_to_sheet(jsonData[sheetName]);
    xlsx.utils.book_append_sheet(workbook, worksheet, sheetName);
  });

  return xlsx.write(workbook, { type: 'buffer', bookType: 'xlsx' });
};

const handler = async (event) => {
  const { config, excel } = JSON.parse(event.body);
  const { s3FilePublic, s3Region, s3Bucket, s3KeyId, s3SecretKey, s3Path } = config;
  
  if (!AUTH_TOKEN || event.headers['x-auth-token'] !== AUTH_TOKEN) {
    return {
      statusCode: 401,
      body: JSON.stringify({ message: 'Invalid auth token' }),
    };
  }

  const excelData = await convertJsonToExcel(excel);

  AWS.config.update({ accessKeyId: s3KeyId, secretAccessKey: s3SecretKey, region: s3Region, signatureVersion: 'v4' });
  const s3 = new AWS.S3();

  const dataset = {
    Bucket: s3Bucket,
    Key: s3Path,
    Body: excelData,
    ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    ACL: s3FilePublic ? 'public-read' : 'private',
  };

  try {
    const response = await s3.upload(dataset).promise();
    return {
      statusCode: 200,
      body: JSON.stringify(response.Location),
    };
  } catch (error) {
    return {
      statusCode: 500,
      body: JSON.stringify({ message: 'Error uploading file to S3', error }),
    };
  }
};

exports.handler = handler;
