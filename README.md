# Export Excel Service 

### APIs

**NOTE:** Don't forgot to add http/https at URL, else redirect won't work.
like: `{"url":"divamtech.com"}` won't work, have to add **http/https** as below.

```sh
curl --location 'https://export-service.webledger.in/api/jsonToExcel' \
--header 'x-auth-token: ' \
--header 'Content-Type: application/json' \
--data '{
  "config": {
    "s3FilePublic": true,
    "s3Region": "<your_s3_region>",
    "s3Bucket": "<your_s3_bucket_name>",
    "s3KeyId": "<your_s3_key_id>",
    "s3SecretKey": "<your_s3_secret_key>",
    "s3Path": "<you_s3_bucket_path>.xlsx"
  },
  "excel": {
    "Sheet 1": [
      {
        "name": "John",
        "age": 30,
        "city": "New York"
      },
      {
        "name": "Alice",
        "age": 25,
        "city": "Los Angeles"
      }
    ],
    "Sheet 2": [
      {
        "name": "Paul",
        "age": 35,
        "city": "New York"
      },
      {
        "name": "Alicea",
        "age": 25,
        "city": "Los Angeles"
      }
    ]
  }
}'
```
## Voila!
