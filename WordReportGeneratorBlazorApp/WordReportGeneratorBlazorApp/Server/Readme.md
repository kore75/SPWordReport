
# WebAPI


```bash
openssl req -x509 -newkey rsa:4096 -keyout devkoeli.sharepoint.com.pem -out devkoeli.sharepoint.com.pem -sha256 -days 365
openssl pkcs12 -export -in devkoeli.sharepoint.com.pem -out devkoeli.sharepoint.com.pfx

```

# Samplate Data
List Id:a0e24368-43ee-434c-ae02-026a179d1abc

Id 1

DocLib=6176fc55-78b7-4e29-b92c-44816540ac7e
https://devkoeli.sharepoint.com/sites/Development/_api/web/lists/getByTitle('Documents')/Id
Id 1