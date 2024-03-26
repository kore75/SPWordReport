
# WebAPI


```bash
openssl req -x509 -newkey rsa:4096 -keyout devkoeli.sharepoint.com.pem -out devkoeli.sharepoint.com.pem -sha256 -days 365
openssl pkcs12 -export -in devkoeli.sharepoint.com.pem -out devkoeli.sharepoint.com.pfx

```