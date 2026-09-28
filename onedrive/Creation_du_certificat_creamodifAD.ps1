#Création du certificat :
 
$cert = New-SelfSignedCertificate `
 -Subject "CN=MSGraphExchangeOnlineAuth20260717_2" `
 -KeyExportPolicy Exportable `
 -KeySpec Signature `
 -KeyUsage DigitalSignature `
 -KeyLength 2048 `
 -KeyAlgorithm RSA `
 -HashAlgorithm SHA256 `
 -CertStoreLocation "Cert:\LocalMachine\My" `
 -NotAfter (Get-Date).AddYears(2)



 
#Export en PFX :
 
$pwd = ConvertTo-SecureString -String "MotDePasseFort123!" -Force -AsPlainText
 
Export-PfxCertificate `
 -Cert $cert `
 -FilePath "$env:USERPROFILE\Desktop\MSGraphExchangeOnlineAuth20260717_2.pfx" `
 -Password $pwd
 
 
#Conversion en .cer :

 
Export-Certificate -Cert $cert -FilePath "$env:USERPROFILE\Desktop\MSGraphExchangeOnlineAuth20260717_2.cer"
 