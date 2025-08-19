# create-entra-app

script to create M365 Entra Apps.
needa an "permission.json" like the exmaples and created an "auth_credentials.json" with ID´s and the ClientSecret.

## Usage

### Create an Entra APP
`.\create-entra-app.ps1 [-PermissionFilePath .\permissions-emailrelay.json] [-CredentialOutputPath .\auth_credentials.json]`

### Get Permission ID´s (config supports name)

Get only one permission ID
`.\create-entra-app.ps1 -GetEntraPermissionID SMTP.SendAsApp -ResourceID 00000002-0000-0ff1-ce00-000000000000`

Get multiple permission IDs
`.\create-entra-app.ps1 -GetEntraPermissionID "SMTP.SendAsApp,POP.AccessAsApp" -ResourceID 00000002-0000-0ff1-ce00-000000000000`

Get both resource and permission id(s)
`.\create-entra-app.ps1 -GetEntraPermissionID "SMTP.SendAsApp,POP.AccessAsApp" -GetEntraResourceID "Office 365 Exchange Online"`