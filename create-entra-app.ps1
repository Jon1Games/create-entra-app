param(    
    [Parameter(Mandatory=$false)]
    [string]$CredentialOutputPath = ".\auth_credentials.json",

    [Parameter(Mandatory=$false)]
    [string]$PermissionFilePath = ".\permissions.json",

    [Parameter(Mandatory=$false)]
    [string]$GetEntraPermissionID,
    
    [Parameter(Mandatory=$false)]
    [string]$ResourceId,

    [Parameter(Mandatory=$false)]
    [string]$GetEntraResourceID
)

function Load-Permissions {
    # Load required permissions from permissions.json
    if (Test-Path $PermissionFilePath) {
        Write-Host "Loading $PermissionFilePath..." -ForegroundColor Yellow
        $Permissions = Get-Content $PermissionFilePath | ConvertFrom-Json
        
        # Transform permission names to IDs if needed
        if ($Permissions.EntraApplication -and $Permissions.EntraApplication.Permission) {
            foreach ($permission in $Permissions.EntraApplication.Permission) {
                # Check if ResourceAppId contains a name instead of an ID (GUID)
                if ($permission.ResourceAppId -and $permission.ResourceAppId -notmatch '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$') {
                    Write-Host "Converting resource name '$($permission.ResourceAppId)' to ID..." -ForegroundColor Yellow
                    $resourceId = Get-EntraResourceID -ResourceName $permission.ResourceAppId
                    if ($resourceId) {
                        $permission.ResourceAppId = $resourceId
                    } else {
                        Write-Error "Failed to get resource ID for '$($permission.ResourceAppId)'"
                        exit 1
                    }
                }
                
                # Check if ResourceAccess contains permission names instead of IDs
                if ($permission.ResourceAccess) {
                    foreach ($access in $permission.ResourceAccess) {
                        if ($access.Id -and $access.Id -notmatch '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$') {
                            Write-Host "Converting permission name '$($access.Id)' to ID..." -ForegroundColor Yellow
                            $permissionId = Get-EntraPermissionID -PermissionName $access.Id -ResourceId $permission.ResourceAppId
                            if ($permissionId) {
                                $access.Id = $permissionId
                            } else {
                                Write-Error "Failed to get permission ID for '$($access.Id)'"
                                exit 1
                            }
                        }
                    }
                }
            }
        }
        
        return $Permissions
    } else {
        Write-Error "$PermissionFilePath not found. This file is required to define the necessary permissions."
        exit 1
    }
}

# Function to connect to Microsoft Entra
function Connect-ToMicrosoftEntra {    
    Write-Host "Connecting to Microsoft Entra..." -ForegroundColor Yellow
    
    try {
        try {
            if (Get-EntraContext -ErrorAction SilentlyContinue) {
                return $true
            }
        }
        catch {
            # Ignore errors if no existing session
        }

        Connect-Entra -NoWelcome -Scopes @(
            "Application.ReadWrite.All",
            "AppRoleAssignment.ReadWrite.All",
            "Directory.Read.All"
        )
        $context = Get-EntraContext
        Write-Host "✓ Successfully connected to Microsoft Entra" -ForegroundColor Green
        Write-Host "  Account: $($context.Account)" -ForegroundColor Cyan
        Write-Host "  Tenant: $($context.TenantId)" -ForegroundColor Cyan
        return $true
    }
    catch {
        Write-Error "Failed to connect to Microsoft Entra: $_"
        return $false
    }
}

# Function to connect to Microsoft Exchange Online
function Connect-ToMicrosoftExchangeOnline {
    Write-Host "Connecting to Microsoft Exchange Online..." -ForegroundColor Yellow

    try {
        # Connect using Microsoft.Exchange.Management.PowerShell
        Connect-ExchangeOnline -ShowBanner:$false
        Write-Host "✓ Successfully connected to Microsoft Exchange Online" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Error "Failed to connect to Microsoft Exchange Online: $_"
        return $false
    }
}

# Function to get the ID of an Entra Resource from the name
function Get-EntraResourceID {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ResourceName
    )

    Write-Host "Getting Entra Resource ID for: $ResourceName" -ForegroundColor Yellow

    try {
        $graphServicePrincipal = Get-EntraServicePrincipal -Filter "DisplayName eq '$ResourceName'"
        if ($graphServicePrincipal) {
            Write-Host "✓ Found Entra Resource ID for '$ResourceName': $($graphServicePrincipal.Id)" -ForegroundColor Green
            Write-Host "✓ Found Entra Resource AppID for '$ResourceName': $($graphServicePrincipal.AppId)" -ForegroundColor Green
            return $graphServicePrincipal.AppId
        } else {
            Write-Warning "Entra Resource '$ResourceName' not found."
            return $null
        }
    }
    catch {
        Write-Error "Failed to get Entra Resource ID: $_"
        return $null
    }
}

# Function to get the ID of an Entra Permission from the name
function Get-EntraPermissionID {
    param(
        [Parameter(Mandatory=$true)]
        [string]$PermissionName,

        [Parameter(Mandatory=$true)]
        [string]$ResourceId
    )

    Write-Host "Getting Entra Permission ID for: $PermissionName" -ForegroundColor Yellow

    try {
        $graphServicePrincipal = Get-EntraServicePrincipal -Filter "AppId eq '$ResourceId'"
        
        # Split permission names by comma or semicolon
        $permissionNames = $PermissionName -split '[,;]' | ForEach-Object { $_.Trim() }
        
        $permissionIds = @()
        foreach ($permName in $permissionNames) {
            $permission = ($graphServicePrincipal.AppRoles | Where-Object { $_.Value -eq $permName })
            if ($permission) {
                $permissionIds += $permission.Id
                Write-Host "✓ Found permission '$permName': $($permission.Id)" -ForegroundColor Green
            } else {
                Write-Warning "Permission '$permName' not found."
            }
        }
        
        if ($permissionIds.Count -gt 0) {
            return $permissionIds
        } else {
            Write-Error "No valid permissions found."
            return $null
        }
    }
    catch {
        Write-Error "Failed to get Entra Permission ID: $_"
        return $null
    }
}

# Function to create app registration using Microsoft.Entra.Applications
function New-EntraAppRegistration {
    param(
        [Parameter(Mandatory=$true)]
        [string]$DisplayName,
        [Parameter(Mandatory=$true)]
        [array]$EntraPermissions,
        [Parameter(Mandatory=$true)]
        [int]$SecretLifetimeMonths
    )
    
    Write-Host "Creating app registration: $DisplayName" -ForegroundColor Yellow
    
    try {
        # Create the app registration using Microsoft.Entra.Applications
        $appParams = @{
            DisplayName = $DisplayName
            RequiredResourceAccess = $EntraPermissions
        }
        
        $app = New-EntraApplication @appParams
        if (!$app) {
            return $null
        }

        Write-Host "✓ App registration created successfully" -ForegroundColor Green
        Write-Host "  Application ID: $($app.AppId)" -ForegroundColor Cyan
        Write-Host "  Object ID: $($app.Id)" -ForegroundColor Cyan

        # Create a Service Principal for the app registration
        $graphServicePrincipal = New-EntraServicePrincipal -AppId $app.AppId
        if (!$graphServicePrincipal) {
            return $null
        }

        Write-Host "✓ Service Principal created successfully" -ForegroundColor Green
        Write-Host "  Service Principal Object ID: $($graphServicePrincipal.Id)" -ForegroundColor Cyan

        # Grand admin consent
        # Get app role IDs from the required permissions
        $appRoles = @()
        foreach ($permission in $EntraPermissions) {
            $tmp = @()
            foreach ($role in $permission.ResourceAccess) {
                $tmp += $role.Id
            }
            $appRoles += @{
                ResourceAppId = $permission.ResourceAppId
                ResourceAccess = $tmp
            }
        }

        # Grant admin consent for each app role
        foreach ($appRole in $appRoles) {
            $graphApiId = $appRole.ResourceAppId
            $ResourceId = (Get-EntraServicePrincipal -Filter "AppId eq '$graphApiId'").Id
            foreach ($roleId in $appRole.ResourceAccess) {
                $adminConsent = New-EntraServicePrincipalAppRoleAssignment -PrincipalId $graphServicePrincipal.Id -ServicePrincipalId $graphServicePrincipal.Id -ResourceId $ResourceId -Id $roleId
                if (!$adminConsent) {
                    return $null
                }
            }
        }

        Write-Host "✓ Admin consent granted successfully" -ForegroundColor Green 

        # Create client secret with 120-month lifetime
        $secretDescription = "Generated by iOS Device Management Script - $(Get-Date -Format 'yyyy-MM-dd')"
        $secretEndDate = (Get-Date).AddMonths($SecretLifetimeMonths)
        
        $passwordCredential = @{
            DisplayName = $secretDescription
            EndDateTime = $secretEndDate
        }
        
        $secret = New-EntraApplicationPassword -ObjectId $app.Id -PasswordCredential $passwordCredential
        if (!$secret) {
            return $null
        }

        Write-Host "✓ Client secret created successfully" -ForegroundColor Green
        Write-Host "  Secret expires: $secretEndDate" -ForegroundColor Cyan
        Write-Host "  Secret lifetime: $SecretLifetimeMonths months" -ForegroundColor Cyan

        return @{
            AppId = $app.AppId
            ObjectId = $app.Id
            TenantId = (Get-EntraContext).TenantId
            ClientSecret = $secret.SecretText
            SecretId = $secret.KeyId
            ExpiryDate = $secretEndDate
        }
    }
    catch {
        Write-Error "Failed to create app registration: $_"
        return $null
    }
}

# Function to save credentials to file (compatible with get-ios-information.ps1)
function Save-Credentials {
    param(
        [hashtable]$Credentials,
        [string]$FilePath,
        [string]$Name
    )
    
    Write-Host "Saving credentials to: $FilePath" -ForegroundColor Yellow
    
    try {
        $credentialsJson = [ordered]@{
            TenantId = $Credentials.TenantId
            ClientId = $Credentials.AppId
            AppName = $Name
            AuthType = "ClientSecret"
            ClientSecret = $Credentials.ClientSecret
            SecretExpires = $Credentials.ExpiryDate.ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
            CreatedDate = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
        } | ConvertTo-Json 
        
        $credentialsJson | Out-File -FilePath $FilePath -Encoding UTF8
        Write-Host "✓ Credentials saved successfully" -ForegroundColor Green
        
        return $true
    }
    catch {
        Write-Error "Failed to save credentials: $_"
        return $false
    }
}

# Main execution
try {
    # Connect to entra as it is always needed
    if (!(Connect-ToMicrosoftEntra)) {
        throw "Failed to connect to Microsoft Entra"
    }

    # GetEntraResourceID
    if ($GetEntraResourceID) {
        $ResourceId = Get-EntraResourceID -ResourceName $GetEntraResourceID
        if ($ResourceId) {
            if (!$GetEntraPermissionID) {
                exit 0
            }
        } else {
            Write-Error "Failed to retrieve Entra Resource ID for '$GetEntraResourceID'"
            exit 1
        }
    }

    # GetEntraPermissionID
    if ($GetEntraPermissionID) {
        if (!$ResourceId) {
            Write-Error "Resource ID is required to get permission ID, give it with -ResourceId"
            exit 1
        }

        $permissionId = Get-EntraPermissionID -PermissionName $GetEntraPermissionID -ResourceId $ResourceId
        if ($permissionId) {
            exit 0
        } else {
            Write-Error "Failed to retrieve Entra Permission ID for '$GetEntraPermissionID'"
            exit 1
        }
    }

    $Permissions = Load-Permissions
    if (!$Permissions) {
        throw "Failed to load permissions"
    }

    if ($Permissions.EntraApplication) {
        $appInfo = New-EntraAppRegistration -DisplayName $Permissions.EntraApplication.Name -EntraPermissions $Permissions.EntraApplication.Permission -SecretLifetimeMonths $Permissions.EntraApplication.SecretLifetimeMonths
        if (!$appInfo) {
            throw "Failed to create app registration"
        }

        if (!(Save-Credentials -Credentials $appInfo -FilePath $CredentialOutputPath -Name $Permissions.EntraApplication.Name)) {
            throw "Failed to save credentials"
        }
    }

    if ($Permissions.ExchangeOnlinePermissions) {
        if (!(Connect-ToMicrosoftExchangeOnline)) {
            throw "Failed to connect to Microsoft Exchange Online"
        }

        $app = $null
        foreach ($permission in $Permissions.ExchangeOnlinePermissions) {
            if (-not $permission.UserPrincipalName -or -not $permission.AccessRights) {
                Write-Error "Invalid Exchange Online permission configuration"
                exit 1
            }

            if (!($app -and $app.DisplayName -eq $permission.EntraAppName)) {
                $app = Get-EntraServicePrincipal -Filter "DisplayName eq '$($permission.EntraAppName)'"
            }

            $exoServicePrincipal = New-ServicePrincipal -AppId $app.AppId -ObjectId $app.Id -DisplayName $permission.EntraAppName
            $mailbox = Add-MailboxPermission -AccessRights $permission.AccessRights -Identity $permission.UserPrincipalName -User $exoServicePrincipal.ObjectId
            if (!$mailbox) {
                Write-Error "Failed to add mailbox permission for $($permission.UserPrincipalName)"
                continue
            }

            Write-Host "✓ Added mailbox permission for $($permission.UserPrincipalName) with access rights $($permission.AccessRights) for $($app.DisplayName)" -ForegroundColor Green

            # Update credentials file with Exchange Online UPNs
            if (Test-Path $CredentialOutputPath) {
                $existingCredentials = Get-Content $CredentialOutputPath | ConvertFrom-Json
                
                # Initialize ExchangeOnlineUPNs array if it doesn't exist
                if (-not $existingCredentials.ExchangeOnlineUPNs) {
                    $existingCredentials | Add-Member -MemberType NoteProperty -Name "ExchangeOnlineUPNs" -Value @()
                }
                
                # Add the UPN if it's not already in the array
                if ($existingCredentials.ExchangeOnlineUPNs -notcontains $permission.UserPrincipalName) {
                    $existingCredentials.ExchangeOnlineUPNs += $permission.UserPrincipalName
                }
                
                # Save updated credentials
                $existingCredentials | ConvertTo-Json | Out-File -FilePath $CredentialOutputPath -Encoding UTF8
                Write-Host "✓ Updated credentials file with UPN: $($permission.UserPrincipalName)" -ForegroundColor Green
            }
        }
    }
    
    Write-Host "✓ Script completed successfully!" -ForegroundColor Green
}
catch {
    Write-Error "Script failed: $_"
    exit 1
}
finally {
    # Disconnect from Microsoft Entra
    try {
        Disconnect-Entra -ErrorAction SilentlyContinue | Out-Null
        Write-Host "Disconnected from Microsoft Graph" -ForegroundColor Yellow
        Disconnect-ExchangeOnline -Confirm:$false | Out-Null
        Write-Host "Disconnected from Microsoft Exchange Online" -ForegroundColor Yellow
    }
    catch {
        # Ignore disconnection errors
    }
}