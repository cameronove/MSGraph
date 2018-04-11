# Defining Azure AD tenant name, this is the name of your Azure Active Directory (do not use the verified domain name)
$Global:Tenant = "<tenant.url>"
$Global:ClientID = "1950a258-227b-4e31-a9cf-717495945fc2" #well-known clientid for Azure

#some arbitrary path to keep your JSON token
# PSCD = PowerShell Configuration Directory
$PSCD = "$env:USERPROFILE\Documents\WindowsPowerShell\Config"
if(-not (Test-Path $PSCD)){
    mkdir $PSCD   
}

#the name of the file to store you oauth token
$Global:FileTokenPath = "$PSCD\MSGraphToken.json"


<#----------Acquire Token Functions----------#>
function Register-GraphAuthToken{

    $RedirectUri = "urn:ietf:wg:oauth:2.0:oob"
    $ResourceAppIdURI = "https://graph.microsoft.com"
    $Authority = "https://login.microsoftonline.com/$Global:Tenant"
    $AuthContext = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext" -ArgumentList $Authority
    $AuthResult = $authContext.AcquireToken($ResourceAppIdURI, $Global:ClientId, $RedirectUri, "Always")
    #Set global header for functions to use
    $Global:Header = @{
        'Content-Type'='application\json'
        'Authorization'=$AuthResult.CreateAuthorizationHeader()
    }

    #Set global token for functions to check
    $Global:Token = $AuthResult

    #Write new token to global token file
    ConvertTo-Json $Global:Token | Out-File -Encoding utf8 -FilePath $Global:FileTokenPath
}

function Update-GraphAuthTokenFromFile{
    
    #if file does not exist or exists but is not token data then register token
    if(-not (Test-Path $Global:FileTokenPath)){
        Register-GraphAuthToken
        return
    }elseif(([io.file]::ReadAllText($Global:FileTokenPath)) -notmatch 'AccessToken'){
        Register-GraphAuthToken
        return
    }
    $FileToken = ConvertFrom-Json -InputObject ([io.file]::ReadAllText($Global:FileTokenPath)) 
    if(((Get-Date) - (Get-Date $FileToken.ExpiresOn)) -gt 14){
        Register-GraphAuthToken
        return
    }
    $Authority = "https://login.microsoftonline.com/$Global:Tenant"
    Write-Host $Authority
    $AuthContext = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext" -ArgumentList $Authority
    $AuthResult = $AuthContext.AcquireTokenByRefreshToken($FileToken.RefreshToken, $Global:ClientId)

    #Set global header for functions to use
    $Global:Header = @{
        'Content-Type'='application\json'
        'Authorization'=$AuthResult.CreateAuthorizationHeader()
    }

    #Set global token for functions to check
    $Global:Token = $AuthResult

    Test-GraphTokenExpiration
}

function Update-GraphAuthToken{
    
    if(-not $Global:Token){
        Update-GraphAuthTokenFromFile
        return
    }
    $Authority = "https://login.microsoftonline.com/$Global:Tenant"
    $AuthContext = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext" -ArgumentList $Authority
    $AuthResult = $AuthContext.AcquireTokenByRefreshToken($Global:Token.RefreshToken, $Global:ClientId)
    #Set global header for functions to use
    $Global:Header = @{
        'Content-Type'='application\json'
        'Authorization'=$AuthResult.CreateAuthorizationHeader()
    }

    #Set global token for functions to check
    $Global:Token = $AuthResult

    #Write new token to global token file
    ConvertTo-Json $Global:Token | Out-File -Encoding utf8 -FilePath $Global:FileTokenPath
}

function Test-GraphTokenExpiration{
    if(-not $Global:Token){
        Update-GraphAuthTokenFromFile
    }
    if($Global:Token.ExpiresOn.UtcDateTime -lt (Get-Date).ToUniversalTime()){
        Update-GraphAuthToken
    }
}

function Initialize-GraphAuthToken{
    #if token is already in memory then refresh it else try to get from file 
    #Update-GraphAuthTokenFromFile will call Register-GraphAuthToken if the file does not exist or does not have token data.
    if($Global:Token){
        Update-GraphAuthToken
    }else{
        Update-GraphAuthTokenFromFile
    }
}

<#----------Generate Global Token if not already done----------#>

Initialize-GraphAuthToken