#you will need to change the location where you load the Auth module from.

function Get-olMemberOf{
    [cmdletbinding()]
    param(
        $Name,
        $ObjectID
    )
    
    if(-not $ObjectID){
        #Get ObjectID if not provided
        $uri = "https://graph.microsoft.com/beta/$Global:Tenant/users?`$filter=displayname eq '$Name' or userprincipalname eq '$Name'"
        $ObjectID = (Invoke-RestMethod -Uri $uri –Headers $Global:Header –Method Get).value | Select-Object -ExpandProperty id
    }

    if($ObjectID){
        $uri = "https://graph.microsoft.com/beta/$Global:Tenant/users/$ObjectID/memberOf"
        $MemberOf = (Invoke-RestMethod -Uri $uri –Headers $Global:Header –Method Get).value | Select-Object displayName,objectID,onPremisesSecurityIdentifier,dirSyncEnabled,mail,description,@{n='Child';e={$Name}}
        return $MemberOf
    }else{
        return
    }   
}

function Get-olGroupMemberShip{
    param(
        $ObjectID,
        $Name
    )

    if(-not $ObjectID){
        #Get ObjectID if not provided
        Test-GraphTokenExpiration
        $uri = "https://graph.microsoft.com/beta/$Global:Tenant/groups?`$filter=displayname eq '$Name' or mail eq '$Name' or mailNickname eq '$Name'"
        $ObjectID = (Invoke-RestMethod -Uri $uri –Headers $Global:Header –Method Get).value | Select-Object -ExpandProperty id
    }

    $MemberResults = [System.Collections.Generic.List[PSObject]]::new()
    $uri = "https://graph.microsoft.com/beta/$Global:Tenant/groups/$($ObjectID)/members"

    Do{
        Test-GraphTokenExpiration
        $Result = Invoke-RestMethod -Uri $uri -Headers $Global:Header -Method Get
        #$Result
        foreach($User in $Result.value){
            $MemberResults.Add($User)
        }
        $uri = $Result.'@odata.nextLink'
    }while($uri)
    
    return $MemberResults
}

function Get-olNestGroupMembership {
    param(
        $GroupEmail
    )

    $MemberEmail = [System.Collections.Generic.List[string]]::new()
    $GroupID = (Get-olGroup -Filter "mail eq '$GroupEmail'").id

    foreach ($object in (Get-olGroupMemberShip -ObjectID $GroupID)) {
        if ($object.'@odata.type' -match '\.user') {
            $MemberEmail.Add($object.mail)
        }
        else {
            $NestedEmail = Get-olNestGroupMembership -GroupEmail $Object.mail
            foreach ($email in $NestedEmail) {
                $MemberEmail.Add($email)
            }
        }
    }

    return $MemberEmail
}

function Get-olUser{
    param(
        $UserPrincipalName,
        $Filter,
        [switch]$All,
        [switch]$ExportNow
    )
    Test-GraphTokenExpiration

    if($UserPrincipalName){
        $uri = "https://graph.microsoft.com/beta/$Global:Tenant/users/$($UserPrincipalName)"
        return (Invoke-RestMethod -Uri $uri –Headers $Global:Header –Method Get)
    }
    if($Filter){
        $FilterResults = @()
        $uri = "https://graph.microsoft.com/beta/$Global:Tenant/users/?`$filter=$Filter"
        do{
            Test-GraphTokenExpiration
            $Result = Invoke-RestMethod -Uri $uri -Headers $Global:Header -Method Get
            if($ExportNow){
                $Result.value 
            }else{
                $FilterResults += $Result.value
            }
            $uri = $Result.'@odata.nextLink'
        }while($uri)
        if($ExportNow){
            return
        }else{
            return $FilterResults
        }
    }
    if($All){
         $FilterResults = @()
        $uri = "https://graph.microsoft.com/beta/$Global:Tenant/users"
        do{
            Test-GraphTokenExpiration
            $Result = Invoke-RestMethod -Uri $uri -Headers $Global:Header -Method Get
            if($ExportNow){
                $Result.value 
            }else{
                $FilterResults += $Result.value
            }
            $uri = $Result.'@odata.nextLink'
        }while($uri)
        if($ExportNow){
            return
        }else{
            return $FilterResults
        }
    }    
}

function Get-olGroup{
    param(
        $ObjectID,
        $Filter,
        [switch]$All,
        [switch]$ExportNow
    )
    Test-GraphTokenExpiration

    if($ObjectID){
        $uri = "https://graph.microsoft.com/beta/$Global:Tenant/groups/$($ObjectID)"
        return (Invoke-RestMethod -Uri $uri –Headers $Global:Header –Method Get)
    }
    if($Filter){
        $FilterResults = @()
        $uri = "https://graph.microsoft.com/beta/$Global:Tenant/groups?`$filter=$Filter"
        do{
            Test-GraphTokenExpiration
            $Result = Invoke-RestMethod -Uri $uri -Headers $Global:Header -Method Get
            if($ExportNow){
                $Result.value 
            }else{
                $FilterResults += $Result.value
            }
            $uri = $Result.'@odata.nextLink'
        }while($uri)
        if($ExportNow){
            return
        }else{
            return $FilterResults
        }
    }
    if($All){
         $FilterResults = @()
        $uri = "https://graph.microsoft.com/beta/$Global:Tenant/groups"
        do{
            Test-GraphTokenExpiration
            $Result = Invoke-RestMethod -Uri $uri -Headers $Global:Header -Method Get
            if($ExportNow){
                $Result.value 
            }else{
                $FilterResults += $Result.value
            }
            $uri = $Result.'@odata.nextLink'
        }while($uri)
        if($ExportNow){
            return
        }else{
            return $FilterResults
        }
    }    
}

function Test-olEmailAddress{
    param(
        $EmailAddress,
        [switch]$ExportNow
    )

    Test-GraphTokenExpiration
    $Filter = "proxyaddresses/any(c:c+eq+'smtp:$EmailAddress')"

    $FilterResults = @()
    $ObjectType = "users","groups"
    foreach($Type in $ObjectType){
        $uri = "https://graph.microsoft.com/beta/$Global:Tenant/$($Type)?`$filter=$Filter"
        do{
            Test-GraphTokenExpiration
            $Result = Invoke-RestMethod -Uri $uri -Headers $Global:Header -Method Get
            if($ExportNow){
                $Result.value 
            }else{
                $FilterResults += $Result.value
            }
            $uri = $Result.'@odata.nextLink'
        }while($uri)
    }

    if($ExportNow){
        return
    }else{
        return $FilterResults
    }
}

function Get-olContact{
    param(
        $Name,
        $Filter,
        $ObjectID,
        [switch]$All,
        [switch]$ExportNow
    )
    Test-GraphTokenExpiration

    if($Name){
        $uri = "https://graph.microsoft.com/beta/$Global:Tenant/contacts/?`$filter=displayname like '$Name' or emailaddress -eq '$Name'"
        return (Invoke-RestMethod -Uri $uri –Headers $Global:Header –Method Get)
    }
    if($Filter){
        $FilterResults = @()
        $uri = "https://graph.microsoft.com/beta/$Global:Tenant/contacts/?`$filter=$Filter"
        do{
            Test-GraphTokenExpiration
            $Result = Invoke-RestMethod -Uri $uri -Headers $Global:Header -Method Get
            if($ExportNow){
                $Result.value 
            }else{
                $FilterResults += $Result.value
            }
            $uri = $Result.'@odata.nextLink'
        }while($uri)
        if($ExportNow){
            return
        }else{
            return $FilterResults
        }
    }
     if($ObjectID){
        $uri = "https://graph.microsoft.com/beta/$Global:Tenant/contacts/$($ObjectID)"
        return (Invoke-RestMethod -Uri $uri –Headers $Global:Header –Method Get)
    }
    if($All){
         $FilterResults = @()
        $uri = "https://graph.microsoft.com/beta/$Global:Tenant/contacts"
        do{
            Test-GraphTokenExpiration
            $Result = Invoke-RestMethod -Uri $uri -Headers $Global:Header -Method Get
            if($ExportNow){
                $Result.value 
            }else{
                $FilterResults += $Result.value
            }
            $uri = $Result.'@odata.nextLink'
        }while($uri)
        if($ExportNow){
            return
        }else{
            return $FilterResults
        }
    }    
}

function Get-olSPSite{
    param{
        $SiteName # i.e. companyname.sharepoint.com
    }

    $uri = "https://graph.microsoft.com/beta/$Global:Tenant/sites/$($SiteName):/"
    return (Invoke-RestMethod -Uri $uri –Headers $Global:Header –Method Get)
}

