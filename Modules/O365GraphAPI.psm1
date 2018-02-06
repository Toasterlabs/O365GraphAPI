Function Get-GraphAPIAuthToken{
    Param(
    
        [parameter(Mandatory=$false)]
        [STRING]$resource = "https://graph.microsoft.com",
        [parameter(Mandatory=$true)]
        [STRING]$ClientID,
        [parameter(Mandatory=$false)]
        [STRING]$loginURL= "https://login.microsoftonline.com",
        [parameter(Mandatory=$true)]
        [STRING]$tenantdomain,
        [parameter(Mandatory=$true)]
        [System.Management.Automation.PSCredential]$credential,
        [parameter(Mandatory=$false)]
        [STRING]$redirectUri = "urn:ietf:wg:oauth:2.0:oob"                          
    )

    # Need the Azure module for this one to work
    Import-Module Azure
    
    # Forming our authority url
    $authority = "$loginURL/$tenantdomain"


    $authContext = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext" -ArgumentList $authority
    $AADCredential = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.UserCredential" -ArgumentList $credential.UserName,$credential.Password
    
    # Get the token!
    $authResult = $authContext.AcquireToken($resource, $clientId,$AADCredential)
    
    # Return the token!
    Return $authResult

} # Token Acquired...

Function Get-GlobalConfig{
    Param(
        [parameter(Mandatory=$true)]
        [STRING]$configFile
    )

    Write-Output "Loading Global Config File"  

    $config = Get-Content $configFile -Raw | ConvertFrom-Json

    # Returning the configuration for later use
    return $config;

} # Configuration loaded!

function get-GraphReport{
    Param(
        [parameter(Mandatory=$true)]
        [validateScript({If(Test-Path $_){$true}Else{Throw "Invalid configuration file path: $_"}})]
        $ConfigFile,
        
        [parameter(Mandatory=$true)]
        [System.Management.Automation.PSCredential]$Credential,

        [Parameter(Mandatory=$true)]
        [ValidateSet("Exchange","Groups","OneDrive","SharePoint","Skype","Tenant","Yammer")]
        $WorkLoad,  

        [Parameter(Mandatory=$false)]
        [ValidateSet("D7","D30","D90","D180")]
        $Period,

        [Parameter(Mandatory=$false)]
        $Date
    )

    DynamicParam {
            # Set the dynamic parameters' name
            $ParameterName = 'ReportType'
            
            # Create the dictionary 
            $RuntimeParameterDictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary

            # Create the collection of attributes
            $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
            
            # Create and set the parameters' attributes
            $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute
            $ParameterAttribute.Mandatory = $true
            $ParameterAttribute.Position = 1
            $ParameterAttribute.ParameterSetName = '__AllParameterSets'

            # Add the attributes to the attributes collection
            $AttributeCollection.Add($ParameterAttribute)

            # Generate and set the ValidateSet
            If ($Workload -eq "Exchange"){$arrSet = @("EmailActivity","getEmailActivityUserDetail","getEmailActivityCounts","getEmailActivityUserCounts","getEmailAppUsageUserDetail","getEmailAppUsageAppsUserCounts","getEmailAppUsageUserCounts","getEmailAppUsageVersionsUserCounts","getMailboxUsageDetail","getMailboxUsageMailboxCounts","getMailboxUsageQuotaMailboxStatusCounts","getMailboxUsageStorage")}
            If ($Workload -eq "Groups"){$arrSet = @("getOffice365GroupsActivityDetail","getOffice365GroupsActivityCounts","getOffice365GroupsActivityGroupCounts","getOffice365GroupsActivityStorage","getOffice365GroupsActivityFileCounts")}
            If ($Workload -eq "OneDrive"){$arrSet = @("getOneDriveActivityUserDetail","getOneDriveActivityUserCounts","getOneDriveActivityFileCounts","getOneDriveUsageAccountDetail","getOneDriveUsageAccountCounts","getOneDriveUsageFileCounts","getOneDriveUsageStorage")}
            If ($Workload -eq "SharePoint"){$arrSet = @("getSharePointActivityUserDetail","getSharePointActivityFileCounts","getSharePointActivityUserCounts","getSharePointActivityPages","getSharePointSiteUsageDetail","getSharePointSiteUsageFileCounts","getSharePointSiteUsageSiteCounts","getSharePointSiteUsageStorage","getSharePointSiteUsagePages")}
            If ($Workload -eq "Skype"){$arrSet = @("getSkypeForBusinessActivityUserDetail","getSkypeForBusinessActivityCounts","getSkypeForBusinessActivityUserCounts","getSkypeForBusinessDeviceUsageUserDetail","getSkypeForBusinessDeviceUsageDistributionUserCounts","getSkypeForBusinessDeviceUsageUserCounts","getSkypeForBusinessOrganizerActivityCounts","getSkypeForBusinessOrganizerActivityUserCounts","getSkypeForBusinessOrganizerActivityMinuteCounts","getSkypeForBusinessParticipantActivityCounts","getSkypeForBusinessParticipantActivityUserCounts","getSkypeForBusinessParticipantActivityMinuteCounts","getSkypeForBusinessPeerToPeerActivityCounts","getSkypeForBusinessPeerToPeerActivityUserCounts","getSkypeForBusinessPeerToPeerActivityMinuteCounts")}
            If ($Workload -eq "Tenant"){$arrSet = @("getOffice365ActivationsUserDetail","getOffice365ActivationCounts","getOffice365ActivationsUserCounts","getOffice365ActiveUserDetail","getOffice365ActiveUserCounts","getOffice365ServicesUserCounts")}
            If ($Workload -eq "Yammer"){$arrSet = @("getYammerActivityUserDetail","getYammerActivityCounts","getYammerActivityUserCounts","getYammerDeviceUsageUserDetail","getYammerDeviceUsageDistributionUserCounts","getYammerDeviceUsageUserCounts","getYammerGroupsActivityDetail","getYammerGroupsActivityGroupCounts","getYammerGroupsActivityCounts")}
            

            #$arrSet = Get-ChildItem -Path .\ -Directory | Select-Object -ExpandProperty FullName
            $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)

            # Add the ValidateSet to the attributes collection
            $AttributeCollection.Add($ValidateSetAttribute)

            # Create and return the dynamic parameter
            $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ParameterName, [string], $AttributeCollection)
            $RuntimeParameterDictionary.Add($ParameterName, $RuntimeParameter)
            return $RuntimeParameterDictionary
    }
    
    # Do this first
    Begin {
        # Bind the parameter to a friendly variable
        $Report = $PsBoundParameters[$ParameterName]
    }

    # Process is required when using dynamic parameters
    Process{
        #If the credential object is empty, prompt the user for credentials
        if(!$Credential) {$Credential = Get-Credential}

        # Retrieving settings
        $globalconfig = Get-GlobalConfig -configFile $ConfigFile

        # Setting us up to get the token
        $ClientID = $globalConfig.AppId
        $redirectURI = $globalconfig.redirectURI
        $loginURL = $globalConfig.LoginURL
        $tenantdomain = $globalConfig.TenantDomain
        $resource = $globalConfig.ResourceAPI

        # Retrieving our Token
        $Token = Get-GraphAPIAuthToken -resource $resource -ClientID $ClientID -loginURL $loginURL -tenantdomain $tenantdomain -credential $Credential -redirectUri $redirectURI

        #Build REST API header with authorization token
        $authHeader = @{
           'Content-Type'='application\json'
           'Authorization'=$token.CreateAuthorizationHeader()
        }

        #If period is specified then add that to the parameters unless it is not supported
        if($period -and $Report -notlike "*Office365Activation*"){
            $str = "period='{0}'," -f $Period
            $parameterset += $str
        }
    
        #If the date is specified then add that to the parameters unless it is not supported
        if($date -and !($report -eq "MailboxUsage" -or $report -notlike "*Office365Activation*" -or $report -notlike "*getSkypeForBusinessOrganizerActivity*")){
            $str = "date='{0}'" -f $Date
            $parameterset += $str
        }
        #Trim a trailing comma off the ParameterSet
        if($parameterset){
            $parameterset = $parameterset.TrimEnd(",")
        }

        #Building our URI 
        $uri = "https://graph.microsoft.com/beta/reports/{0}({1})" -f $report, $parameterset

        # Capturing the result our our query to the graph API in a variable
        $result = Invoke-RestMethod -Uri $uri –Headers $authHeader –Method Get

        # Since this result is returned in annoying CSV format by default (and the .content is useless), we'll need to convert this...
        $resultarray = ConvertFrom-Csv -InputObject $result
    }

    # Do this last
    End{
        # Return the results of our query in a proper format we can use for further processing
        return $resultarray
    }
}