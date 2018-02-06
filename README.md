# O365GraphAPI

## Description
Powershell functions to retrieve information from the O365 Graph API. Currently able to pull all reports listed on https://developer.microsoft.com/en-us/graph/docs/api-reference/beta/resources/report

Has the following functions:

### get-globalconfig
*.Param*: ConfigFile
Imports configuration from json

### Get-Graphoauthtoken
*.Param*: Resource, ClientID, LoginURL, TenantDomain, RedirectURI, Credentials
Constructs header for use with other functions. Did not make this work with an app secret, so you need to pass credentials to it. Easiest is to do this with $cred = get-credential

*.Prerequisites*
Requires the Azure PS module to be installed. Install using "install-module Azure"

### Get-GraphReport
*Inspired by the script Get-Office365Report from Damien Wiese*

*.Param*: ConfigFile, Workload, ReportType, Period, Date
Reads information from the configfile, gets the token, retrieves the report, and converts it to a useable format for further processing

*.Usage*
get-GraphReport -configfile <Path to config JSON> -Workload <Workload to report from> -ReportType <Type of report to get based on workload> -Credential <Credential to retrieve token with> -period <Period to run report against> 
get-GraphReport -configfile <Path to config JSON> -Workload <Workload to report from> -ReportType <Type of report to get based on workload> -Credential <Credential to retrieve token with> -date <Specific day to run report against>

## Configuration
Requires a graph API application to be setup in the azure management portal . Set the redirect URI to "urn:ietf:wg:oauth:2.0:oob". Not really using this, but it needs to be set...

Configuration file layout:

```
{
    "Description": "Global configuration file for O365 Investigations",
    "AppID": "<App ID goes here>",
    "TenantDomain": "<your tenant domain (contoso.onmicrosoft.com or contoso.com)>",
    "LoginURL" : "https://login.windows.net",
    "ResourceAPI" : "https://manage.office.com"
	"redirectURI" : "urn:ietf:wg:oauth:2.0:oob" 
}
```

## Usage
Install azure PS module:
install-module Azure

Import module:
Import-module <path to O365graphap.psm1> -force

Run report:
get-graphreport -configfile <Path to config JSON> -Workload <Workload to report from> -ReportType <Type of report to get based on workload> -Credential <Credential to retrieve token with> -date <Specific day to run report against>

## Disclaimer
Use this code at your own risk, verify anything you run in production. If you decide to run this code all risk stays with you. 

### Legalese:
> The sample scripts are provided AS IS without warranty of any kind. I disclaim all implied warranties including, without limitation, any implied warranties of merchantability or of fitness for a particular purpose. The entire risk arising out of the use or performance of the sample scripts and documentation remains with you. In no event shall I,or anyone else involved in the creation, production, or delivery of the scripts be liable for any damages whatsoever (including, without limitation, damages for loss of business profits, business interruption, loss of business information, or other pecuniary loss) arising out of the use of or inability to use the sample scripts or documentation, even if I have been advised of the possibility of such damages.