<#
=======================================================================================
File Name: get-rptinactiveadaccounts.ps1
Created on: 
Created with VSCode
Version 1.0
Last Updated: 
Last Updated by: John Shelton | c: 260-410-1200 | e: john.shelton@lucky13solutions.com

Purpose:

Notes: 

Change Log:


=======================================================================================
#>
#
# Define Parameter(s)
#
param (
  [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
  [string] $InactiveDays = $(throw "-InactiveDays is required."),
  [string[]] $OUs = $(throw "-OU(s) are required")
)
#
#
# Configure HTML Header
#
$HTMLHead = "<style>"
$HTMLHead += "BODY{background-color:white;}"
$HTMLHead += "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"
$HTMLHead += "TH{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:royalblue}"
$HTMLHead += "TD{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:gainsboro}"
$HTMLHead += "</style>"
#
#
# Define Output Variables
#
$ExecutionStamp = Get-Date -Format yyyyMMdd_hh-mm-ss
$path = "c:\temp\"
$FilenamePrepend = 'temp_'
$FullFilename = "get-rptinactiveadaccounts.ps1"
$FileName = $FullFilename.Substring(0, $FullFilename.LastIndexOf('.'))
$FileExt = '.xlsx'
$OutputFile = $path + $FilenamePrePend + '_' + $FileName + '_' + $ExecutionStamp + $FileExt
#
$PathExists = Test-Path $path
IF($PathExists -eq $False)
  {
  New-Item -Path $path -ItemType  Directory
  }
#
ForEach($OU in $OUs){
  $OUShortStringTemp = $OU | Select-String "(?:CN|OU|DC)=.*?(?=(?<!\\),)|DC=.*$" | Select-Object -ExpandProperty Matches | Select-Object -ExpandProperty Value
  $OUShortName = $OUShortStringTemp.Substring(3)
  $OUShortName = $OUShortName -replace '\s','_'
  $InactiveUsersDetail = @()
  $InactiveUsers = @()
  $Count++
  $TimeSpan = [timespan]::FromDays($InactiveDays)
  $InactiveUsers += Write-Host $OU | Search-ADAccount -AccountInactive -TimeSpan $TimeSpan -UsersOnly -SearchBase $OU | Where-Object {$_.LastLogonDate -and $_.Enabled}
  ForEach($InactiveUser in $InactiveUsers){
    $InactiveUsersDetail += Get-ADUser $InactiveUser.SamAccountName -Properties * | Select CanonicalName, Created, Department, Description, DisplayName, DistinguishedName, Division, EmployeeID, Enabled, GivenName, SurName, LastLogonDate, PasswordExpired, SamAccountName
  }
  $InactiveUsersDetail | Sort-Object LastLogonDate | Export-Excel -Path $OutputFile -WorkSheetname "$OUShortName" -TableName "InactiveUsersTbl$OUShortName" -TableStyle Custom -AutoSize
}
