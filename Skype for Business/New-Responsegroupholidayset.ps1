<#

.SYNOPSIS
This powershell script uses the AUS goverement data API to collect information about Public Holidays then convert them into a usable format for Skype for Business Response Groups

.DESCRIPTION
The script will take input on what State, Front End Server, and whether to include National and output it to use on Skype for Business. This script does not require anything other than internet access

.EXAMPLE
./New-Responsegroupholidayset.ps1 -state WA -feserver skypefe1.contonso.local -includenat:$true
./New-Responsegroupholidayset.ps1 -state NSW -feserver skypefe1.contonso.local -includenat:$false

.NOTES
Created by James Rogers

01/11/2017 - Updated dataset, and paramaters

#>

[CmdletBinding()] 
param( 
  [Parameter(Mandatory=$true)] 
  [ValidateSet("WA","NSW","VIC","QLD","NT","ACT","SA","TAS")]
  [string]$state,
  
  [Parameter(Mandatory=$true)]
  [string]$feserver,

  [Parameter(Mandatory=$false)] 
  [switch]$includenat = $true
)

#---------------------------------------------------------[Initialisations]--------------------------------------------------------
function Convert-DateString ([string]$Date,[string[]]$Format) {
  $result = New-Object DateTime
  $convertible = [datetime]::TryParseExact(
    $Date,
    $Format,
    [System.Globalization.CultureInfo]::InvariantCulture,
    [System.Globalization.DateTimeStyles]::None,
    [ref]$result)
  if ($convertible) { $result }
}

#----------------------------------------------------------[Declarations]----------------------------------------------------------
#ensure variables are blank
$i = 1
$year = @()

#get data from govement
#for up to date info on the API go to; https://data.gov.au/dataset/australian-holidays-machine-readable-dataset
$request = 'http://data.gov.au/api/3/action/datastore_search?resource_id=a24ecaf2-044a-4e66-989c-eacc81ded62f'

#---------------------------------------------------------[Main Execution]---------------------------------------------------------

Write-Host -ForegroundColor Green -BackgroundColor Black "Building Holidays for" $state `n
Write-Host -ForegroundColor Green -BackgroundColor Black "On the Front end Server" $feserver `n
Write-Host -ForegroundColor Green -BackgroundColor Black "Copy and paste this script and run on the front end server" `n

#hit gov api get databack as json and convert
$json = Invoke-WebRequest $request | ConvertFrom-Json

#try get data applicable to specified state, and national if included.
if ($includenat) {
  $outvalue = $json.result.records | Where-Object { $_."Applicable To" -match $state -or $_."Applicable To" -match 'NAT' } | Select-Object Date,"Holiday Name"
}
else {
  $outvalue = $json.result.records | Where-Object { $_."Applicable To" -match $state } | Select-Object Date,"Holiday Name"
}

#write read if nothing is returned
if (($outvalue).count -lt 1) {
  Write-Error "No Data Returned"
  Write-Error "Data returned $json"
  Write-Error "Data selected $outvalue"
}

else {
  #format each holiday
  $outvalue | ForEach-Object {
    $number = $i
    $i++
    $date = Convert-DateString -Date $_.Date -Format yyyyMMdd
    $name = $_."Holiday Name"
    $startdate = (Get-Date $date -Format "dd/MM/yyyy hh:mm tt")
    $startdateyy = (Get-Date $date -Format yyyy)
    $stopdate = (Get-Date ((Get-Date $date).AddDays(1)) -Format "dd/MM/yyyy hh:mm tt")
    "$" + $number + " = New-CsRgsHoliday -Name `"" + $name + " " + $startdateyy + "`" -StartDate `"" + $startdate + "`" -EndDate `"" + $stopdate + "`""
    $year += $startdateyy
  }
  #output group
  $a = 1..($outvalue).count
  $holidaygroup = [System.String]::Join(", $",$a)
  $holidaygroup = "$" + $holidaygroup

  "New-CsRgsHolidaySet -Parent `"service:ApplicationServer:" + $feserver + "`" -Name `"Holidays for " + $state + "`" -HolidayList(" + $holidaygroup + ")"

}
