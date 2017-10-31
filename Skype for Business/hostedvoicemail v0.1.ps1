#requires -version 2
<#
.SYNOPSIS
  Configuration of Hosted voicemail for office 365 users on Skype for Business
.DESCRIPTION
  This script searches active directory for users with the attribute ExchangeHostedVoiceMail=1. Then enables them for Hosted voicemail within Skype for Business
.PARAMETER <Parameter_Name>
    None
.INPUTS
  None
.OUTPUTS
  Log file stored in "execute.log"
.NOTES
  Version:        0.1
  Author:         James Rogers
  Creation Date:  12/10/2017
  Purpose/Change: Initial script development
#>

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

#bring in functions
."functions\writelog.ps1"
$logpath = "execute.log"

#----------------------------------------------------------[Declarations]----------------------------------------------------------

#declare policy name
$voicemailpolicy = "Hosted Voicemail Polciy Name"

#declare server to run remote command
$DCServer = "Domain Controller"


#---------------------------------------------------------[Main Execution]---------------------------------------------------------

#log script started
Write-Log -Path $logpath -Level Info -Message "--------------------Script started--------------------"

#return users based on command run on dc to $users variable, user must be enabled and have the msExchUCVoicemailSettings attribute
$users = Get-ADUser -LdapFilter "(&(objectCategory=person)(objectClass=user)(!userAccountControl:1.2.840.113556.1.4.803:=2)(msExchUCVoiceMailSettings=ExchangeHostedVoiceMail=1))" -Server $DCServer  |  Select-Object UserPrincipalName,SamAccountName


#loop of users returned from AD
foreach ($user in $users) {
  $sam = $user.SamAccountName
  $upn = $user.UserPrincipalName
  #try get the user first
  try {
    $csuser = Get-CsUser -Identity $upn
  }
  #catch if user doesnt exist, skip rest of loop for that user
  catch {
    Write-Log -Path $logpath -Level Info -Message "$sam not found doing nothing"
    return
  }
  #check if users is enabled for EV if not do nothing
  if ($csuser.EnterpriseVoiceEnabled -eq $false) {
    Write-Log -Path $logpath -Level Info -Message "$sam not enabled for ev doing nothing"
  }
  else {
    Write-Log -Path $logpath -Level Info -Message "$sam enabled for ev attempting to set policies"

    # Check if user has correct hosted voicemail policy, if not assign
    if ($csuser.HostedVoicemailPolicy.FriendlyName -ne $voicemailpolicy) {
      Grant-CsHostedVoicemailPolicy -Identity $csuser.Identity -PolicyName $voicemailpolicy
      Write-Log -Path $logpath -Level Info -Message "$sam granted voicemail policy"
    }
    else {
      Write-Log -Path $logpath -Level Info -Message "$sam has correct policy not changing"
    }

    # Check if user is enabled for hosted voicemail, if not enable
    if ($csuser.HostedVoiceMail -ne $true) {
      Set-CsUser -Identity $csuser.Identity -HostedVoiceMail $true
      Write-Log -Path $logpath -Level Info -Message "$sam enabled hosted voicemail"
    }
    else {
      Write-Log -Path $logpath -Level Info -Message "$sam already enabled for hosted voicemail"
    }
  }
}

#log script started
Write-Log -Path $logpath -Level Info -Message "--------------------Script ended--------------------"
