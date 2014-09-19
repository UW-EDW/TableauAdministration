#****************************************************************
#
# Script Name: sync-ad.ps1
# Version: 1.0
# Author: Travis Boyle
# Date: 9/19/2014
#
# Description: script that can be scheduled to run that will sync a list of Active Directory groups with 
#   corresponding Tableau Server groups.  
#
# Usage: first run, run from Powershell command line, to create credential file.  After first run, can be
#   scheduled
# 		.\sync-ad.ps1 tableauAdministratorUsername
#     powershell.exe -file <path>\ad-sync.ps1 -TabCmdUser <username>
#*****************************************************************

[CmdletBinding()]
Param(
   [Parameter(Mandatory=$true,Position=1)]
      [string]$TabCmdUser,
   [string]$pass   
)


$WorkingDirectory = ""; # Update with your working directory path
$CredsFile = "$WorkingDirectory\$TabCmdUser.txt"; # Filename to store credential
$adGroupListFile = "$WorkingDirectory\AD_Group_Sync_List.txt";
$Domain = ""; # Update with AD Domain
$datetime = get-date;
$logdate = get-date -format "yyyyMMdd"
$logFile = "$WorkingDirectory\logs\nightly-adsync.$logDate.txt";
$Server = ""; # Update with URL of Tableau Server - for use in tabcmd  

# attempts to add path to tabcmd - if tabcmd path is not available, either add directly to path, or update function
function set-path {
   if (!($env:path.Contains("Tableau Server"))){
      if (test-path "C:\Program Files\Tableau\Tableau Server\8.2\extras\Command Line Utility\tabcmd.exe"){
         $env:path = $env:path + ";C:\Program Files\Tableau\Tableau Server\8.2\extras\Command Line Utility\"
      }
      elseif (test-path "C:\Program Files\Tableau\Tableau Server\8.2\bin\tabcmd.exe"){
         $env:path = $env:path + ";C:\Program Files\Tableau\Tableau Server\8.2\bin\"
      }
      elseif (test-path "D:\Program Files\Tableau\Tableau Server\8.2\bin\tabcmd.exe"){
         $env:path = $env:path + ";D:\Program Files\Tableau\Tableau Server\8.2\bin\"
      }
      elseif (test-path "C:\Program Files (x86)\Tableau\Tableau Server\8.2\bin\tabcmd.exe"){
         $env:path = $env:path + ";C:\Program Files (x86)\Tableau\Tableau Server\8.2\bin\"
      }
      elseif (test-path "D:\Program Files (x86)\Tableau\Tableau Server\8.2\bin\tabcmd.exe"){
         $env:path = $env:path + ";D:\Program Files (x86)\Tableau\Tableau Server\8.2\bin\"
      }
   }
}

function build-credentials{
   #STORED CREDENTIAL CODE
   #$CredsFile = "C:\working\PScripts\pass\$TabCmdUser.txt"
   $FileExists = Test-Path $CredsFile
   if  ($FileExists -eq $false) {
      Write-Host 'Credential file not found. Enter your password:' -ForegroundColor Red
      Read-Host -AsSecureString | ConvertFrom-SecureString | Out-File $CredsFile
      $password = get-content $CredsFile | convertto-securestring
   }else{
      Write-Host 'Using your stored credential file' -ForegroundColor Green
      $password = get-content $CredsFile | convertto-securestring
   }
   $Script:Cred = new-object -typename System.Management.Automation.PSCredential -argumentlist $Domain\$TabCmdUser,$password
}

function build-tabcmdlist{
   $Script:tabcmdlist = @();
   $synclist = get-content $adGroupListFile;
   [string]$Script:tabuser = $Script:Cred.username;
   $Script:tabpass = ([Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($Script:Cred.Password)));
   $Script:tabcmdlist += " login -s $Script:Server -u ""$tabuser"" -p ""$tabpass""";
   $tabpass = $null;
   #$Script:tabcmdlist += " createProject -n $ProjectName -d $Description";
   for($i=0;$i -lt $synclist.length; $i++){
      $Script:tabcmdlist += " syncgroup " + $synclist[$i] + " --license interactor --no-complete";
   }
   $Script:tabcmdlist += " logout";
}

function process-tabcmd{
   for($i=0;$i -lt $tabcmdlist.length; $i++){
      $command = "tabcmd.exe" + $tabcmdlist[$i];
      if($i -gt 0){$command | out-file -filepath $logFile -append;}
      Invoke-Expression "$command" | out-file -filepath $logFile -append;
   }
}

set-path;
build-credentials;
build-tabcmdlist;
process-tabcmd;
