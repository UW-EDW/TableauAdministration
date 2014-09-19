#****************************************************************
#
# Script Name: Create-ADProject.ps1
# Version: 1.0
# Author: Travis Boyle
# Date: 9/19/2014
#
# Description: Create Project and Groups, if applicable, add AD Groups to sync list
#
# Usage: .\Create-ADProject.ps1 <server> <projectname> <projectdescription> <use AD groups (t/f)> <publisher group> <editor group> <viewer group> (optional) -tabcmduser <username> -pass <password>
#*****************************************************************



[CmdletBinding()]
Param(
   [Parameter(Mandatory=$true,Position=1)]
      [string]$Server,

   [Parameter(Mandatory=$true,Position=2)]
      [string]$ProjectName,
      
   [Parameter(Mandatory=$true,Position=3)]
      [string]$Description,
      
   [Parameter(Mandatory=$true,Position=4)]
      [bool]$UseAD,
      
   [Parameter(Position=5)]
      [string]$PublisherGroup,
      
   [Parameter(Position=6)]
      [string]$EditorGroup,
      
   [Parameter(Position=7)]
      [string]$ViewerGroup,

   [string]$TabCmdUser,
   [string]$pass,
   [switch]$testing
   
)

#Configuration
$WorkingDirectory = ""; # Update with working directory path
$CredsFile = "$WorkingDirectory\$TabCmdUser.txt"; 
$adGroupListFile = "$WorkingDirectory\AD_Group_Sync_List.txt"; # List of AD groups, for use in sync-ad.ps1
$logFile = "$WorkingDirectory\logs\$ProjectName.txt";
$Domain = ""; #update with AD Domain
$datetime = get-date;

"Creating $ProjectName project on $Server, $datetime" | out-file -filepath $logFile -append

function transform-server{
   if($Server.StartsWith("https://")){
      $Script:ConStrServer = $Server.substring(8);
   }else{
      $Script:ConStrServer = $Server;
      $Script:Server = "https://"+$Server;
   }

   if($Script:ConStrServer -eq "bitools.uw.edu"){
      $Script:ConStrServer = "edwtab3.cac.washington.edu";
      $Server = "https://$Script:ConStrServer";
   }
}

if($testing){
   $Server;
   $conStrServer;
}



function build-group-names{
   if($UseAD){
      
      
      if(!$PublisherGroup){
         $Script:PublisherGroup = Read-Host " Enter Publisher AD group";
      }
      
      if(!$EditorGroup){
         $Script:EditorGroup = Read-Host " Enter Editor AD group";
      }
         
      if(!$ViewerGroup){
         $Script:ViewerGroup = Read-Host " Enter Viewer AD group";
      }
      
      $Script:PublisherGroup | out-file -filepath $adGroupListFile -append;
      $Script:EditorGroup | out-file -filepath $adGroupListFile -append;
      $Script:ViewerGroup | out-file -filepath $adGroupListFile -append;
      $Script:SetPermissionsSQL = "SELECT setpermissionsad('$projectname', '$EditorGroup', '$PublisherGroup', '$ViewerGroup')"

   }else{
      
      if(!$PublisherGroup){
         $Script:PublisherGroup = "$projectname Publishers";
         $Script:PublisherGroup = $Script:PublisherGroup -replace " ","_";
      }
      
      if(!$EditorGroup){
         $Script:EditorGroup = "$projectname Editors";
         $Script:EditorGroup = $Script:EditorGroup -replace " ","_";
      }
         
      if(!$ViewerGroup){
         $Script:ViewerGroup = "$projectname Viewers";
         $Script:ViewerGroup = $Script:ViewerGroup -replace " ","_";
      }
   $Script:SetPermissionsSQL = "SELECT setpermissions('$projectname')"
   }
}

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

   if(!$TabCmdUser){
      $TabCmdUser = Read-Host "Enter your TabCMD username"
   }

   $FileExists = Test-Path $CredsFile
   if  ($FileExists -eq $false) {
      Write-Host 'Credential file not found. Enter your password:'
      Read-Host -AsSecureString | ConvertFrom-SecureString | Out-File $CredsFile
      $password = get-content $CredsFile | convertto-securestring
   }else{
      Write-Host 'Using your stored credential file' 
      $password = get-content $CredsFile | convertto-securestring
   }
   $Script:Cred = new-object -typename System.Management.Automation.PSCredential -argumentlist $Domain\$TabCmdUser,$password
}

function build-tabcmdlist{
   $Script:tabcmdlist = @();
   [string]$Script:tabuser = $Script:Cred.username;
   $Script:tabpass = ([Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($Script:Cred.Password)));
   $Script:tabcmdlist += " login -s $Script:Server -u ""$tabuser"" -p ""$tabpass""";
   $tabpass = $null;
   $Script:tabcmdlist += " createProject -n $ProjectName -d $Description";
   if($UseAD){
      $Script:tabcmdlist += " syncgroup $PublisherGroup --license interactor --no-complete";
      $Script:tabcmdlist += " syncgroup $EditorGroup --license interactor --no-complete";
      $Script:tabcmdlist += " syncgroup $ViewerGroup --license interactor --no-complete";
   }else{
      $Script:tabcmdlist += " creategroup $PublisherGroup";
      $Script:tabcmdlist += " creategroup $EditorGroup";
      $Script:tabcmdlist += " creategroup $ViewerGroup";
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


function execute-storedproc{
   $conn = New-Object System.Data.Odbc.OdbcConnection;
   $odbcuser = $tabuser.substring($domain.length+1);
   $connStr = "Driver={PostgreSQL UNICODE(x64)}; Server=$conStrServer; Port=8060; Database=workgroup; Integrated Security = true;";
   $connStr | out-file -filepath $logFile -append;
   $conn.ConnectionString = $connStr;
   [void]$conn.open();
   $cmd = new-object System.Data.Odbc.OdbcCommand($SetPermissionsSQL,$conn);
   if($cmd.ExecuteNonQuery()){
      "Executed $SetPermissionsSQL on $conStrServer"  | out-file -filepath $logFile -append;
   }else{
      "Stored proc failed - check Postgress"  | out-file -filepath $logFile -append;
   }
   [void]$conn.close();
}

transform-server;
build-group-names; 
set-path; #ensure tabcmd is in path
build-credentials; #options to use $tabcmduser/$pass or $tabcmduser and file
build-tabcmdlist; #tabcmds to run (login / createproject / syncgroup (x3) / logout)
process-tabcmd; #run list of commands built with build-tabcmdlist
execute-storedproc; #build connection to server, run storedproc


if($testing){
   Write-Host "$ProjectName, $Description, $UseAD"
   Write-Host "$PublisherGroup, $EditorGroup, $ViewerGroup"
   write-host "$setPermissionsSQL"
   for($i=1;$i -lt $tabcmdlist.length; $i++){
         $tabcmdlist[$i];
   }
}
