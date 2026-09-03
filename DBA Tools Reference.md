**DbaTools Reference**
[Quiz](#Quiz)

Set-PsRepository -Name PSGallery  
Install-Moduel dbaTools -Scope CurrentUser  
Import-Moduel dbaTools  
  
Find-DbaCommand -Pattern "database"  
Find-DbaCommand -Tag "Migration"  
Get-Help Test-DbaLastBackup -Examples  
  
Get-DbaDatabase -SqlInstance sqlserver1, sqlserver2 -Database MyDB  
Get-DbaDiskSpace -ComputerName sqlserver1  
Get-DbaProcess -SqlInstnace sqlserver1  
Get-DbaLastBackup -SqlInstance sqlserver1  
Get-DbaLastGoodCheckDb -SqlInstance sqlserver1  
  
Backup-DbaDatabaes -SqlInstance abc -ExludeSystem -Compress -BackupDirectory "\\NAS\abc"  
Test-DbaLastBackup -SqlInstance abc -Databaes HR  
  
Invoke-DbaQuery -SqlInstnace sqlserver1 -Query "SELECT @@SERVERNAME"  
  
Copy-DbaLogin -Source abc -Destination def  
Start-DbaMigration -Source abc -Destination def -SharedPath "NAS\Backup" -WhatIf  

##quiz[quiz]
1. 
2. 
3. 
4.




