# Dit script verzorgt de volgende taken:
# Stap 1 - Het kopieren van de trenddatabases en Install map naar de lokale map TrendTot
# Stap 2 - Het opzetten van de Stack webdav verbinding
# Stap 3 - Het overzetten naar Stack van de lokaal klaargezette bestanden bij stap 1
# Stap 4 - Het verbreken van de Stack webdav verbinding
# Stap 5 - Vermelding in het register schrijven of de backup gelukt is of niet.

# Dit ter vervanging van de diverse losse scripts.
# Zaken tussen < > tekens worden vanuit het serverconfiguratorscript (Stap 10) ingevuld.

# Algemeen
$Time=Get-Date
echo "$Time - Backupscript Schouten Techniek gestart."

$stckur = "<Step10_stckusr>"
$stckww = "<Step10_stckpw>"

Try{
# Stap 0a - Stack verbinding afsluiten voor het geval deze na vorige sessie actief is gebleven.
C:\WINDOWS\system32\net use Z: /delete

# Stap 0b - Backup_log.txt verwijderen indien groter dan 10MB.
    $path = '<dr_install>:\Install\SysteembeheerST\Backup_log.txt'
    Get-ChildItem $path |
    Where-Object {$_.Length -gt 10Mb} | 
    Remove-Item 

# Stap 1 - Het kopieren van de trenddatabases en Install map naar de lokale map TrendTot
$Time=Get-Date
echo "$Time - Stap 1/5: Kopieren lokale bestanden"
ROBOCOPY "<dr_ipub>:\Inetpub\wwwroot\WEBVisionNT\PROJEKTE\Beheer\energietagebuch" "<dr_install>:\TrendTot\Energiedagboek" /COPYALL /S /SEC /R:50 /W:60 /LOG+:<dr_install>:\Install\SysteembeheerST\Backup_log.txt /NFL /NDL
ROBOCOPY "<dr_ipub>:\Inetpub\wwwroot\WEBVisionNT\PROJEKTE\Beheer\Trend" "<dr_install>:\TrendTot" /COPYALL /S /SEC /R:50 /W:60 /LOG+:<dr_install>:\Install\SysteembeheerST\Backup_log.txt /NFL /NDL
$localcopyfail = $lastexitcode #Robocopy is geen Powershell script, dus errorhandling via omweg

# Stap 2 - Het opzetten van de Stack webdav verbinding
$Time=Get-Date
echo "$Time - Stap 2/5: Verbinden met Stack server"
$stckurd = [System.Text.Encoding]::Unicode.GetString([System.Convert]::FromBase64String($stckur))
$stckwwd = [System.Text.Encoding]::Unicode.GetString([System.Convert]::FromBase64String($stckww))

C:\WINDOWS\system32\net use Z: "https://schoutentechniek.stackstorage.com/remote.php/webdav/Regeltechniek/Backup_WEBVision_servers/<Step10_project>_<Step10_location>" /User:$stckurd $stckwwd /y
$stackconnectfail = $lastexitcode

# Stap 3 - Het overzetten naar Stack van de lokaal klaargezette bestanden bij stap 1
$Time=Get-Date
echo "$Time - Stap 3/5: Bestanden op backupserver zetten"
ROBOCOPY2003 "<dr_install>:\TrendTot" "Z:\Backup_automatisch\TrendTot" /MIR /LOG+:<dr_install>:\Install\SysteembeheerST\Backup_log.txt /NFL /NDL
ROBOCOPY2003 "<dr_install>:\Install" "Z:\Backup_automatisch\Install" /MIR /LOG+:<dr_install>:\Install\SysteembeheerST\Backup_log.txt /NFL /NDL 
$backupcopyfail = $lastexitcode

# Stap 4 - Het verbreken van de Stack webdav verbinding
$Time=Get-Date
echo "$Time - Stap 4/5: Verbreken verbinding Stack server"
C:\WINDOWS\system32\net use Z: /delete 
$stackdisconnectfail = $lastexitcode


if(($localcopyfail -le 7) -And ($backupcopyfail -le 7) -And ($stackconnectfail -le 2) -And ($stackdisconnectfail -eq 0) ){ # Alleen succesvol als niet-Powershell scripts ook foutloos gerund zijn.
$Succes = 1
}
else{
$Succes = 0
$ErrorMessage = "Exitcode robocopy lokaal kopieren: $localcopyfail. `nExitcode net use (stackconnect): $stackconnectfail. `nExitcode robocopy backup kopieren: $backupcopyfail. `nExitcode net use (stackdisconnect): $stackdisconnectfail. `n`nZie Backup_log.txt voor aanvullende informatie of google op Robocopy/net use exit codes."
}

} 

Catch{
$Succes = 0
$ErrorMessage = $_.Exception.Message
$FailedItem = $_.Exception.ItemName
}

# Stap 5 - Vermelding in het register schrijven of de backup gelukt is of niet.
$Time=Get-Date
echo "$Time - Stap 5/5: Vermelding in register schrijven"

$logSourceExists = Test-Path "HKLM:\System\CurrentControlSet\Services\EventLog\Application\Systeembeheer ST"
if (! $logSourceExists){
New-Eventlog -LogName "Application" -Source "Systeembeheer ST"
}

if ($Succes -eq 0){
Write-EventLog -LogName "Application" -Source "Systeembeheer ST" -EventID 1001 -EntryType Error -Message "Backup niet gelukt: `n`n$ErrorMessage `n`n $FailedItem"
echo "Foutmelding in register geschreven."
}
elseif($Succes -eq 1){
Write-EventLog -LogName "Application" -Source "Systeembeheer ST" -EventID 1000 -EntryType Information -Message "Backup gelukt"
echo "Gebeurtenis in register geschreven."
}

C:\WINDOWS\system32\net use Z: /delete #Stack verbinding voor de zekerheid nogmaals afsluiten.