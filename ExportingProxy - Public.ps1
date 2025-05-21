# PowerShell-Skript zum Exportieren von SMTP-Adressen mit dem Microsoft Graph PowerShell SDK

# Authentifizierung mit Microsoft Graph
Connect-MgGraph -Scopes "User.Read.All", "Group.Read.All", "Mail.Read" -NoWelcome

# Dateipfade und E-Mail-Parameter definieren
$aktuellesDatumUndUhrzeit = Get-Date -Format "yyyyMMdd_HHmmss"
$filePath = "C:\Batches\smtplist\sendedFiles\latestFile_$aktuellesDatumUndUhrzeit.csv"
$include = "C:\Batches\smtplist\include.txt"
$logFilePath = "C:\Batches\smtplist\Log\Export_$aktuellesDatumUndUhrzeit.log"

$From = "#Private"
$To = "#Private"
$SMTPServer = "#Private"
$SMTPPort = 25

# Logging starten
Start-Transcript -Path $logFilePath

# Arrays zum Speichern der E-Mail-Adressen initialisieren
$allsmtpAddresses = @()

# Alle Benutzer und deren Proxy-Adressen abrufen
$users = Get-MgUser -Property DisplayName, Mail, ProxyAddresses -All

# Alle Proxy-Adressen abrufen
foreach ($user in $users) {
    $allsmtpAddresses = $user.ProxyAddresses | Where-Object {$_ -like 'SMTP:*'} | ForEach-Object { $_ -replace 'SMTP:' }
}

# Alle sekundären Proxy-Adressen abrufen
foreach ($user in $users) {
    $allsmtpAddresses += $user.ProxyAddresses | Where-Object {$_ -like 'smtp:*'} | ForEach-Object { $_ -replace 'smtp:' }
}

# Primäre E-Mail-Adressen aller Benutzer abrufen
foreach ($user in $users) {
    $allsmtpAddresses += $user.Mail
}

# Alle Gruppen und deren Proxy-Adressen abrufen
$groups = Get-MgGroup -Property DisplayName, Mail, ProxyAddresses -All

# Primäre Proxy-Adressen aller Gruppen abrufen
foreach ($group in $groups) {
    $allsmtpAddresses += $group.ProxyAddresses | Where-Object {$_ -like 'SMTP:*'} | ForEach-Object { $_ -replace 'SMTP:' }
}

# Sekundäre Proxy-Adressen aller Gruppen abrufen
foreach ($group in $groups) {
    $allsmtpAddresses += $group.ProxyAddresses | Where-Object {$_ -like 'smtp:*'} | ForEach-Object { $_ -replace 'smtp:' }
}

# Primäre E-Mail-Adressen aller Gruppen abrufen
foreach ($group in $groups) {
    $allsmtpAddresses += $group.Mail
}

# Duplikate und leere Einträge entfernen
$allsmtpAddresses = $allsmtpAddresses | Where-Object { $_ -ne $null -and $_ -ne ""}

# Duplikate entfernen, nur den ersten Eintrag behalten
$allsmtpAddresses = $allsmtpAddresses | Select-Object -Unique

# "ok" an jede SMTP-Adresse anhängen
$allsmtpAddresses = $allsmtpAddresses | ForEach-Object { "$_,ok" }

# Unerwünschte Adressen filtern
$filteredLines = $allsmtpAddresses | Where-Object {
    $_ -notmatch "#Private" -and
    $_ -notlike "#Private" -and
    $_ -notlike "#Private"
}

# In Kleinbuchstaben umwandeln und in Datei speichern
$filteredLines.ToLower() | Out-File -FilePath $filePath

# Inhalt aus include.txt anhängen
$inhaltDerQuelldatei = Get-Content -Path $include
Add-Content -Path $filePath -Value $inhaltDerQuelldatei

# Zeilen in der Datei zählen
$dateiInhalt = Get-Content -Path $filePath -Raw
$zeilen = $dateiInhalt -split "`n"
$anzahlZeilen = $zeilen.Count
Write-Host "Anzahl E-Mail-Adressen: $anzahlZeilen"

# E-Mail-Parameter definieren
$aktuellesDatumUndUhrzeit = Get-Date -Format "dd.MM.yyyy HH:mm:ss"
$Subject = "#Private: $aktuellesDatumUndUhrzeit. Total $anzahlZeilen E-Mail-Adressen"
$Body = "#Private"

# E-Mail mit Anhang senden
Send-MailMessage -From $From -To $To -Subject $Subject -Body $Body -SmtpServer $SMTPServer -Port $SMTPPort -Attachments $filePath

# Logging beenden
Stop-Transcript

# Datei archivieren
Copy-Item -Path $filePath -Destination "C:\Batches\smtplist\sendedFiles\Archiv"
Remove-Item -Path "C:\Batches\smtplist\sendedFiles\latestFile.csv" -Force
Rename-Item -Path $filePath -NewName "C:\Batches\smtplist\sendedFiles\latestFile.csv"

# Verbindung zu Microsoft Graph trennen
Disconnect-MgGraph

# Ende des Skripts
