#Syntax: .\change_wmsvc-certificate.ps1 *Text file with servernames* *path and private certificate* *secret key for private certificate*
#Example: .\change_wmsvc-certificate.ps1 Liste.txt "C:\privat.pfx" cert_password

#Variablen
$script:ver = "1.4"
$script:verdate = "17.03.2022"

$script:listname = $Args[0]
$script:pfxfile = $Args[1]
$script:pfxkey = $Args[2]

$script:startdate = get-date -uformat "%d.%m.%Y"
$script:starttime = get-date -uformat "%R"


#Informationsblock Anfang
Write-Output "=========================================================================="
Write-Output "Change WMSvc Certificate Script Ver. $ver, $verdate"
Write-Output "Written by Andyt for face of buildings planning stimakovits GmbH"
Write-Output "Promoted development by BlackSeals.net Technology"
Write-Output "Copyright 2017-2022 by Reisenhofer Andreas"
Write-Output "=========================================================================="
Write-Output "Gestartet am $startdate um $starttime Uhr..."
Write-Output ""

#Prüfe Liste
$script:checkpathlist = test-path -path ($listname)
$script:checkpathpfx = test-path -path ($pfxfile)

if (($checkpathlist -eq "True") -and ($checkpathpfx -eq "True")) {
	#Remotesitzungen aufbauen
	Write-Output "Verbindung wird aufgebaut..."
	$script:sessions = New-PSSession –ComputerName (get-content $listname)
	
	#Lade ServerManager-Powershell Modul
	Invoke-Command –Session $Sessions -ScriptBlock {Import-Module ServerManager}

	#Prüfen ob WMSVC installiert ist und bei Bedarf installieren
	Write-Output "Kontrolle ob IIS-Verwaltung (WMSVC) installiert ist..."
	Invoke-Command –Session $Sessions -ScriptBlock {$checkwindowsfeature = Get-WindowsFeature | where {$_.Name -like "Web-Mgmt-Service"}}
	Invoke-Command –Session $Sessions -ScriptBlock {if ($checkwindowsfeature.InstallState -eq "Available") {Add-WindowsFeature Web-Mgmt-Service}}

	#Aktivierung von MWSVC nach Installation
	Write-Output "IIS-Verwaltung wird aktiviert..."
	Invoke-Command –Session $Sessions -ScriptBlock {if ($checkwindowsfeature.InstallState -eq "Installed") {Set-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\WebManagement\Server -Name EnableRemoteManagement -Value 1}}

	#WMSVC wird beendet und nicht mehr automatisch gestartet
	Write-Output "IIS-Verwaltung wird beendet und deaktiviert..."
	Invoke-Command –Session $sessions -ScriptBlock {$checkservice = Get-Service | where {$_.Status -like "Running"} | where {$_.Name -like "WMSVC"}}
	Invoke-Command –Session $sessions -ScriptBlock {if ($checkservice.Status -eq "Running") {Set-Service -name WMSVC -StartupType Manual}}
	Invoke-Command –Session $Sessions -ScriptBlock {if ($checkservice.Status -eq "Running") {Stop-service WMSVC}}
	
	#Vorhandene Zertifikate entfernen
	Write-Output "Altes Zertifikat wird entfernt..."
	Invoke-Command -Session $sessions -ScriptBlock {Get-ChildItem -Path Cert:\LocalMachine\My | Where {$_.FriendlyName -like "*WMSvc"} | Where {$_.Issuer -like "*Certificate-Authority*"} | Remove-Item}
	
	#Zertifikat kopieren und installieren
	Write-Output "Neues Zertifikat wird installiert..."
	(get-content $listname) | Foreach-Object {Copy-Item -Path "$pfxfile" -Destination "\\$_\c$"}
	$pfxfile = ((Get-Item $pfxfile ).Name)
	Invoke-Command -Session $sessions {param($spfxkey=$pfxkey, $spfxfile=$pfxfile) certutil -f -p $spfxkey -importpfx C:\$spfxfile} -ArgumentList $pfxkey,$pfxfile
	(get-content $listname) | Foreach-Object {Remove-Item -Path "\\$_\c$\$pfxfile"}
	
	#Prüfen ob IS-Powershell Modul installiert ist und bei Bedarf installieren
	Invoke-Command –Session $Sessions -ScriptBlock {$checkwindowsfeature = Get-WindowsFeature | where {$_.Name -like "Web-Scripting-Tools"}}
	Invoke-Command –Session $Sessions -ScriptBlock {if ($checkwindowsfeature.InstallState -eq "Available") {Add-WindowsFeature Web-Scripting-Tools}}
	
	#Lade IIS-Powershell Modul
	Invoke-Command -Session $sessions {Import-Module WebAdministration}
	
	#Suche neues Zertifikat, Entferne altes Zertifikat und Installiere neues Zertifikat
	Write-Output "Zuweisung des neuen Zertifikates..."
	Invoke-Command -Session $sessions -ScriptBlock {$cert = Get-ChildItem -Path Cert:\LocalMachine\My | Where {$_.FriendlyName -like "*WMSvc"} | Where {$_.Issuer -like "*Certificate-Authority*"} | Where {$_.NotAfter -gt (get-date)} | Select-Object -ExpandProperty Thumbprint}
	Invoke-command -Session $sessions -ScriptBlock {Remove-Item -Path IIS:\SslBindings\0.0.0.0!8172}
	Invoke-Command -Session $sessions -ScriptBlock {Get-Item -Path "cert:\localmachine\my\$cert" | New-Item -Path IIS:\SslBindings\0.0.0.0!8172}
	
	#WMSVC wird gestartet und automatisch gestartet
	Write-Output "IIS-Verwaltung wird gestartet und aktiviert..."
	Invoke-Command –Session $sessions -ScriptBlock {$checkservice = Get-Service | where {$_.Status -like "Stopped"} | where {$_.Name -like "WMSVC"}}
	Invoke-command –Session $sessions -ScriptBlock {if ($checkservice.Status -eq "Stopped") {Set-Service -name WMSVC -StartupType Automatic}}
	Invoke-command –Session $sessions -ScriptBlock {if ($checkservice.Status -eq "Stopped") {Start-service WMSVC}}
	
	#Schließe alle offenen Remotesitzungen
	Write-Output "Verbindung werden geschlossen..."
	Get-PSSession | Remove-PSSession
		
} else {
    Write-Host "Es wurde ein wichtiger Pfad nicht gefunden!" -ForegroundColor red
	if ($checkpathlist -eq "True") {
		Write-Host "Die Datei $listname wurde gefunden." -ForegroundColor green
	} else {
		Write-Host "Die Datei $listname konnte nicht gefunden werden." -ForegroundColor red	
	}
	if ($checkpathpfx -eq "True") {
		Write-Host "Die Datei $pfxfile wurde gefunden." -ForegroundColor green
	} else {
		Write-Host "Die Datei $pfxfile konnte nicht gefunden werden." -ForegroundColor red
	}
	Write-Host ""
	Write-Host ""
	
	$script:erroravailable = 1
}

#Informationsblock Ende
Write-Output "Abarbeitung am $(get-date -uformat "%d.%m.%Y") um $(get-date -uformat "%R") Uhr beendet."
Write-Output ""