$Machine = [system.environment]::MachineName
$User = (Get-WMIObject -ClassName Win32_ComputerSystem).Username
$UpTime = (get-date) - (gcim Win32_OperatingSystem).LastBootUpTime | select Hours, minutes
$NomExe = "chemin_vers_votre_programme.exe"

Add-Content C:\Info-user.txt $User, $Machine, $UpTime
Start-process -FilePath $NomExe
