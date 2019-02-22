<#
.CREATED BY:
    Alex Meys
. CREATED ON:
    18/02/2015
.SYNOPSIS
    Basic cleanup program (Exchange, IIS, Windows en Updates)
.DESCRIPTION
    Opschonen van Windows Temp bestanden zowel als Exchange Logs en Program logs van Exchange.
    Tevens zullen de IIS Logs ook verwijderd worden (op basis van dagen ingesteld).
    IIS: 10 dagen
    Exchange: 30 Dagen
    Windows: 10 dagen
    Updates: Nvt
    Logs van Programma: Elke keer opnieuw bij het openen.
#>

#Datum gegevens (voor vergelijking met logdatum
$vandaag = Get-Date
$verleden_iis = ($vandaag).AddDays(-10)
$verleden_exch = ($vandaag).AddDays(-30)
$verleden_Win = ($vandaag).AddDays(-5)
$verleden_log = ($vandaag).AddDays(-1)

$dag = ($vandaag).Year
$maand = ($vandaag).Month
$jaar =  ($vandaag).Day
$min = ($vandaag).Minute
$sec = ($vandaag).Second

$objShell = New-Object -ComObject Shell.Application 
$objFolder = $objShell.Namespace(0xA)

#Get Username
$me = $Env:USERPROFILE

#Maak log directory
$logfile = "C:\log"
$bestaat = Test-Path($logfile)

if ($bestaat -eq $True)
{
    ls -Recurse $logfile | Where {$_.Name -Like "*.txt"} | Remove-Item
}
else
{
    mkdir "C:\log"
}

#Verwijderen van Windows TEMP Logs
Write-Host "Windows Temp Files (5 dagen)"
$a = ls -Recurse -Force C:\Windows\Temp\ | Where {$_.LastWriteTime -lt $verleden_Win}
$a | Out-File "C:\log\opkuis_Win_$jaar$maand$dag.txt"
$a | Remove-Item -Force -ErrorAction SilentlyContinue -Recurse

$b = ls -Recurse -Force "C:\Windows\Microsoft.NET\Framework64\V4*\Temporary ASP.NET Files\" | Where {$_.LastWriteTime -lt $verelden_Win}
$b | Out-File "C:\log\opkuis_Win_1_$jaar$maand$dag.txt"
$b | Remove-Item -Force -ErrorAction SilentlyContinue -Recurse

$c = ls -Recurse -Force -ErrorAction SilentlyContinue "C:\users\*\appdata\local\temp\" | Where {$_.LastWriteTime -lt $verleden_Win}
$c | Out-File "C:\log\opkuis_Win_2_$jaar$maand$dag.txt"
$c | Remove-Item -Force -ErrorAction SilentlyContinue -Recurse

$d = ls -Recurse -Force $me\Downloads\ | Where {$_.LastWriteTime -lt $verelden_Win}
$d | Out-File "C:\log\opkuis_Win_3_$jaar$maand$dag.txt"
$d | Remove-Item -Force -ErrorAction SilentlyContinue -Recurse

#Verwijderen van IIS logs ouder dan $verleden:
write-Host "inetpub Files (10 dagen)"
$e = ls -Recurse -Force C:\InetPub\Logs\LogFiles\ | Where {$_.Name -Like "u*.log" -And $_.LastWriteTime -lt $verleden_iis }
$e = Out-File "C:\log\opkuis_IIS_1_$jaar$maand$dag.txt"
$e | Remove-Item -Force -ErrorAction SilentlyContinue -Recurse

#Verwijderen logs van Exchange Programma:
Write-Host "Exchange Program Files (30 dagen)"

$locatie0 = "C:\Program Files\Microsoft\Exchange Server\V15\Logging" 
$loc0 = Test-Path $locatie0
If ($loc0 -eq $True)
{
    #Leegmaken van sp variable voor nieuwe input
    $sp = ""
    $sp = ls -Recurse -Force $locatie0 | Where {$_.LastWriteTime -lt $verleden_exch} 
    $sp | Out-File "C:\log\opkuis_Exchange_$jaar$maand$dag$min$sec.txt"
    $sp | Remove-Item -Force -ErrorAction SilentlyContinue -Recurse
}
else
{
    write-Host " "
    Write-Host "Geen Exchange 2013 gevonden"
    write-Host " "

}

Write-Host "Exchange Logs (30 dagen) - Indien niet aanwezig geen output"
$locaties = @("C:\Program Files\Microsoft\Exchange Server\V15","E:\Program Files\Microsoft\Exchange Server\V15","C:\Program Files\Microsoft\Exchange Server\V14","E:\Program Files\Microsoft\Exchange Server\V14", "E:\Exchsrv\", "E:\Exchange\", "E:\Program Files\Microsoft\Exchange Server\Mailbox\")

Foreach ($locatie in $locaties)
{
    if ((Test-Path($locatie)) -eq $True)
	{
        $sp2 = ""
	    $sp2 = ls -Recurse -Force $locatie  | Where {$_.LastWriteTime -lt $verleden_exch -and $_.Name -Like "E000*EE*.log"} 
        $sp2 | Out-File "C:\log\opkuis_Exchange_$jaar$maand$dag$min$sec.txt" 
        $sp2 | Remove-Item -Force -ErrorAction SilentlyContinue -Recurse
	}
	else 
	{
	    continue
	}
}
#Verwijderen van Software Updates
Write-Host " "
Write-Host "Opkuisen van Updates"
Get-Service -Name wuauserv | Stop-Service -Force -ErrorAction SilentlyContinue
$f = ls -Recurse -Force -ErrorAction SilentlyContinue "C:\Windows\SoftwareDistribution\*"
$f | Out-File "C:\log\opkuis_update_$jaar$maand$dag.txt"
$f | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue

#Start updates clean search
Get-Service -Name wuauserv | Start-Service

#Leegmaken van prullenbak
Write-Host " "
Write-Host "Opkuisen Prullenbak"
$objFolder.Items() | %{Remove-Item $_.Path -Force -Recurse}

Write-Host " "
Write-Host " "
Write-Host "****** *******"
Write-Host "** All Done **"
Write-Host "****** *******"
Write-Host " "
