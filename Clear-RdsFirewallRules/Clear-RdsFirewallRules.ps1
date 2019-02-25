############################
###        Header        ###
############################

# 2019 - Josef Albert
# Clear-RdsFirewallRules

# Sources
# https://www.ugg.li/windows-server-2016-rds-sh-effekte/
# https://social.technet.microsoft.com/Forums/en-US/8dad5b1e-8236-4792-85fe-8725d74bbbcb/start-menu-not-coming-up-server-2016-rds?forum=winserverTS

############################
###      Transcript      ###
############################

$TranscriptFolder = 'C:\Scripts\Logs\Clear-RdsFirewallRules'
$TranscriptFileName = "$(Get-Date -Format 'yyyy-MM-dd_HH-mm')" + '_Clear-RdsFirewallRules.log'
$TranscriptPath = $TranscriptFolder + '\' + $TranscriptFileName

if (!(Test-Path "$TranscriptFolder")) {
    New-Item "$TranscriptFolder" -ItemType Directory
}

# Delete old logs (30 days)
Get-ChildItem "$TranscriptFolder" -Recurse -File | Where CreationTime -lt (Get-Date).AddDays(-30)  | Remove-Item -Force

Start-Transcript -Path "$TranscriptPath"

############################
###        Script        ###
############################

Clear-Host

Write-Host ''
Write-Host '####################################'
Write-Host '###    Clear-RdsFirewallRules    ###'
Write-Host '####################################'
Write-Host ''

$profiles = get-wmiobject -class win32_userprofile

Write-Host 'Getting Firewall Rules...'

$AllRules1 = Get-NetFirewallRule -All -ErrorAction SilentlyContinue

if ($?) {
    $Rules1 = $AllRules1 | Where-Object {$profiles.sid -notcontains $_.owner -and $_.owner }
    $Rules1Count = $Rules1.count
    Write-Host ' -> OK' -ForegroundColor Green
    Write-Host " -> $Rules1Count Rules"
} else {
    Write-Host ' -> Error!' -ForegroundColor Red
    Write-Host ' -> Stopping Script!' -ForegroundColor Red
    break
}

Write-Host ''
Write-Host 'Getting Firewall Rules from ConfigurableServiceStore...'

$AllRules2 = Get-NetFirewallRule -All -PolicyStore ConfigurableServiceStore -ErrorAction SilentlyContinue

if ($?) {
    $Rules2 = $AllRules2 | Where-Object { $profiles.sid -notcontains $_.owner -and $_.owner }
    $Rules2Count = $Rules2.count
	Write-Host ' -> OK' -ForegroundColor Green
    Write-Host " -> $Rules2Count Rules"
} else {
    $ErrorReadingRules2 = $true
    Write-Host ' -> Warning!' -ForegroundColor Yellow
    Write-Host ' -> Reading ConfigurableServiceStore failed... Possible Overflow!' -ForegroundColor Yellow
}

$Total = $Rules1.count + $Rules2.count


$Result = Measure-Command {

    $start = (Get-Date)
    $i = 0.0

    Write-Host ''
    Write-Host "Deleting $($Rules1.Count) Firewall Rules..."
  
    # action
    try {
        foreach($rule1 in $Rules1){

            Remove-ItemProperty -Path "HKLM:\System\CurrentControlSet\Services\SharedAccess\Parameters\FirewallPolicy\FirewallRules" -Name $rule1.name

            # progress
            $i = $i + 1.0
            $prct = $i / $total * 100.0
            $elapsed = (Get-Date) - $start
            $totaltime = ($elapsed.TotalSeconds) / ($prct / 100.0)
            $remain = $totaltime - $elapsed.TotalSeconds
            $eta = (Get-Date).AddSeconds($remain)

            # display
            $prctnice = [math]::round($prct,2) 
            $elapsednice = $([string]::Format("{0:d2}:{1:d2}:{2:d2}", $elapsed.hours, $elapsed.minutes, $elapsed.seconds))
            $speed = $i/$elapsed.totalminutes
            $speednice = [math]::round($speed,2) 
            Write-Progress -Activity "Deleting Rules ETA $eta elapsed $elapsednice loops/min $speednice" -Status "$prctnice" -PercentComplete $prct -SecondsRemaining $remain
        }
        Write-Host ' -> OK' -ForegroundColor Green
    } catch {
        Write-Host ' -> Error!' -ForegroundColor Red
    }

    Write-Host ''
    Write-Host "Deleting $($Rules2.Count) Firewall Rules from ConfigurableServiceStore..."
    if (!$ErrorReadingRules2) {
        try {
            foreach($rule2 in $Rules2) {
            
                # action  
                Remove-ItemProperty -Path "HKLM:\System\CurrentControlSet\Services\SharedAccess\Parameters\FirewallPolicy\RestrictedServices\Configurable\System" -Name $rule2.name

                # progress
                $i = $i + 1.0
                $prct = $i / $total * 100.0
                $elapsed = (Get-Date) - $start
                $totaltime = ($elapsed.TotalSeconds) / ($prct / 100.0)
                $remain = $totaltime - $elapsed.TotalSeconds
                $eta = (Get-Date).AddSeconds($remain)

                # display
                $prctnice = [math]::round($prct,2) 
                $elapsednice = $([string]::Format("{0:d2}:{1:d2}:{2:d2}", $elapsed.hours, $elapsed.minutes, $elapsed.seconds))
                $speed = $i/$elapsed.totalminutes
                $speednice = [math]::round($speed,2) 
                Write-Progress -Activity "Deleting Rules from ConfugurableServiceStore ETA $eta elapsed $elapsednice loops/min $speednice" -Status "$prctnice" -PercentComplete $prct -secondsremaining $remain
            }
            Write-Host ' -> OK' -ForegroundColor Green
        } catch {
            Write-Host ' -> Error!' -ForegroundColor Red
        }
    } else {

        Write-Host ' -> Deleting not possible because of an reading error above...' -ForegroundColor Yellow
        Write-Host ' -> Trying to recreate ConfigurableServiceStore...' -ForegroundColor Yellow

        Write-Host ''
        Write-Host 'Recreating ConfigurableServiceStore...'

        try {
            Remove-Item "HKLM:\SYSTEM\CurrentControlSet\Services\SharedAccess\Parameters\FirewallPolicy\RestrictedServices\Configurable\System"
            New-Item "HKLM:\SYSTEM\CurrentControlSet\Services\SharedAccess\Parameters\FirewallPolicy\RestrictedServices\Configurable\System"
            Write-Host ' -> OK' -ForegroundColor Green
        } catch {
            Write-Host ' -> Error!' -ForegroundColor Red
            Write-Host ' -> Stopping Script!' -ForegroundColor Red
            break
        }

    } 
}

Write-Host ''
Write-Host 'Statistics...'
Write-Host ''
$end = Get-Date
Write-Host " -> Start: $end"
Write-Host " -> ETA: $eta"
Write-Host " -> Runtime: " $result.minutes min $result.seconds sec
Write-Host ''

Stop-Transcript

Write-Host ''
Write-Host 'Script will automatically close in 30 seconds...'
Start-Sleep -s 30