########################################################################
# Created by David Cannon  Date:8-17-15                                #
#   This script updates a machine using windows update or WSUS server  #
########################################################################
###########
#Variables#
###########
#creates a name for the log file based on the date
$LogFileName = Get-Date -UFormat "%m-%d-%Y"
New-Item -ItemType File -Path "\\<Path-to-log-file\$LogFileName.log" -ea SilentlyContinue
$log = "\\<Path-to-log-file>\$LogFileName.log"
$status = Get-Service wuauserv
#add a break in the log file with the computer name 
Add-Content -Path $log "########## Starting log for $env:COMPUTERNAME ##########"
#######################################
#See if the update service is running #
#######################################
if ($status.Status -eq "Stopped"){Write-Host "Stopped"
    Add-Content -Path $log "$env:COMPUTERNAME Setting Update Service to Manual"
    Set-Service wuauserv -StartupType Manual
    Add-Content -Path $log "$env:COMPUTERNAME Update Service set to Manual"
    Add-Content -Path $log "$env:COMPUTERNAME Starting Update Service"
    Start-Service wuauserv
    Add-Content -Path $log "$env:COMPUTERNAME Update Service Started"
#Define update criteria.
    $Criteria = "IsInstalled=0 and Type='Software'"
#Search for relevant updates.
    Add-Content -Path $log "$env:COMPUTERNAME Searching for updates"
    $Searcher = New-Object -ComObject Microsoft.Update.Searcher
    $SearchResult = $Searcher.Search($Criteria).Updates
    Add-Content -Path $log "$env:COMPUTERNAME Done Searching for updates"
#if NO Updates    
    if($Error){
        Add-Content -Path $log "[Error]:$env:COMPUTERNAME No Updates Available"
    #stopping services
        Add-Content -Path $log "$env:COMPUTERNAME Update Service Stopping"
        Stop-Service wuauserv
        Add-Content -Path $log "$env:COMPUTERNAME Update Service Stopped"
        Add-Content -Path $log "$env:COMPUTERNAME Setting Service to Disabled"
        Set-Service wuauserv -StartupType Disabled
        Add-Content -Path $log "$env:COMPUTERNAME Service set to Disabled"   
        Add-Content -Path $log "########## End of log for $env:COMPUTERNAME ##########"
        Add-Content -Path $log "########## $env:COMPUTERNAME No Reboot Necessary ##########"
        }
        Else{
#Download updates.
    Add-Content -Path $log "$env:COMPUTERNAME Downloading Updates from WSUS Server"
    $Session = New-Object -ComObject Microsoft.Update.Session
    $Downloader = $Session.CreateUpdateDownloader()
    $Downloader.Updates = $SearchResult
    $Downloader.Download()
    Add-Content -Path $log "$env:COMPUTERNAME Downloading Updates from WSUS Server Complete"
#Install updates.
    Add-Content -Path $log "$env:COMPUTERNAME Installing Updates"
    $Installer = New-Object -ComObject Microsoft.Update.Installer
    $Installer.Updates = $SearchResult
    $Result = $Installer.Install()
    Add-Content -Path $log "$env:COMPUTERNAME Installing Updates Complete"
#stopping services
    Add-Content -Path $log "$env:COMPUTERNAME Update Service Stopping"
    Stop-Service wuauserv
    Add-Content -Path $log "$env:COMPUTERNAME Update Service Stopped"
    Add-Content -Path $log "$env:COMPUTERNAME Setting Service to Disabled"
    Set-Service wuauserv -StartupType Disabled
    Add-Content -Path $log "$env:COMPUTERNAME Service set to Disabled"        
#Reboot if required by updates.
    If ($Result.rebootRequired) { shutdown.exe /t 0 /r 
    Add-Content -Path $log "$env:COMPUTERNAME Rebooting PC"
    Add-Content -Path $log "########## End of log for $env:COMPUTERNAME ##########"}
    else{
    Add-Content -Path $log "########## End of log for $env:COMPUTERNAME ##########"
    Add-Content -Path $log "########## $env:COMPUTERNAME No Reboot Necessary ##########"
    }
}
}
########################
#It is already Running #
########################
ELSEIF($status.Status -eq "Running"){
    Add-Content -Path $log "$env:COMPUTERNAME Update Service Was Running"
    
#Define update criteria.
    $Criteria = "IsInstalled=0 and Type='Software'"

#Search for relevant updates.
    Add-Content -Path $log "$env:COMPUTERNAME Searching for updates"
    $Searcher = New-Object -ComObject Microsoft.Update.Searcher
    $SearchResult = $Searcher.Search($Criteria).Updates
    Add-Content -Path $log "$env:COMPUTERNAME Done Searching for updates"
    
#if NO Updates    
    if($Error){
        Add-Content -Path $log "[Error]:$env:COMPUTERNAME No Updates Available"
    #stopping services
        Add-Content -Path $log "$env:COMPUTERNAME Update Service Stopping"
        Stop-Service wuauserv
        Add-Content -Path $log "$env:COMPUTERNAME Update Service Stopped"
        Add-Content -Path $log "$env:COMPUTERNAME Setting Service to Disabled"
        Set-Service wuauserv -StartupType Disabled
        Add-Content -Path $log "$env:COMPUTERNAME Service set to Disabled"   
        Add-Content -Path $log "########## End of log for $env:COMPUTERNAME ##########"
        Add-Content -Path $log "########## $env:COMPUTERNAME No Reboot Necessary ##########"
        }
        Else{
#Download updates.
    Add-Content -Path $log "$env:COMPUTERNAME Downloading Updates from WSUS Server"
    $Session = New-Object -ComObject Microsoft.Update.Session
    $Downloader = $Session.CreateUpdateDownloader()
    $Downloader.Updates = $SearchResult
    $Downloader.Download()
    Add-Content -Path $log "$env:COMPUTERNAME Downloading Updates from WSUS Server Complete"
#Install updates.
    Add-Content -Path $log "$env:COMPUTERNAME Installing Updates"
    $Installer = New-Object -ComObject Microsoft.Update.Installer
    $Installer.Updates = $SearchResult
    $Result = $Installer.Install()
    Add-Content -Path $log "$env:COMPUTERNAME Installing Updates Complete"
#stopping services
    Add-Content -Path $log "$env:COMPUTERNAME Update Service Stopping"
    Stop-Service wuauserv
    Add-Content -Path $log "$env:COMPUTERNAME Update Service Stopped"
    Add-Content -Path $log "$env:COMPUTERNAME Setting Service to Disabled"
    Set-Service wuauserv -StartupType Disabled
    Add-Content -Path $log "$env:COMPUTERNAME Service set to Disabled"        
#Reboot if required by updates.
    If ($Result.rebootRequired) { shutdown.exe /t 0 /r 
    Add-Content -Path $log "$env:COMPUTERNAME Rebooting PC"
    Add-Content -Path $log "########## End of log for $env:COMPUTERNAME ##########"}
    else{
    Add-Content -Path $log "########## End of log for $env:COMPUTERNAME ##########"
    Add-Content -Path $log "########## $env:COMPUTERNAME No Reboot Necessary ##########"
    }
}

}
