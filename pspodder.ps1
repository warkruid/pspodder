#requires -version 3.0 -Modules BitsTransfer, Remove-Directory
#=============================================================================
# Name   : PsPodder
# Comment: Based on Bashpodder (http://lincgeekcom/bashpodder)
#          
# Author : Enneman
# Description:
#          read rss feeds from pspodder.conf and fetch the podcasts in a directory 
#          each day
#          for example http://podcasts.scpr.org/loh_down
# Requirements: Remove-Directory module
# Bugs   : Not every rss feed is parsed cleanly? 
# Todo   : Better errorhandling         
#          Check correct scope of variables
#          Style? 
#          Port configfile and logfile to sqlite?
#          Nail down bug in parsing of feeds
#          Add commandline options?
#=============================================================================

#=============================================================================
# Set variables
#=============================================================================
$UserDir    = $env:USERPROFILE
$PodcastDir = 'D:\PODCASTS'

# Choose a directory were you can write to. Note: if malware/ransomware protection
# is activated you only have permission to write in Downloads or temp on you C: drive.
if(!(Test-Path -Path $PodcastDir )){
    $PodcastDir = (${UserDir} + '\Downloads\PODCASTS' -f $UserDir)
}
$LogFile    = ('{0}\pspodder.log' -f $PSScriptRoot)
$ConfFile   = ('{0}\pspodder.conf' -f $PSSCriptRoot) 
$ErrorFile  = ('{0}\pspodder.err' -f $PSScriptRoot) 
$DayDir     = ('{0}\{1}' -f $PodcastDir, (Get-Date).ToString('yyyy-MM-dd'))
$Playlist   = ('{0}\{1}-podcasts.m3u' -f $DayDir, (Get-Date).ToString('yyyy-MM-dd'))
$MaxItems   = 3 # Max downloads for each url
$MaxDays    = 3 # Days to keep
$Player  = 'C:\Program Files\VideoLAN\VLC\vlc.exe'
#=============================================================================
# Checks
#=============================================================================

# Check if a network connection exists
if (! (Get-NetAdapter -Physical | Where-Object -Property status -EQ -Value 'up')) 
{
  Write-Output -InputObject 'No Network connection found' 
  exit
}

# Check if the file with the rss urls exists
if (!(Test-Path -Path $ConfFile -IsValid)) 
{
  Write-Output -InputObject 'No podcast file found' 
  exit
}
else 
{
  # Load entire file
  $PodcastConf = Get-Content -Path ('{0}' -f $ConfFile)
}

if (!(Test-Path -Path $LogFile))
{
  Write-Output -InputObject 'Create an empty log file' 
  $LogFile = New-Item -ItemType File -Path ('{0}' -f $LogFile) 
}

# Test if the Error file exists
if (!(Test-Path -Path $ErrorFile))
{
  Write-Output -InputObject 'Create an empty error file'
  $ErrorFile = New-Item -ItemType File -Path ('{0}' -f $ErrorFile) 
}

#=============================================================================
# Call Module Remove-Directory
#=============================================================================
Remove-Directory -directory $PodcastDir -interval 'days' -count $MaxDays 

Import-Module -Name BitsTransfer

#=============================================================================
# Main
#=============================================================================

try
{
  foreach ($line in $PodcastConf)
  {
    if ($line -match "^#.*") { continue }
    Write-Output -InputObject ('Checking: {0}' -f $line)
    
    if (!(Test-Path -Path $DayDir))
    {
      Write-Output  -InputObject ('Creating directory: {0}' -f $DayDir)
      New-Item -ItemType Directory -Path $DayDir
    }  
    Set-Location -Path $DayDir   
        
    # Haul in net.webclient object  from .NET

    Try
    {
      $a = ([xml](New-Object -TypeName net.webclient -ErrorAction Continue).downloadstring(('{0}' -f $line)))
    }
    Catch {
      Write-Error -Message ('Errors occured downloading from {0}' -f $line) 
      $_ | Out-File -FilePath ('{0}' -f $ErrorFile) -Append
    }
    $count = 0 # Number of items downloaded


    $a.rss.channel.item | ForEach-Object -Process {  
      Try 
      {
        $url = New-Object -TypeName System.Uri -ArgumentList ($_.enclosure.url) -ErrorAction Stop
      
      }
      Catch 
      {
        Write-Error -Message ('Errors occured extracting RSS info from {0}' -f $line) 
        $_ | Out-File -FilePath ('{0}' -f $ErrorFile) -Append
      }
      $file = $url.Segments[-1]
      $count++

      # If not already downloaded $MaxItems of files, and $file not in $logfile

      if (($count -le $MaxItems) -and !(Select-String -Path ('{0}' -f $LogFile) -Pattern ('{0}' -f $file))) 
      {
        Write-Output -InputObject ('Downloading: {0}' -f $file)
        Write-Output -InputObject ('URL = {0}' -f $url)
        # Test if file name already exists, if yes, generate a random string
        # and prefix it to the original filename
            
        While ((Test-Path -Path $DayDir/$file)) 
        {
          [String]$RandomString = Get-Random
          $file = $RandomString + $file
        }
        Try 
        {
          #(New-Object -TypeName System.Net.WebClient -ErrorAction Stop).DownloadFile($url,('{0}\{1}' -f $DayDir, $file))
          (Start-BitsTransfer -Source $url -Destination ('{0}\{1}' -f $DayDir, $file) -ErrorAction Stop)
        }
        catch 
        {
          # get error record
          [Management.Automation.ErrorRecord]$e = $_

          # retrieve information about runtime error
          $info = New-Object -TypeName PSObject -Property @{
            Exception = $e.Exception.Message
            Reason    = $e.CategoryInfo.Reason
            Target    = $e.CategoryInfo.TargetName
            Script    = $e.InvocationInfo.ScriptName
            Line      = $e.InvocationInfo.ScriptLineNumber
            Column    = $e.InvocationInfo.OffsetInLine
          }
        
          # output information. Post-process collected info, and log info (optional)
          $info | Out-File -FilePath ('{0}' -f $ErrorFile) -Append
        }
        $file | Out-File -FilePath ('{0}' -f $LogFile) -Append
      }
    }
  }
}

catch [System.Net.WebException]
{
  # get error record
  [Management.Automation.ErrorRecord]$e = $_

  # retrieve information about runtime error
  $info = [PSCustomObject]@{
    Exception = $e.Exception.Message
    Reason    = $e.CategoryInfo.Reason
    Target    = $e.CategoryInfo.TargetName
    Script    = $e.InvocationInfo.ScriptName
    Line      = $e.InvocationInfo.ScriptLineNumber
    Column    = $e.InvocationInfo.OffsetInLine
  }
  
  # output information. Post-process collected info, and log info (optional)
  $info
}



# Make a podcast.m3u file

Write-Output -InputObject 'Generating playlist' 
Get-ChildItem -Exclude *.m3u |
Where-Object -FilterScript {
  !$_.PsIsContainer
} |

ForEach-Object -Process {
  $_.Name
} > $Playlist


#TODO make an option for downloading OR streaming
# Check if vlcplayer exists
if ((Test-Path -Path $Player -IsValid)) 
{
  Start-Process -FilePath $Player $Playlist 
}

#====================================================================================
# End
#====================================================================================