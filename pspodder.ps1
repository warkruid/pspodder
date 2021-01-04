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
# Bugs   : 1. Not every rss feed is parsed cleanly? 
#          2. Output to .m3u contains garbage chars.
# Todo   : Better errorhandling         
#          Check correct scope of variables
#          Style? 
#          Port configfile and logfile to sqlite?
#          Add commandline options?
#=============================================================================

#=============================================================================
# Set variables
#=============================================================================
$text1 = '{0}'
$itemtype = 'File'
$dateformat = 'yyyy-MM-dd'
$UserDir    = $env:USERPROFILE
$PodcastDir = 'D:\PODCASTS'

# Choose a directory were you can write to. Note: if malware/ransomware protection
# is activated you only have permission to write in Downloads or temp on you C: drive.
if(!(Test-Path -Path $PodcastDir )){
    $PodcastDir = ('{0}\Downloads\PODCASTS' -f $UserDir)
}
$LogFile    = ('{0}\pspodder.log' -f $PSScriptRoot)
$ConfFile   = ('{0}\pspodder.conf' -f $PSSCriptRoot) 
$ErrorFile  = ('{0}\pspodder.err' -f $PSScriptRoot) 
$DayDir     = ('{0}\{1}' -f $PodcastDir, (Get-Date).ToString($dateformat))
$Playlist   = ('{0}\{1}-podcasts.m3u' -f $DayDir, (Get-Date).ToString($dateformat))
$MaxItems   = 3 # Max downloads for each url
$MaxDays    = 3 # Days to keep
$Player     = "$env:ProgramFiles(x86)\VideoLAN\VLC\vlc.exe"
#Write-Output -InputObject $Player
#exit

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
  $PodcastConf = Get-Content -Path ($text1 -f $ConfFile)
}

if (!(Test-Path -Path $LogFile))
{
  Write-Output -InputObject 'Create an empty log file' 
  $LogFile = New-Item -ItemType $itemtype -Path ($text1 -f $LogFile) 
}

# Test if the Error file exists
if (!(Test-Path -Path $ErrorFile))
{
  Write-Output -InputObject 'Create an empty error file'
  $ErrorFile = New-Item -ItemType $itemtype -Path ($text1 -f $ErrorFile) 
}

#Function Get-FileMetaData 
#{
  <#
      .SYNOPSIS
      Describe purpose of "Get-FileMetaData" in 1-2 sentences.

      .DESCRIPTION
      Add a more complete description of what the function does.

      .PARAMETER folder
      Describe parameter -folder.

      .EXAMPLE
      Get-FileMetaData -folder Value
      Describe what this call does

      .NOTES
      Place additional notes here.

      .LINK
      URLs to related sites
      The first link is opened by Get-Help -Online Get-FileMetaData

      .INPUTS
      List of input types that are accepted by this function.

      .OUTPUTS
      List of output types produced by this function.
  #>

  # Cribbed from https://gallery.technet.microsoft.com/scriptcenter/get-file-meta-data-function-f9e8d804 
  #Param([Parameter(Mandatory,HelpMessage='Show metadata of file')][string[]]$folder) 
  #foreach($sFolder in $folder) 
  #{ 
  #  $a = 0 
  #  $objShell = New-Object -ComObject Shell.Application 
  #  $objFolder = $objShell.namespace($sFolder) 
 
  #  foreach ($File in $objFolder.items()) 
  #  {  
  #    $FileMetaData = New-Object -TypeName PSOBJECT 
  #    for ($a ; $a  -le 266; $a++) 
  #     {  
  #       if($objFolder.getDetailsOf($File, $a)) 
  #         { 
  #           $hash += @{$($objFolder.getDetailsOf($objFolder.items, $a))  = 
  #                 $($objFolder.getDetailsOf($File, $a)) } 
  #          $FileMetaData | Add-Member
  #          $hash.clear()  
  #         }  
  #     }   
  #    $a=0 
  #    $FileMetaData 
  #  } 
  #}  
#}
 
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
    if ($line -match '^#.*') { continue }
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
      $a = ([xml](New-Object -TypeName net.webclient -ErrorAction Continue).downloadstring(($text1 -f $line)))
    }
    Catch {
      Write-Error -Message ('Errors occured downloading from {0}' -f $line) 
      $_ | Out-File -FilePath ($text1 -f $ErrorFile) -Append
    }
    $count = 0 # Number of items downloaded


    $a.rss.channel.item | ForEach-Object -Process {  
                            $text2 = '{0}'
                            $erroraction = 'Stop'
                            Try 
                            {
                              $url = New-Object -TypeName System.Uri -ArgumentList ($_.enclosure.url) -ErrorAction $erroraction
      
                            }
                            Catch 
                            {
                              Get-ChildItem -Recurse | Where-Object{ $_.PSIsContainer } Write-Error -Message ('Errors occured extracting RSS info from {0}' -f $line) 
                              $_ | Out-File -FilePath ($text2 -f $ErrorFile) -Append
                            }
                            $file = $url.Segments[-1]
                            $count++

                            # If not already downloaded $MaxItems of files, and $file not in $logfile

                            if (($count -le $MaxItems) -and !(Select-String -Path ($text2 -f $LogFile) -Pattern ($text2 -f $file))) 
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
                                (Start-BitsTransfer -Source $url -Destination ('{0}\{1}' -f $DayDir, $file) -ErrorAction $erroraction)
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
                                $info | Out-File -FilePath ($text2 -f $ErrorFile) -Append
                              }
                              $file | Out-File -FilePath ($text2 -f $LogFile) -Append
                            }
                          }
  }
}

catch [Net.WebException]
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

# Check if the m3u file already exists. If so remove it

try
{
  if (Test-Path -Path $PlayList -IsValid) 
  {
    Remove-Item -Path $PlayList -ErrorAction Stop
  }
}
# NOTE: When you use a SPECIFIC catch block, exceptions thrown by -ErrorAction Stop MAY LACK
# some InvocationInfo details such as ScriptLineNumber.
# REMEDY: If that affects you, remove the SPECIFIC exception type [System.Management.Automation.ItemNotFoundException] in the code below
# and use ONE generic catch block instead. Such a catch block then handles ALL error types, so you would need to
# add the logic to handle different error types differently by yourself.
catch [Management.Automation.ItemNotFoundException]
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

#Get-FileMetaData -folder $DayDir | Select-Object -Property Name, album, Title
# Make a podcast.m3u file

Write-Output -InputObject 'Generating playlist' 
Get-ChildItem -Path $DayDir  | Where-Object  { !$_.PsIsContainer -and $_.Extension -ne '.m3u' } | ForEach-Object -Process {$_.Name | Out-File -FilePath $PlayList -Encoding ascii -Append }


# TODO make an option for downloading OR streaming
# Check if player exists
Write-Output -InputObject $Player
if ((Test-Path -Path $Player -IsValid)) 
{
  $vlcplayer = Get-Process -Name vlc -ErrorAction SilentlyContinue
  if (!  $vlcplayer) {
    Start-Process -FilePath $Player -ArgumentList $Playlist 
   }
}
Write-Output -InputObject 'End'
# Todo Playlist still contains invalid chars. HOW?? Fixed! -encoding ascii
#====================================================================================
# End
#====================================================================================
