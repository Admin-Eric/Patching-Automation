### Standalone Microsoft Updater
### Created by Eric Stevens
### Last updated 03/16/2021

## If Using PDQ Deploy, change $PDQ_Patch_Repo to path for the script directory. Example: $PDQ_Patch_Repo = \\Path_to_repo\Powershell_Installation

$PDQ_Patch_Repo = $null

if ($PDQ_Patch_Repo -ne $null)
  {
    $PathToRepo = $PDQ_Patch_Repo
  }
  else {$PathToRepo = $PSScriptRoot}

<#
Subdirectories are assigned to each variable. The system environment variable is used to determine which directory OS
standalone patches will be pulled from. All other directories contain standalone installers for various products
that correlate to the folder name.
#>

$env:SEE_MASK_NOZONECHECKS = 1
$winversion = [System.Environment]::OSVersion.Version.Build
$officeupdates = Get-ChildItem "$PathToRepo\Office\*" -include "*.msi","*.exe"
$redistributables = Get-ChildItem "$PathToRepo\Redistributables\*" -include "*.msi","*.exe"
$updatefolder =  Get-ChildItem "$PathToRepo\$winversion\*" -include "*.msu"
$SSUfolder =  Get-ChildItem "$PathToRepo\$winversion\SSU\*" -include "*.msu"
$certificates = Get-ChildItem -path "$PathToRepo\Certificates\*" -include "*.crt","*.cer"
$revokelist = Get-ChildItem -Path "$PathToRepo\CRL\*" -include "*.crl"
$miscupdates = Get-ChildItem -Path "$PathToRepo\Misc Installers\*"

### This is the function that installs updates for only those redistributables which are already on the machine.

function Install-VCs {
    $VC2019x86 = Get-WmiObject -Class Win32_Product -Filter "name LIKE '%Visual C++ 2019 x86%'"
    $VC2019x64 = Get-WmiObject -Class Win32_Product -Filter "name LIKE '%Visual C++ 2019 x64%'"
    $VC2013x86 = Get-WmiObject -Class Win32_Product -Filter "name LIKE '%Visual C++ 2013 x86%'"
    $VC2013x64 = Get-WmiObject -Class Win32_Product -Filter "name LIKE '%Visual C++ 2013 x64%'"
    $VC2012x86 = Get-WmiObject -Class Win32_Product -Filter "name LIKE '%Visual C++ 2012 x86%'"
    $VC2012x64 = Get-WmiObject -Class Win32_Product -Filter "name LIKE '%Visual C++ 2012 x64%'"
    $VC2010x86 = Get-WmiObject -Class Win32_Product -Filter "name LIKE '%Visual C++ 2010  x86%'"
    $VC2010x64 = Get-WmiObject -Class Win32_Product -Filter "name LIKE '%Visual C++ 2010  x64%'"
    $VC2008x86 = Get-WmiObject -Class Win32_Product -Filter "name LIKE '%Visual C++ 2008 x86%'"
    $VC2008x86 = Get-WmiObject -Class Win32_Product -Filter "name LIKE '%Visual C++ 2008 x64%'"

  if ($VC2019x86 -ne $null)
    {
        Write-Host "VC Redistributable version 2015-2019 32-bit found. Updating..."
        Start-Process -FilePath "$PathToRepo\Redistributables\vcredist2019_x86.exe" -ArgumentList "/quiet /norestart" -wait -PassThru
    }
    else {write-Host "VC Redistributable version 2015-2019 32-bit not found. Skipping..."}
  if ($VC2019x64 -ne $null)
    {
        Write-Host "VC Redistributable version 2015-2019 64-bit found. Updating..."
        Start-Process -FilePath "$PathToRepo\Redistributables\vcredist2019_x64.exe" -ArgumentList "/quiet /norestart" -wait -PassThru
    }
    else {write-Host "VC Redistributable version 2015-2019 64-bit not found. Skipping..."}
  if ($VC2013x86 -ne $null)
    {
        Write-Host "VC Redistributable version 2013 32-bit found. Updating..."
        Start-Process -FilePath "$PathToRepo\Redistributables\vcredist2013_x86.exe" -ArgumentList "/quiet /norestart" -wait -PassThru
    }
    else {write-Host "VC Redistributable version 2013 32-bit not found. Skipping..."}
  if ($VC2013x64 -ne $null)
    {
        Write-Host "VC Redistributable version 2013 64-bit found. Updating..."
        Start-Process -FilePath "$PathToRepo\Redistributables\vcredist2013_x64.exe" -ArgumentList "/quiet /norestart" -wait -PassThru
    }
    else {write-Host "VC Redistributable version 2013 64-bit not found. Skipping..."}
  if ($VC2012x86 -ne $null)
    {
        Write-Host "VC Redistributable version 2012 32-bit found. Updating..."
        Start-Process -FilePath "$PathToRepo\Redistributables\vcredist2012_x86.exe" -ArgumentList "/quiet /norestart" -wait -PassThru
    }
    else {write-Host "VC Redistributable version 2012 32-bit not found. Skipping..."}
  if ($VC2012x64 -ne $null)
    {
        Write-Host "VC Redistributable version 2012 64-bit found. Updating..."
        Start-Process -FilePath "$PathToRepo\Redistributables\vcredist2012_x64.exe" -ArgumentList "/quiet /norestart" -wait -PassThru
    }
    else {write-Host "VC Redistributable version 2012 64-bit not found. Skipping..."}
  if ($VC2010x86 -ne $null)
    {
        Write-Host "VC Redistributable version 2010 32-bit found. Updating..."
        Start-Process -FilePath "$PathToRepo\Redistributables\vcredist2010_x86.exe" -ArgumentList "/quiet /norestart" -wait -PassThru
    }
    else {write-Host "VC Redistributable version 2010 32-bit not found. Skipping..."}
  if ($VC2010x64 -ne $null)
    {
        Write-Host "VC Redistributable version 2010 64-bit found. Updating..."
        Start-Process -FilePath "$PathToRepo\Redistributables\vcredist2010_x64.exe" -ArgumentList "/quiet /norestart" -wait -PassThru
    }
    else {write-Host "VC Redistributable version 2010 64-bit not found. Skipping..."}
  if ($VC2008x86 -ne $null)
    {
        Write-Host "VC Redistributable version 2008 32-bit found. Updating..."
        Start-Process -FilePath "$PathToRepo\Redistributables\vcredist2008_x86.exe" -ArgumentList "/quiet /norestart" -wait -PassThru
    }
    else {write-Host "VC Redistributable version 2008 32-bit not found. Skipping..."}
  if ($VC2008x64 -ne $null)
    {
        Write-Host "VC Redistributable version 2008 64-bit found. Updating..."
        Start-Process -FilePath "$PathToRepo\Redistributables\vcredist2008_x64.exe" -ArgumentList "/quiet /norestart" -wait -PassThru
    }
    else {write-Host "VC Redistributable version 2008 64-bit not found. Skipping..."}
}

## This optional function is applied at the last step of the script. Uncomment it if you'd like to create the update.csv script to view installed OS security patches.

function Verify-Updates {

  Write-host "
  Running optional validation function. Check updates.csv for updated security OS patches. WARNING: updates.csv might not be accurate unless system is rebooted to finish LCU installation." -ForegroundColor Green

  Get-HotFix | 

  ## Replace the date below with the month/day/year of the day you patched the network. 

  Where-Object InstalledOn -GE "02/08/2021" | 
  Select-Object PSComputerName, Description, HotFixID, InstalledBy, InstalledOn | 
  Export-Csv -Path $PathToRepo\updates.csv -NoTypeInformation -Append
}

write-host "
Initializing Microsoft Standalone Updater, last updated 03/16/2021.
Windows version $winversion patches installing.
" -ForegroundColor Green

### The latest Windows SSU update is installed first to ensure other OS updates are installed correctly.

Write-Host "
Installing latest Servicing Stack Update.
" -ForegroundColor Green

foreach ($SSU in $SSUfolder)

  {
    $SSUlocation = $SSU.fullname
    Start-Process -FilePath "wusa.exe" -ArgumentList "$SSUlocation /quiet /norestart" -wait -PassThru
  }

### Other included OS updates are installed afterwards.

write-host "
Installing latest Cumulative, Adobe Flash and .NET framework updates.
" -ForegroundColor Green


foreach ($update in $updatefolder)

  {
    $updatelocation = $update.fullname
    Start-Process -FilePath "wusa.exe" -ArgumentList "$updatelocation /quiet /norestart" -wait -PassThru
  }

### Office 2016/2013 64-bit patches are installed after OS updates.

write-host "
Installing Office 2013/2016 updates.
" -ForegroundColor Green


foreach ($office in $officeupdates)

	{
		Start-Process -FilePath ($office.fullname) -ArgumentList "/quiet /norestart" -wait -PassThru
	}

### Redistributables are installed next.

write-host "
Installing VC Redistributables.
" -ForegroundColor Green

Install-VCs

<# 
A Miscellaneous folder is included in the powershell script directory in case the user wants
to add additional updates/applications (such as firefox, chrome, etc.) to the update process. 
By default, this directory is left empty and will be up to the admin to decide on if or when
other applications can be added.
#>

write-host "
Installing Miscellaneous updates and applications.
" -ForegroundColor Green


foreach ($misc in $miscupdates)

	{
		Start-Process -FilePath ($misc.fullname) -ArgumentList "/quiet /norestart" -wait -PassThru
	}

### After installations, any new MS certificates are added to the local machine.

Write-Host "
Importing latest Microsoft certificates.
" -ForegroundColor Green

foreach ($cert in $certificates)

  {
    Import-Certificate -filepath ($cert.fullname) -CertStoreLocation Cert:\LocalMachine\Root
  }

### The certificate revocation list is updated on the local machine with any additional certificates.

Write-Host "
Importing latest Microsoft certificate revocation lists.
" -ForegroundColor Green


foreach ($revoke in $revokelist)

  {
    certutil.exe -addstore CA ($revoke.fullname)
  }

<#
Optional Validation function applied last. Warning: Validation might not be accurate if system has not been rebooted since updated. 
Reboot is required to finish installation of LCU. Uncomment the next Verify-Updates line to run the function.
#>

#Verify-Updates

Write-host "

Updates finished. Please reboot to complete installation." -ForegroundColor Green