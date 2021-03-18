***** MICROSOFT STANDALONE UPDATER *****
***** By Eric Stevens *****
***** Last updated: 03/16/2021 *****

This script was created to bring standalone Windows 10 and Windows Server systems up-to-date with the latest monthly security patches released by Microsoft. Please note that all updates are cumulative--you won't need to install previous security patches if you've missed them, the current monthly updates in these directories supersede older ones.

You can also run this script from a fileshare or other method, such as PDQ, for remotely updating systems on a P2P or domain network.


Currently Supported Products:

- Windows 10 OS Versions 1607, 1803, 1909, 20H2 cumulative updates and servicing stack updates
- Windows Server OS versions 2012r2, 2016, 2019 cumulative updates and servicing stack updates
- .NET framework 3.5, 4.7.2, 4.8 updates
- Adobe Flash security updates
- MS Office 2013 64-bit security updates
- MS Office 2016 64-bit security updates
- MS Edge latest stable-channel release
- Latest VC redistributables
- Latest WMF updates
- Latest Microsoft-issued certificates
- Latest Microsoft-issued certificate revocation lists


*** Standalone Directions ***

Step 1: Copy Powershell_Installation directory to local drive or fileshare.

Step 2: Inside the copied directory, right-click start.cmd --> Run as Administrator.

Step 3: A cmd window will appear and display current patching progress. Prompt will notify you when patching is complete.


*** PDQ Deploy Directions ***

Step 1: Copy Powershell_Installation directory into your PDQ repository.

Step 2: Inside the directory, open (don't run) the Microsoft_Automated_Patching.ps1 file with notepad or a similar text editor.

Step 3: On line 7, change $PDQ_Patch_Repo from "$null" to the path to the directory. Example: "\\PATH_TO_PDQ_REPO\Powershell_Installation"

Step 4: Create a 1-step PDQ deployment package that runs the Microsoft_Automated_Patching.ps1 script within the directory.

Step 5: Deploy it to any machine.


*** Optional: Patch Validation ***

At the end of the script on line 224, you'll find a "#Verify-Updates" comment. If you uncomment this line (take out the # symbol) then an updates.csv file will be created in the Powershell_Installation directory that lists the installed Windows OS Security Patches. You'll need to change the listed date on line 118 to the date you installed patches, and be please aware that the latest LCU might not appear in this file unless the machine is rebooted to complete the installation. This .CSV file will append any new data whenever this function is run.