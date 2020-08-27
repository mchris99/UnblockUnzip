# PowerShell script to unzip files from a specific (unzipped) folder into a new folder.
# Requires: Source folder filled with zipped folders.
# Involves: Unblocking all zipped folders in source before unzipping them to a destination.

# Type the file name of this script in powershell to execute.
# To use this script, your execution policy must allow scripts to be run and you should run powershell as an administrator.
# More information on running powershell scripts:
# https://www.itprotoday.com/powershell/how-run-powershell-script

# Communicate with user: 
Write-Host "This program takes a single folder with multiple zipped files and unblocks/unzips them into a specified folder, preserving the names of each zipped folder."
Write-Host "If the destination folder does not exist, it will be created."
Write-Host "If a certain file already exists in the destination folder, unzipping will be skipped for that particular file."
Write-Host ""

# Receive paths as input:
# Example source folder: C:\Users\USERNAME\Desktop\ZippedFolders\
$source = Read-Host -Prompt "Enter the full path of the folder containing the zipped files"
# Example destination folder: C:\Users\USERNAME\Desktop\UnzippedFolders\
$dest = Read-Host -Prompt "Enter the full path of the folder you wish the files to be unzipped to"
# Check to make sure that both paths end with "\":
if ($source -notmatch '\\$') {
    $source += '\'
}
if ($dest -notmatch '\\$') {
    $dest += '\'
}

# Unblock all files in source:
Get-ChildItem $source -Recurse | Unblock-File

# Unzip each folder into a new folder of the same name inside dest:
# https://powershell.org/forums/topic/unzip-multiple-zips-to-multiple-folders-from-a-single-source-folder/
$zip = Get-ChildItem -Path $source -Filter *.zip
foreach ($z in $zip){
    Expand-Archive -Path $z.FullName -DestinationPath $dest$($z.BaseName) -Verbose
}