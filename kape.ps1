$KapeZipLocation = "C:\ProgramData\Microsoft\Windows Defender Advanced Threat Protection\Downloads\kape.zip"
$WorkingDirectory = "C:\ProgramData\kape"

$H = @"
_   __  ___  ______ _____   _____ _____ _      _      _____ _____ _____ ___________ 
| | / / / _ \ | ___ \  ___| /  __ \  _  | |    | |    |  ___/  __ \_   _|  _  | ___ \
| |/ / / /_\ \| |_/ / |__   | /  \/ | | | |    | |    | |__ | /  \/ | | | | | | |_/ /
|    \ |  _  ||  __/|  __|  | |   | | | | |    | |    |  __|| |     | | | | | |    / 
| |\  \| | | || |   | |___  | \__/\ \_/ / |____| |____| |___| \__/\ | | \ \_/ / |\ \ 
\_| \_/\_| |_/\_|   \____/   \____/\___/\_____/\_____/\____/ \____/ \_/  \___/\_| \_|
                                                                                     
                                                                                                                                
"@
Write-Host $H

# Check if the extraction directory exists, if not, create it
if (-not (Test-Path -Path $WorkingDirectory -PathType Container)) {
    New-Item -ItemType Directory -Path $WorkingDirectory -Force | Out-Null
}
Write-Host "Unzipping Kape..."
# Unzip the file using the built-in ComObject Shell.Application
$shell = New-Object -ComObject Shell.Application
$zipFile = $shell.NameSpace($KapeZipLocation)
$destination = $shell.NameSpace($WorkingDirectory)
$destination.CopyHere($zipFile.Items())
Write-Host "Unzipping Complete.."
# Wait for the extraction process to complete 
while ($destination.Items().Count -ne $zipFile.Items().Count) {
    Start-Sleep -Seconds 1
}
Write-Host "Kape Collection running..."
# Execute the kape.exe with the given parameters
$KapeExecutable = "C:\ProgramData\kape\kape.exe"
$Args = "--tsource C:\ --tdest C:\ProgramData\kape\output --tflush --target !SANS_Triage --zip kapeoutput"
Start-Process -FilePath $KapeExecutable -ArgumentList $Args -Wait
$F = @"
##############################################
#            COLLECTION COMPLETE             #
#            Collect with getfile            #
#              File Located at:              #
#        C:\ProgramData\kape\output\         #
##############################################
"@
Write-Host $F
