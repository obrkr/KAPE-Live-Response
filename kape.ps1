$zipLocation = "C:\ProgramData\Microsoft\Windows Defender Advanced Threat Protection\Downloads\kape.zip"
$workingDirectory = "C:\ProgramData\kape"

# Check if the extraction directory exists, if not, create it
If(!(test-path -PathType container $workingDirectory))
{
      New-Item -ItemType Directory -Path $workingDirectory
}

# Unzip the file using the built-in ComObject Shell.Application
$shell = New-Object -ComObject Shell.Application
$zipFile = $shell.NameSpace($zipLocation)
$destination = $shell.NameSpace($workingDirectory)
$destination.CopyHere($zipFile.Items())

# Wait for the extraction process to complete 
while ($destination.Items().Count -ne $zipFile.Items().Count) {
    Start-Sleep -Seconds 1
}

# Execute the kape.exe with the given parameters
$kapeExe = "C:\kape\kape.exe"
$args = "--tsource C:\ --tdest C:\ProgramData\kape\output --tflush --target !SANS_Triage --zip kapeoutput"
Start-Process -FilePath $kapeExe -ArgumentList $args -Wait
