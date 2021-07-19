# Deployment Script for RingCentral.
# Install/Uninstall RC Versions 10.0 & 10.1

# Global Admin Credentials ---------------------------------------------------------------------------
$GlobalUN = "Globaltrax\ejacobsen"
$GlobalPW =  Get-Content C:\Scripts\ejacobsen@omni-healthcare.com.txt | ConvertTo-SecureString -Force
$GlobalCred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $GlobalUN, $GlobalPW
#----------------------------------------------------------------------------------------------

Function Deploy-EXE
{
    param(
        [Parameter(Mandatory=$True)][PSCredential]$Credential,
        [Parameter(Mandatory=$True)][string]$Hostname,
        [Parameter(Mandatory=$True)][string]$DownloadLink,
        [Parameter(Mandatory=$True)][string]$FileName,
        [Parameter(Mandatory=$False)][string]$Arguments
    )

    # Create Session and use MSIEXEC to execute command on remote computer.
    $Sesh = New-PSSession -ComputerName $Hostname -Credential $Credential
    Invoke-Command -Session $Sesh {
        # Check C:\Temp path.
        $DownloadLink = $Args[0]; $Filename = $Args[1]; $Arguments = $Args[2]
        $dir = "C:\Temp\packages"
        If( -Not (Test-Path $dir)) {mkdir $dir}
        
        # Download File.
        $webClient = New-Object System.Net.WebClient
        $url = $DownloadLink
        $file = "$($dir)\$Filename.exe"
        $webClient.DownloadFile($url,$file)
       
        # Run Exe with Arguments.
        If($Arguments -ne "") { Start-Process "$($dir)\$Filename.exe" -ArgumentList $Arguments -Wait -PassThru }
        Else { Start-Process "$($dir)\$Filename.exe" -Wait -PassThru }
        $ENV:COMPUTERNAME + ": Success"
    } -ArgumentList $DownloadLink, $Filename, $Arguments
    $Sesh | Remove-PSSession
}


Deploy-EXE -Credential (Get-Credential) `
-Hostname (Read-Host 'Hostname: ') `
-DownloadLink 'https://dl.sharefile.com/cfo' `
-FileName 'CitrixFilesForOutlook-v6.5.12.1' `
-Arguments '/S'
