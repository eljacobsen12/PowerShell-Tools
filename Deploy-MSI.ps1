# Deployment Script for RingCentral.
# Install/Uninstall RC Versions 10.0 & 10.1

# Global Admin Credentials ---------------------------------------------------------------------------
$GlobalUN = "Globaltrax\ejacobsen"
$GlobalPW =  Get-Content C:\Scripts\ejacobsen@omni-healthcare.com.txt | ConvertTo-SecureString -Force
$GlobalCred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $GlobalUN, $GlobalPW
#----------------------------------------------------------------------------------------------

# Used to uninstall programs under User Profile.
Function Get-UserProfiles
{
    param(
        [Parameter(Mandatory=$True)]
        [string]$ComputerName
    )

    Invoke-Command -ComputerName $ComputerName -Credential $Cred {
        $Profiles = gwmi win32_userprofile | select @{LABEL="LastUsed";EXPRESSION={$_.ConvertToDateTime($_.lastusetime)}}, LocalPath, SID
        $List = New-Object System.Collections.Generic.List[PSObject]
        
        ForEach($Profile in $Profiles) {
            $Properties = [Ordered]@{'LocalPath' = $Profile.LocalPath;
                                     'LastUsed' = $Profile.'LastUsed';
                                     'SID' = $Profile.SID
            }
            $Object = New-Object -TypeName PSObject -Property $Properties
            $List.Add($Object)
        }
        $List | Format-Table
    }
}

Function Deploy-MSI
{
    param(
        [Parameter(Mandatory=$True)]
        [PSCredential]$Credential,
        [string[]]$Computers,
        [string]$Repository,
        [string]$Filename,
        [ValidateSet("Install", "Uninstall", "Repair")]
        [string]$InstallOption,
        [string]$OtherOptions
    )

    ForEach($ComputerName in $Computers) {
        If(Test-Connection -ComputerName $ComputerName -Count 1 -ErrorAction SilentlyContinue) {
            # Create the MSIEXEC string command.
            If($InstallOption -eq 'Install') {$Option = 'i'}
            ElseIf($InstallOption -eq 'Uninstall') {$Option = 'x'}
            ElseIf($InstallOption -eq 'Repair') {$Option = 'f'}

            # Copy the file to the remote Computer using ROBOCOPY.
            If(Test-Path -Path "\\$ComputerName\C$\Windows\Temp\$Filename.msi") {
                 # If file is older than limit, force ROBOCOPY.
                 $CreatedDate = (Get-ChildItem "\\$ComputerName\C$\Windows\Temp\$Filename.msi").CreationTime
                 $limit = (Get-Date).AddDays(-5)
                 If ($CreatedDate -lt $limit) {
                    # Force ROBOCOPY to copy over more recent version of file.
                    robocopy "$Repository" "\\$ComputerName\C$\Windows\Temp" "$Filename.msi" | Out-Null
                 }
            }
            Else { robocopy "$Repository" "\\$ComputerName\C$\Windows\Temp" "$Filename.msi" | Out-Null }

            # Use MSIEXEC to execute command on remote computer.
            $Sesh = New-PSSession -ComputerName $ComputerName -Credential $Credential
            Invoke-Command -Session $Sesh {
                $Filename = $Args[0]
                $Option = $Args[1]
                
                If($Args[2] -ne $null) {msiexec.exe /$Option "C:\Windows\Temp\$Filename.msi" ALLUSERS=1 /qn /norestart $Args[2]}
                Else {msiexec.exe /$Option "C:\Windows\Temp\$Filename.msi" ALLUSERS=1 /qn /norestart}

                # ********* *********** ************ **************** ************** **********
                Start-Sleep -Seconds 2 #Seemed like the session was being close before the command was finished.
                $ENV:COMPUTERNAME + ": Success"
            } -ArgumentList $Filename, $Option, $OtherOptions
            $Sesh | Remove-PSSession
        }
        Else {
            Write-Host "$ComputerName Offline"
        }
    }
}

Deploy-MSI -Credential $GlobalCred -Computers ("WS81") -Repository '\\globaldc1\gpo software\ShareFile' -Filename 'CitrixFilesForWindows_x64_v20.3.28.0' -InstallOption Install
#Deploy-MSI -Credential $GlobalCred -Computers ("WS81") -Repository '\\globaldc1\gpo software\RingCentral\Meetings' -Filename 'RCMeetingsClientSetup-19.4.23042' -InstallOption Install #-OtherOptions "ADDLOCAL=ALL IACCEPTSQLNCLILICENSETERMS=YES"
#Deploy-MSI -Computers ("WS81") -Filename 'RingCentral-18.08.1-x64' -InstallOption Install
#Deploy-MSI -Computers ("WS35") -Filename 'RingCentral Desktop' -InstallOption Install
#Deploy-MSI -Computers ("WS35") -Filename 'FixMe.IT Unattended' -InstallOption Install
#Deploy-MSI -Computers ("NB5","NB6","NB19","NB20","NB21","NB22","WS93","WS31","WS65","NB24","WS17") -Filename 'RingCentral Phone-10.1.2' -InstallOption Install
#Deploy-MSI -Computers (Get-Content "C:\Users\ejacobsen\OneDrive - Global Financial\Scripts\NC Computers.csv") -Filename 'RingCentral Phone-10.1.2' -InstallOption Install
#$Time = Measure-Command { Deploy-MSI -Computers 'WS35' -Filename 'RingCentral Desktop' -InstallOption Uninstall }