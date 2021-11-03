If (Test-Path -Path Cert:\LocalMachine\CA\22BBE981F0694D246CC1472ED2B021DC8540A22F)
        {
            $Cert = Get-Item -Path Cert:\LocalMachine\CA\22BBE981F0694D246CC1472ED2B021DC8540A22F
            $Store = Get-Item -Path Cert:\LocalMachine\CA
            $Store.open("ReadWrite")
            $Store.Remove($Cert)
            $Store.Close()
            $Untrusted = Get-Item -Path Cert:\LocalMachine\Disallowed
            $Untrusted.open("ReadWrite")
            $Untrusted.add($Cert)
            $Untrusted.close()

            $Result = "Moved $Cert to the Disallowed Folder."
            $Result | Out-File -Verbose C:\Users\1180219788A\Desktop\Cert_Success.txt -Append -Force
    }
Else
    {
        $Result = "The certificate is not in the AuthRoot Folder."
        $Result | Out-File -Verbose C:\Users\1180219788A\Desktop\Cert_Fail.txt -Append -Force
    }
