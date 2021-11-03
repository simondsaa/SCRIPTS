$Computers = Get-Content C:\Users\1180219788A\Desktop\NCC.txt

ForEach ($Computer in $Computers)
    {

        $vers = Get-WmiObject -ComputerName $Computer -Class Win32_OperatingSystem

        If ($vers.caption -like "*Windows 7*" -and $vers.OSArchitecture -like "64-bit")
            
            {  $Computer | Out-File -FilePath C:\Users\1180219788A\Desktop\Win7x64.txt -Append -Force  }

        ElseIf ($vers.caption -like "*Windows 7*" -and $vers.OSArchitecture -like "32-bit")

            {  $Computer | Out-File -FilePath C:\Users\1180219788A\Desktop\Win7x32.txt -Append -Force  }
        
        Else 
            
            {  $Computer | Out-File -FilePath C:\Users\1180219788A\Desktop\Win10.txt -Append -Force   }
    }