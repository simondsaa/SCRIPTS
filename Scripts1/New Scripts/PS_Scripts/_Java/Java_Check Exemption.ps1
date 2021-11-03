$Computers = Get-Content C:\Users\1274873341C\Desktop\Desktop\PS_Scripts\_Java\Vulnerable_Systems.txt
#Remove-Item "C:\Users\1274873341C\Desktop\Desktop\PS_Scripts\_Java\Java_*.txt"


#This section checks exemption group membership for all of the Java targets identified from ACAS

    ForEach ($Computer in $Computers)
    {

                   
        $Groups =   Get-ADPrincipalGroupMembership ((Get-ADComputer $Computer).DistinguishedName)
        
    
        If (($Groups).Name -like "Java Push Exemption*")
        {
            Write-Host "$Computer is a member of Java Exemption Group"
            "$Computer" | Out-File C:\Users\1274873341C\Desktop\Desktop\PS_Scripts\_Java\Java_Exempt.txt -Append -Force
        }
    
        Else
        {
            Write-Host "$Computer can be patched"
            "$Computer" | Out-File C:\Users\1274873341C\Desktop\Desktop\PS_Scripts\_Java\Java_Patch.txt -Append -Force
        }
    }


#This section gathers the computers not within the exemption group and checks their OS architecture for installing either 64-bit or 32-bit Java

    $targetcomputers = Get-Content C:\Users\1274873341C\Desktop\Desktop\PS_Scripts\_Java\Java_Patch.txt
    
    foreach ($targetcomputer in $targetcomputers)
    {
        $OSInfo = Get-Wmiobject Win32_OperatingSystem -ComputerName $targetcomputer -ErrorAction SilentlyContinue
        If (($OSInfo).OSArchitecture -eq "64-Bit")
        {
            Write-host "$targetcomputer is 64 Bit"
            "$targetcomputer" | Out-File C:\Users\1274873341C\Desktop\Desktop\PS_Scripts\_Java\Java_64.txt -Append -Force       
        }
        
        If ($OSInfo.OSArchitecture -eq "32-Bit")
        {
            Write-host "$targetcomputer is 32 Bit"
            "$targetcomputer" | Out-File C:\Users\1274873341C\Desktop\Desktop\PS_Scripts\_Java\Java_86.txt -Append -Force 
        }
    }

<#This section performs the psexec installation of Java for 64-bit systems

    $64bit_computers = Get-Content C:\Users\1274873341C\Desktop\Desktop\PS_Scripts\_Java\Java_64.txt

    foreach ($64bit_computer in $64bit_computers)
    {
        if (Test-Connection -ComputerName $64bit_computer -quiet)
            {
                
                cd "C:\Users\1274873341.adm\Desktop\Program Scripts\Shockwave"
                & psexec \\$64bit_computer -h -d -c -f "C:\Users\1274873341.adm\Desktop\Program Scripts\Java\CommandInstall64.bat" C:

            }
        else
            {
                "$64bit_computer is not online"
            }
    }


#This section performs the psexec installation of Java for 32-bit systems

    $32bit_computers = Get-Content \\xlwu-fs-05pv\Tyndall_PUBLIC\Patching\Java\Java_32.txt

    foreach ($32bit_computer in $32bit_computers)
    {
        if (Test-Connection -ComputerName $32bit_computer -quiet)
            {
                
                cd "C:\Users\1274873341.adm\Desktop\Program Scripts\Shockwave"
                & psexec \\$32bit_computer -h -d -c -f "C:\Users\1274873341.adm\Desktop\Program Scripts\Java\CommandInstall32.bat" C:

            }
        else
            {
                "$32bit_computer is not online"
            }
    }
    #>