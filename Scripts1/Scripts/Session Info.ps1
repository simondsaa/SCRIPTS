#-----------------------------------------------------------------------------------------#
#                                  Written by SrA David Roberson                          #
#                                  Tyndall AFB, Panama City, FL                           #
#                                     Created 01 July 2014                                #
#-----------------------------------------------------------------------------------------#
$Computers = Get-Content C:\Users\1252862141.adm\Desktop\Scripts1\Pop.txt
ForEach ($Computer in $Computers)
{
If (Test-Connection $Computer -quiet -BufferSize 16 -Ea 0 -count 1)
    { 
    Get-PSSession -ComputerName $Computer -ErrorAction SilentlyContinue
    "`n"
    } 
Else
    {
    Write-Host -ForegroundColor Red "$Computer is not available"
    "`n"
    }  
}