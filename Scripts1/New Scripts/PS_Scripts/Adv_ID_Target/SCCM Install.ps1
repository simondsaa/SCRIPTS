<#
$computers = gc "C:\Users\1274873341C\Desktop\Desktop\PS_Scripts\Adv_ID_Target\Target_Computers.txt"


foreach($computer in $computers){
    
    $SCCM = New-Object -ComObject UIResource.UIResourceMgr
    $SCCM.GetAvailableApplications() | Select ID, PackageID, PackageName | Where {$_.PackageName -like "INE51750"} | Format-List
    }
    #>

    $SCCM = New-Object -ComObject UIResource.UIResourceMgr
    $SCCM.GetAvailableApplications() | Select ID, PackageID, PackageName | Where {$_.PackageName -like "INE51750"} | Format-List