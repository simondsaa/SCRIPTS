.\MultiThread_x64 -f Workstation_List.txt -c "cscript ClientHealth.vbs /machine:%S /repair" -t 50

cscript .\BuildUpdatedTargetList.vbs /D:C:\Users\1274873341C\Desktop\Desktop\PS_Scripts\4-Rs_Commands_and_Scripts\Scripts\XMLOutput_ClientHealth /F:C:\Users\1274873341C\Desktop\Desktop\PS_Scripts\4-Rs_Commands_and_Scripts\Scripts\Workstation_List.txt

.\MultiThread_x64 -f Workstation_List_NewTargetList.txt -c "cscript ClientHealth.vbs /machine:%S /repair" -t 50




cscript ClientHealth.vbs /File:C:\Users\Public\Documents\!Run_Folder\Computers.txt /Repair

cscript ClientHealth.vbs /ProcessXML_SCCM

copy *.txt File_Name.txt