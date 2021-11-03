Synopsis
   SDCTargeting - Gathers report on machines put into Security Groups for SDC Deployment Solutions
   SDCTargeting - Gathers report on all machines that are ready to upgrade to 5.3.1
   **Needs to be ran as administrator
   **Needs RSAT tools including functions like Get-ADComputer
   **The "Pre-Pre-Flight.txt" needs to be renamed to "Pre-Pre-Flight.ps1"

DESCRIPTION
    This script generates a reports on the machines in each of the 2 Security groups:
        GLS_BASE NAME_SDC Servicing Script (Available)
        GLS_BASE NAME_ SDC Servicing Script (Mandatory)

    This script should be ran on a daily basis to check the current status of machines within the security groups.
    
    For the SDC_Servicing groups the report gathers 2 main results:
        Machines that are now on 5.3.1
        Machines that are still on 5.2
            Machines that failed pre-flight checks
            Machines that never recieved advertisment
            Machines that were interrupted during upgrade phase
    
    Legacy Security Groups are currently unavailable, please wait for next updated script.

    For any machine that failed to upgrade, important logs are gathered in host computer to centeralize and process logs to determine reason for failure.

ARGUEMENTS
	BaseName: Enter the base name you want to target EXCLUDING "AFB". Ex. "BaseName:Scott"

	AFCONUS: Enter accordingly to the targeted base. Ex. "AFCONUS:EAST"

	Upgradable_SDC_Servicing_Machines: Yes or No (Y or N) depending on if you would like a list of machines that would fall into this category. Ex. "Upgradable_SDC_Servicing_Machines:Yes"

	Upgradable_Legacy_Machines: Yes or No (Y or N) depending on if you would like a list of machines that would fall into this category. Ex. "Upgradable_Legacy_Machines:Yes"

	SDC_Servicing_Mandatory: Yes or No (Y or N). This will check this security group for eligible machines. Ex. "SDC_Servicing_Mandatory:Yes"

	SDC_Servicing_Available: Yes or No (Y or N). This will check this security group for eligible machines. Ex. "SDC_Servicing_Available:Yes"

	PathToWriteLogs: Enter a full path to the location you would like the .csv file to be created. Ex. "PathToWriteLogs:C:\Users\Administrator\Dekstop"
