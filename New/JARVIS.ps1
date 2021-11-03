#############JARVIS##################################################

#####################Loading Sequence################################

$location
$operating_system
$time
$profile_name



#Get Today's Appointments


#####Open Internet Explorer Stuff###########




#####Open One Note##########################
Function CloseOpen_Proc()
{
    Get-process  -name "*Chrome*"
    Get-Process  -name "*Outlook*"
    Get-Process  -name "*iexplore*"
    Get-Process  -name "*excel*"


    Get-process |where {($_.SI) -eq "1"}

}







#Make_Conversation(){










Function Laugh(){
                    speak_it("HA HA.  HA HA. HA HA.   Rawr!")

                    }

Function Reactive(){
                    
                    $powerpoint = get-process | ?{$_.name -like 'POWERPNT'}
                    
                    $tanium_c = get-process |?{$_.name -like "TaniumCLient"}
                    $tanium_c.count
           
                    $powerpoint.starttime

                        if($powerpoint - $True){
                                speak_it("Go go powerpoint rangers!..........GO!")
                        }



}

###git info pull
