#############JARVIS##################################################
###################Animations########################################



Function Loading_Art($Seconds){
                                $I= 0
                                While($I -lt $Seconds){
                                Write-host     "#                #"
                                start-sleep -Milliseconds 50
                                Write-host     " #              # "
                                start-sleep -Milliseconds 50
                                Write-host     "  #            #  "
                                start-sleep -Milliseconds 50
                                Write-host     "   #          #   "
                                start-sleep -Milliseconds 50
                                Write-host     "    #        #    "
                                start-sleep -Milliseconds 50
                                Write-host     "     #      #     "
                                start-sleep -Milliseconds 50
                                Write-host     "      #    #      "
                                start-sleep -Milliseconds 50
                                Write-host     "       #  #       "
                                start-sleep -Milliseconds 50
                                Write-host     "        ##        "
                                start-sleep -Milliseconds 50
                                Write-host     "        ##        "
                                start-sleep -Milliseconds 50
                                Write-host     "       #  #       "
                                start-sleep -Milliseconds 50
                                Write-host     "      #    #      "
                                start-sleep -Milliseconds 50
                                Write-host     "     #      #     "
                                start-sleep -Milliseconds 50
                                Write-host     "    #        #    "
                                start-sleep -Milliseconds 50
                                Write-host     "   #          #   "
                                start-sleep -Milliseconds 50
                                Write-host     "  #            #  "
                                start-sleep -Milliseconds 50
                                Write-host     " #              # "
                                start-sleep -Milliseconds 50
                                Write-host     "#                #"
                                start-sleep -Milliseconds 50
                                                      $I++}
                                          
                                 }
######GREETINGS FOR TIME OF DAY FUNCTION##############################################################
Function TOD_Greeting(){
                            $TOD = (get-date).TimeOfDay
                            $Hour = $TOD.Hours
                            
                            If($Hour -lt 12){
                                                    $greeting = "Good Morning Aaron! It's me! Girl Jarvis. I hope you had a great night."
                                                    
                                                    }
                            ElseIf($Hour -ge 12 -and $Hour -le 17){

                                                                    $greeting = "Good Afternoon Aaron! Girl Jarvis here. I am happy to see you."
                                                                   
                                                                   }
                            ElseIf($Hour -gt 17 -and $Hour -le 19){

                                                                      $greeting = "Good Evening Aaron! Girl Jarvis at your service. It's almost time to go home!"
                                                                              
                                                                              }
                            Else{ $greeting = "Hello Aaron!  You are here late and girl jarvis is getting sleepy. Don't work too hard!" } 

                            $greeting
                            Speak_It($greeting)
                            }

######SPEECH FUNCTION###Speech Engine - .net Microsoft Native########################################

##Speaking on Windows Machines######################################################################
Function Speak_It([string]$text){
                                    Add-Type -AssemblyName System.speech
                                    $speak = New-Object System.Speech.Synthesis.SpeechSynthesizer
                                    $speak.SelectVoice('Microsoft Zira Desktop')
                                    $speak.Speak($text)  
                            }

###################################################################################################

######################################################################################################

####OUTLOOK#######

#Get current unread messages  ....and maybe read them to me
function Start_Outlook(){
$TTS_Starting = "I am now starting outlook for you. It should just be a moment."
Speak_it($TTS_Starting)
Start-Process -FilePath "C:\Program Files (x86)\Microsoft Office\Office15\OUTLOOK.EXE" -LoadUserProfile

Loading_Art(30)
###Get Unread Emails
Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook"
$Outlook = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNameSpace("MAPI")
$Inbox = $NameSpace.Folders.Item(2).Folders |where {($_.FolderPath) -eq "\\aaron.simonds@us.af.mil\Inbox"}
$UI_Num = $Inbox.UnReadItemCount 

$UI_string = "You have $UI_Num unread emails in your inbox. Check them at your leisure."

###
####Custome Searches#####Secuirty anti-virus won't let me read the body because of the com interface.
####From Tracie Oster####
####From Michael Esquivias###
####From Brandon Devault#####
 

Speak_It($UI_string)

}
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
#####Open Chrome Stuff######################
function Chrome_Stuff(){
$web1 = "mail.yahoo.com"
$UI_1 = "your yahoo mail"
$web2 = "https://music.amazon.com/my/playlists/Cleaning/304c8a15-bb04-4809-bded-e68eb6ab6782"
$UI_2 = "or play some music"
$other_things = $UI_1 + $UI_2 

start-process "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" -ArgumentList $web1,$web2 -WindowStyle Maximized


$UI_string = "I will go ahead and open up your chrome items for you. Please take a look at $other_things, as I prepare the rest of your desktop."

Speak_It($UI_string)
}

Function DoneFor_Today()
{
    Loading_Art(4)

    ###Audios Versioned by Time of Day


    ###Shut Down All processes
    CloseOpen_Proc


    ###Restart Computer
    
    Restart-Computer -wait 60

    Speak_It("Your computer is set to restart in 60 seconds.")

}
Function First_Login(){

    Loading_Art(4)

    TOD_Greeting

    Chrome_Stuff

    Loading_Art(8)

    Start_Outlook

    $UI_String = "I am done for now.  But if there is anything else you need let me know.  Lady Jarvis, OUT!"
    Speak_It($UI_String)
}
