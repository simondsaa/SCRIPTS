######SPEECH FUNCTION###Speech Engine - .net Microsoft Native########################################

##Speaking on Windows Machines######################################################################
Function Speak_It([string]$text){
                                    Add-Type -AssemblyName System.speech
                                    $speak = New-Object System.Speech.Synthesis.SpeechSynthesizer
                                    $speak.SelectVoice('Microsoft Zira Desktop')
                                    $speak.Speak($text)  
                            }

###################################################################################################

