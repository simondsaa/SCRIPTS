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

