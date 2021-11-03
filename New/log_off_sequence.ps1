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