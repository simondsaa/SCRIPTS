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
$Inbox = $NameSpace.Folders.Item(2).Folders |where {($_.FolderPath) -eq "\\aaron.rosenmund@us.af.mil\Inbox"}
$UI_Num = $Inbox.UnReadItemCount 

$UI_string = "You have $UI_Num unread emails in your inbox. Check them at your leisure."

###
####Custome Searches#####Secuirty anti-virus won't let me read the body because of the com interface.
####From Tracie Oster####
####From Michael Esquivias###
####From Brandon Devault#####
 

Speak_It($UI_string)

}
