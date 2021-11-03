cd 'C:\Program Files (x86)\Microsoft Office\Office16'

Import-Module .\Microsoft.Lync.Model.dll
$Client = [Microsoft.Lync.Model.LyncClient]::GetClient()
$Conversation = $client.ConversationManager.AddConversation()
$person=$client.ContactManager.GetContactByUri('person@domain.com')
$conversation.AddParticipant($person)
 
 
Get-EventSubscriber|Unregister-Event
  # For each participant in the conversation
  $conversation.Participants | Where { !$_.IsSelf } | foreach {
    Register-ObjectEvent -InputObject $_.Modalities[1] -EventName "InstantMessageReceived" -SourceIdentifier "person $i" -action { 
    $global:conv = $event
    $msg = $conv.SourceEventArgs.Text.trim()
    write-host $msg
    switch -Wildcard ($msg) {
     "What*" {$Conversation.Modalities['InstantMessage'].BeginSendMessage((Invoke-Expression $msg.split()[-1]), {}, 0)}
     "Hello" {$Conversation.Modalities['InstantMessage'].BeginSendMessage("Hello human", {}, 0)}
     "stupid robot" {$Conversation.Modalities['InstantMessage'].BeginSendMessage("Humanity is overrated", {}, 0)}
    }
    }
    $i++
   }