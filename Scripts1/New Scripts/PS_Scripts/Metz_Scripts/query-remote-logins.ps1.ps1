Param($computername)

$events = "4648","4624"



Function Query-remote-logins($computername, $eventID)
{
$computername = "XLWUW-421NKX"
$username = (gwmi win32_computersystem -computername $computername).username

$query = @"

<QueryList>

  <Query Id="0" Path="Security">

    <Select Path="Security">*[System[Provider[@Name='Microsoft-Windows-Security-Auditing']

    and (EventID=$eventID)]]

    and *[EventData[Data[@Name='TargetUsername'] != "$computername`$"]]
    and *[EventData[Data[@Name='TargetUserName'] != "$username"]]
    

</Select>

  </Query>

</QueryList>

"@

get-winevent -computername $computername -filterxml $query
}

Foreach($eventID in $events){Query-remote-logins $computername $eventID}





