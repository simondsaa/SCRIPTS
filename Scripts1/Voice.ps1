Invoke-Command -Computername xlwuw-491s40 -ScriptBlock { function Say-Text {
    param ([Parameter(Mandatory=$true, ValueFromPipeline=$true)] [string] $Text)
    [Reflection.Assembly]::LoadWithPartialName('System.Speech') | Out-Null   
    $object = New-Object System.Speech.Synthesis.SpeechSynthesizer 
    $object.Speak($Text) 
}
}
sleep 10
Say-Text "Test."
