#####Open Chrome Stuff######################
function Chrome_Stuff(){
$web1 = "mail.google.com"
$UI_1 = "your google mail, "
$web2 = "https://cyber-defense.sans.org/blog"
$web5 = "https://isc.sans.edu/diary.html"
$UI_2 = "new cyber security information from sans, "
$web3 = "calendar.google.com/calendar/r/month"
$UI_3 = "personal calendar,"
$web4 = "https://music.amazon.com/home?ie=UTF8&ref_=sv_dmusic_7"
$UI_4 = "or play some music"
$other_things = $UI_1 + $UI_2 + $UI_3 + $UI_4

start-process "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" -ArgumentList $web1,$web2,$web3,$web4,$web5 -WindowStyle Maximized


$UI_string = "I will go ahead and open up your chrome items for you. Please take a look at $other_things, as I prepare the rest of your desktop."

Speak_It($UI_string)
}

