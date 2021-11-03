# Filter to define note frequencies.
Filter n {Switch($_){C{262}D{294}E{330}F{349}G{392}A{440}}}

# Notes for Twinkle, Twinkle, Little Star.
$t="CCGGAAGFFEEDDCGGFFEEDGGFFEEDCCGGAAGFFEEDDC"

# Run through each phrase in the song.
1..6|%{
    # Play first six notes as quarter notes.
    $t[0..5]|n|%{[console]::beep($_,600)}
    # Play seventh note as half note.
    $t[6]|n|%{[console]::beep($_,1200)}
    # Left-shift $t by 7 notes.
    $t=$t.SubString(7)
}