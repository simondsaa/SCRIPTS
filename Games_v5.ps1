<#

Written by I.T Delinquent so no stealin' !
Be a champ, visit my website :) mharwood.uk    
Currently version 5, I think...

#>

#SETTING SECRET KEY FOUND TO FALSE AND ASSIGNING RANDOM COLORS
$Global:secretkeyfound = $false
$global:found_secret_key_random_colours = "green","yellow","cyan"

#CHANGING THE POWERSHELL WINDOW SIZE FOR GAMES
function change_screen_size_for_games{
    $pshost = Get-Host
    $pswindow = $pshost.UI.RawUI
    $newsize = $pswindow.WindowSize   
    $newsize = $pswindow.buffersize
    $newsize.height = 22
    $newsize.width = 65
    $pswindow.buffersize = $newsize
    $pswindow.WindowTitle = "Games Galore!"
}

change_screen_size_for_games

#FAKE LOADING SCREEN
function fake_loading_screen{
    $pshost = Get-Host 
    $pswindow = $pshost.Ui.RawUI
    $pswindow.WindowTitle = "Games Galore! - Loading"
    Clear-Host
    Write-Host @"










    0% [ ##########                                    ] 100 %
"@
    Start-Sleep -Milliseconds 350
    Clear-Host
    Write-Host @"










    0% [ #####################                         ] 100%
"@
    Start-Sleep -Milliseconds 200
    Clear-Host
    Write-Host @"










    0% [ ########################                      ] 100%
"@
    Start-Sleep -Milliseconds 400
    Clear-Host 
    Write-Host @"










    0% [ #############################                 ] 100%
"@
    Start-Sleep -Milliseconds 275
    Clear-Host 
    Write-Host @"










    0% [ ##########################################    ] 100%
"@
    Start-Sleep -Milliseconds 320
    Clear-Host 
    Write-Host @"










    0% [ ############################################# ] 100%
"@
    Start-Sleep -Milliseconds 200
    Clear-Host 
    Write-Host @"










    0% [ ############################################# ] 100%
"@  -ForegroundColor Green
    Start-Sleep -Milliseconds 600
    Clear-Host 
}

fake_loading_screen

#EXIT SCREEN
function exit_screen_for_games{
    Clear-Host
    Write-Host @"







                     ____ __     __ ______ 
                    |  _ \\ \   / /|  ____|
                    | |_) |\ \_/ / | |__   
                    |  _ <  \   /  |  __|  
                    | |_) |  | |   | |____ 
                    |____/   |_|   |______|
                                                                          
"@
Start-Sleep -Seconds 1
exit
}

#SECRET KEY
function secret_key_10001{
    Clear-Host 
    Write-Host @"






    
            8 8 8 8                     ,ooo.
            8a8 8a8                    oP   ?b
           d888a888zzzzzzzzzzzzzzzzzzzz8     8b
             `""^""'                    ?o___oP'

            Congratulations! You found the secret
"@
Start-Sleep -Seconds 2
$secretkeyfound = $true
main_menu
}

#MAIN MENU
function main_menu{

    Clear-Host

    change_screen_size_for_games

    Write-Host @"
================== Games Galore! - Main Menu ====================
"@

    #IF THE PLAYER HAS ALREADY FOUND THE SECRET KEY
    if ($secretkeyfound){

        #ENSURE WE GET A DIFFERENT COLOUR EACH TIME FOR THE TEXT
        do {
            $found_secret_key_random_color = Get-Random $found_secret_key_random_colours
        }until ($found_secret_key_random_color -ne $found_secret_key_random_text_colour)

        #SET THE TEXT COLOUR TO THIS RANDOM COLOUR
        $found_secret_key_random_text_colour = $found_secret_key_random_color

        Write-Host @"

           You have already found the secret key!

"@ -ForegroundColor $found_secret_key_random_text_colour
    }else{
        Write-Host @"

                  Can you find the secret key?

"@
    }

    Write-Host "Guessing Games------------" -BackgroundColor DarkCyan
    Write-Host @"
 1: Press '1' for the 1-10 guess game
 2: Press '2' for the 1-100 guess game
 3: Press '3' for the verb guessing game
"@
    Write-Host "Multiplication Games------" -BackgroundColor DarkCyan
    Write-Host @"
 4: Press '4' for the easy multiplication game (1-10)
 5: Press '5' for the intermediate multiplication game (10-20)
 6: Press '6' for the hard multiplication game (20-50)
"@
    Write-Host "Other Games---------------" -BackgroundColor DarkCyan
    Write-Host @"
 7: Press '7' for rock, paper, scissors
 8: Press '8' for quickfire counting
 9: Press '9' for blackjack
"@  
    Write-Host ""
    Write-Host @"
Press 'H' for help  |  Press 'Q' to quit  
"@ -BackgroundColor Blue

    $main_menu_selection = Read-Host "select an option"

    #SWITCH FOR USER INPUT ON MAIN MENU
    switch ($main_menu_selection){
        "1" {1_to_10_guessgame}
        "2" {1_to_100_guessgame}
        "3" {verb_guessing_game}
        "4" {easy_math_game}
        "5" {intermediate_math_game}
        "6" {hard_math_game}             
        "7" {rock_paper_scissors}
        "8" {quick_fire_counting}
        "9" {blackjack}   
        "h" {main_help_menu}
        #LOGIC FOR COSMETIC CHANGES ON SECRET KEY TEST AND MAIN MENU BEHAVIOUR
        "/" { if (!$secretkeyfound){secret_key_10001}else{main_menu} }
        "q" {
            Clear-Host; 
            #MAKING SURE THE INPUT IS VALID
            do {$are_you_sure_message = Read-Host "Are you sure you want to quit? Y or N?"} while (("y","n") -notcontains $are_you_sure_message); 
            switch ($are_you_sure_message){
            "y" {exit_screen_for_games}
            "n"{main_menu} 
            } 
        }
        #IF INPUT ISN'T VALID, THE MAIN MENU IS DISPLAYED AGAIN
        default {main_menu}
    }
}

#1 TO 10 GUESS GAME
function 1_to_10_guessgame{

    #CHANGING NAME OF WINDOW
    $pshost = Get-Host 
    $pswindow = $pshost.ui.RawUI
    $pswindow.WindowTitle = "1 to 10 Guess Game"

    Clear-Host
    Write-Host "You only have one guess!"
    Start-Sleep -Milliseconds 500
    
    #GENERATING A RANDOM NUMBER BETWEEN 1 AND 10 INCLUDING 10
    $1_to_10_random_number = Get-Random -Minimum 1 -Maximum 11

    #DO THIS UNTIL THE GUESS IS BETWEEN 1 AND 10 AND PASSES THE NUM CHECK
    do {
        try{
            $1_to_10_numok = $true
            [int]$1_to_10_guess = Read-Host "Make a guess between 1 and 10"
        }catch{
            $1_to_10_numok = $false
        }
    }until ($1_to_10_guess -ge 1 -and $1_to_10_guess -le 10 -and $1_to_10_numok)

    #WHEN GUESS IS BETWEEN 1 AND 10 AND PASSES NUM CHECK IN ABOVE DO/TRY STATEMENTS
    if ($1_to_10_guess -eq $1_to_10_random_number){
        Write-Host "You got it right!"
    }else{
        Write-Host "You got it wrong! the number was $1_to_10_random_number"
    }

    Start-Sleep -Seconds 1

    #ASK USER IF THEY WANT TO REPLAY UNTIL INPUT IS A Y OR N
    do {$1_to_10_play_again = Read-Host "Do you want to play again? Y or N"} while (("y","n") -notcontains $1_to_10_play_again)

    #SWITCH TO EITHER PLAY AGAIN OR GO TO MAIN MENU
    switch ($1_to_10_play_again){
        "y" {1_to_10_guessgame}
        "n" {main_menu}
        default {main_menu}
    }
}

#1 TO 100 GUESS GAME
function 1_to_100_guessgame{

    #SETTING THE GUESS TRIES TO 0
    $1_to_100_guess_tries = 0

    #CHANGING NAME OF WINDOW
    $pshost = Get-Host
    $pswindow = $pshost.UI.RawUI
    $pswindow.WindowTitle = "1 to 100 Guess Game"
    Clear-Host

    #GENERATING A RANDOM NUMBER BETWEEN 1 AND 100 INCLUDING 100
    $1_to_100_random_number = Get-Random -Minimum 1 -Maximum 101

    Write-Host "Try to guess the number between 1 and 100"

    #DOES THIS WHILE THE GUESS IS NOT EQUAL TO THE RANDOMLY GENERATED NUMBER ABOVE
    while ($1_to_100_guess -ne $1_to_100_random_number){
        [int]$1_to_100_guess = Read-Host "Enter guess"

        #DETERMINES WHAT TO OUTPUT TO THE CONSOLE AND ADDS TO THE GUESS TRIES NUMBER IF TOO HIGH OR TOO LOW
        #IF INPUT IS NOT BETWEEN 1 AND 100 OR ISN'T A NUMBER/EMPTY
        if ($1_to_100_guess -gt 100 -or $1_to_100_guess -lt 0){
            Write-Host "Guess must be between 0 and 100! - that doesn't count!"
        #IF INPUT IS GREATER THAN THE RANDOMLY GENERATED NUMBER / ADDS 1 TO THE GUESS TRIES NUMBER
        }elseif ($1_to_100_guess -gt $1_to_100_random_number){
            Write-Host "$1_to_100_guess is too high!"
            $1_to_100_guess_tries = $1_to_100_guess_tries + 1
        #IF INPUT IF LESS THAN THE RANDOMLY GENERATED NUMBER / ADD 1 TO THE GUESS TRIES NUMBER
        }elseif ($1_to_100_guess -lt $1_to_100_random_number){
            Write-Host "$1_to_100_guess is too low!"
            $1_to_100_guess_tries = $1_to_100_guess_tries + 1
        }else{
            #UNHANDLED ERROR
        }

    }Write-Host "Correct! the number was $1_to_100_random_number. It took you $1_to_100_guess_tries tries"

    #ASK USER IF THEY WANT TO REPLAY UNTIL INPUT IS A Y OR N
    do {$1_to_100_play_again = Read-Host "Do you want to play again? Y or N"} while (("y","n") -notcontains $1_to_100_play_again)

    #SWITCH TO EITHER PLAY AGAIN OR GO TO MAIN MENU
    switch ($1_to_100_play_again){
        "y" {1_to_100_guessgame}
        "n" {main_menu}
        default {main_menu}
    }
}

#VERB GUESSING GAME
function verb_guessing_game{

    #CHANGING NAME OF WINDOW
    $pshost = Get-Host
    $pswindow = $pshost.UI.RawUI
    $pswindow.WindowTitle = "Verb Guessing"  

    #GETTING VERBS FOR THE GAME
    $verb_guessing_game_get_verbs = Get-Verb | Select-Object -ExpandProperty verb

    #RESETTING SCORE
    [int]$verb_guessing_game_score = 0

    #DO THIS (PLAY THE GAME) UNTIL THE INPUT DOESNT MATCH THE VERB WITHOUT VOWALS
    do{
        #PICKING A RANDOM VERB, REMOVING THE VOWELS AND PUTTING IT INTO A VARIABLE
        $verb_guessing_game_random_verb = Get-Random $verb_guessing_game_get_verbs
        $verb_guessing_game_random_verb_vowels_removed = $verb_guessing_game_random_verb -replace '[aeiou]','_'

        #INCREASING SCORE
        $verb_guessing_game_score = $verb_guessing_game_score + 1

        Clear-Host

        #WRITING VERB WITH NO VOWELS TO CONSOLE AND PROMPTING FOR INPUT
        Write-Host $verb_guessing_game_random_verb_vowels_removed
        $verb_guessing_game_input = Read-Host "What is the word?"
    }until($verb_guessing_game_input -ne $verb_guessing_game_random_verb)

    #WRITING RESULTS TO CONSOLE ONCE THE USER AS FAILED
    Clear-Host
    Write-Host "Incorrect! The verb was $verb_guessing_game_random_verb"
    Write-Host "Your score was $verb_guessing_game_score"
    Start-Sleep -Seconds 2

    #ASK USER IF THEY WANT TO REPLAY UNTIL INPUT IS A Y OR N
    do {$verb_guessing_game_play_again = Read-Host "Do you want to play again? Y or N"} while (("y","n") -notcontains $verb_guessing_game_play_again)

    #SWITCH TO EITHER PLAY AGAIN OR GO TO MAIN MENU
    switch ($verb_guessing_game_play_again){
        "y" {verb_guessing_game}
        "n" {main_menu}
        default {main_menu}
    }
}

#EASY MATH GAME (1-10)
function easy_math_game{

    #CHANGING NAME OF WINDOW
    $pshost = Get-Host
    $pswindow = $pshost.UI.RawUI
    $pswindow.WindowTitle = "Easy Math Game"

    #GENERATING 2 RANDOM NUMBERS BETWEEN 1 AND 10
    $easy_math_game_number = Get-Random -Minimum 1 -Maximum 11
    $easy_math_game_number_to_be_solved = Get-Random -Minimum 1 -Maximum 11

    #MULTIPLYING NUMBERS TOGETHER
    $easy_math_game_multiplied_numbers = $easy_math_game_number * $easy_math_game_number_to_be_solved

    Clear-Host
    Write-Host "Here is your question, can you work out what X is?"
    Write-Host "$easy_math_game_number * X = $easy_math_game_multiplied_numbers"

    #DO THIS UNTIL THE SOLVE IS BETWEEN 1 AND 10 AND PASSES THE NUM CHECK
    do {
        try{
            $easy_math_game_numok = $true
            [int]$easy_math_game_input = Read-Host "Enter your answer here (1-10)"
        }catch {
            $easy_math_game_numok = $false
        }
    }until ($easy_math_game_input -ge 1 -and $easy_math_game_input -le 10 -and $easy_math_game_numok)

    #WHEN THE INPUT IS BETWEEN 1 AND 10 AND NUMOK CHECK IS PASSED
    if ($easy_math_game_input -eq $easy_math_game_number_to_be_solved){
        Write-Host "You got it right!"
    }else{
        Write-Host "You got it wrong, the answer was $easy_math_game_number_to_be_solved"
    }

    Start-Sleep -Seconds 1

    #ASK USER IF THEY WANT TO REPLAY UNTIL INPUT IS A Y OR N
    do {$easy_math_game_play_again = Read-Host "Do you want to play again? Y or N"} while (("y","n") -notcontains $easy_math_game_play_again)

    #SWITCH TO EITHER PLAY AGAIN OR GO TO MAIN MENU
    switch ($easy_math_game_play_again){
        "y" {easy_math_game}
        "n" {main_menu}
        default {main_menu}
    }
}

#INTERMEDIATE MATH GAME (10-20)
function intermediate_math_game{

    #CHANGING NAME OF WINDOW
    $pshost = Get-Host
    $pswindow = $pshost.UI.RawUI
    $pswindow.WindowTitle = "Intermediate Math Game"

    #GENERATING 2 RANDOM NUMBERS BETWEEN 10 AND 20
    $intermediate_math_game_number = Get-Random -Minimum 10 -Maximum 21
    $intermediate_math_game_number_to_be_solved = Get-Random -Minimum 10 -Maximum 21

    #MULTIPLYING NUMBERS TOGETHER
    $intermediate_math_game_multiplied_numbers = $intermediate_math_game_number * $intermediate_math_game_number_to_be_solved

    Clear-Host
    Write-Host "Here is your question, can you work out what X is?"
    Write-Host "$intermediate_math_game_number * X = $intermediate_math_game_multiplied_numbers"

    #DO THIS UNTIL THE SOLVE IS BETWEEN 10 AND 20 AND PASSES THE NUM CHECK
    do {
        try{
            $intermediate_math_game_numok = $true
            [int]$intermediate_math_game_input = Read-Host "Enter your answer here (10-20)"
        }catch{
            $intermediate_math_game_numok = $false
        }
    }until ($intermediate_math_game_input -ge 10 -and $intermediate_math_game_input -le 20 -and $intermediate_math_game_numok)

    #WHEN THE INPUT IS BETWEEN 10 AND 20 AND NUMOK CHECK IS PASSED
    if ($intermediate_math_game_input -eq $intermediate_math_game_number_to_be_solved){
        Write-Host "You got it right!"
    }else{
        Write-Host "You got it wrong, the answer was $intermediate_math_game_number_to_be_solved"
    }

    Start-Sleep -Seconds 1

    #ASK USER IF THEY WANT TO REPLAY UNTIL INPUT IS A Y OR N
    do {$intermediate_math_game_play_again = Read-Host "Do you want to play again? Y or N"} while (("y","n") -notcontains $intermediate_math_game_play_again)

    #SWITCH TO EITHER PLAY AGAIN OR GO TO MAIN MENU
    switch ($intermediate_math_game_play_again){
        "y" {intermediate_math_game}
        "n" {main_menu}
        default {main_menu}
    }
}

#HARD MATH GAME (20 - 50)
function hard_math_game{

    #CHANGING NAME OF WINDOW
    $pshost = Get-Host
    $pswindow = $pshost.UI.RawUI
    $pswindow.WindowTitle = "Hard Math Game"

    #GENERATING 2 RANDOM NUMBERS BETWEEN 20 AND 50
    $hard_math_game_number = Get-Random -Minimum 20 -Maximum 51
    $hard_math_game_number_to_be_solved = Get-Random -Minimum 20 -Maximum 51

    #MULTIPLYING NUMBERS TOGETHER
    $hard_math_game_multiplied_numbers = $hard_math_game_number * $hard_math_game_number_to_be_solved

    Clear-Host
    Write-Host "Here is your question, can you work out what X is?"
    Write-Host "$hard_math_game_number * X = $hard_math_game_multiplied_numbers"

    #DO THIS UNTIL THE SOLVE IS BETWEEN 20 AND 50 AND PASSES THE NUM CHECK
    do {
        try{
            $hard_math_game_numok = $true
            [int]$hard_math_game_input = Read-Host "Enter your answer here (20-50)"
        }catch{
            $hard_math_game_numok = $false
        }
    }until ($hard_math_game_input -ge 20 -and $hard_math_game_input -le 50 -and $hard_math_game_numok)

    #WHEN THE INPUT IS BETWEEN 20 AND 50 AND NUMOK CHECK IS PASSED
    if ($hard_math_game_input -eq $hard_math_game_number_to_be_solved){
        Write-Host "You got it right!"
    }else{
        Write-Host "You got it wrong, the answer was $hard_math_game_number_to_be_solved"
    }

    Start-Sleep -Seconds 1

    #ASK USER IF THEY WANT TO REPLAY UNTIL INPUT IS A Y OR N
    do {$hard_math_game_play_again = Read-Host "Do you want to play again? Y or N"} while (("y","n") -notcontains $hard_math_game_play_again)

    #SWITCH TO EITHER PLAY AGAIN OR GO TO MAIN MENU
    switch ($hard_math_game_play_again){
        "y" {hard_math_game}
        "n" {main_menu}
        default {main_menu}
    }
}

#ROCK PAPER SCISSORS
function rock_paper_scissors{

    #CHANGING NAME AND SIZE OF WINDOW
    $pshost = Get-Host
    $pswindow = $pshost.UI.RawUI
    $newsize = $pswindow.windowsize
    $pswindow.WindowTitle = "Rock Paper Scissors"  

    #RESETTING SCORES
    [int]$rock_paper_scissors_user_score = 0
    [int]$rock_paper_scissors_computer_score = 0

    #DO THIS (PLAY THE GAME) UNTIL EITHER THE USER OR COMPUTER HAS A SCORE OF 3
    do{
        #CHOSING AN OPTION FOR THE COMPUTER AND BUILDING THE LOSING TEXT
        $rock_paper_scissors_computer_input = Get-Random ("Rock","Paper","Scissors")
        $rock_paper_scissors_losing_text = "You lose, the computer chose $rock_paper_scissors_computer_input"

        Clear-Host  

        #DISPLAYING SCORES ON THE CONSOLE
        Write-Host "You score - $rock_paper_scissors_user_score" -ForegroundColor Green
        Write-Host "Computer's score - $rock_paper_scissors_computer_score" -ForegroundColor Red

        #GET USER INPUT AND MAKE SURE ITS VALID
        do {$rock_paper_scissors_user_input = Read-Host "Rock, paper or scissors?"} while (("Rock","Paper","Scissors") -notcontains $rock_paper_scissors_user_input)
        
        #SWITCH TO EVALUATE IF THE USER OR COMPUTER WON THE ROUND
        switch ($rock_paper_scissors_user_input){

            {$_ -eq "rock" -and $rock_paper_scissors_computer_input -eq "paper"}{Write-Host $rock_paper_scissors_losing_text; Start-Sleep -Seconds 1; $rock_paper_scissors_computer_score = $rock_paper_scissors_computer_score + 1}
            {$_ -eq "paper" -and $rock_paper_scissors_computer_input -eq "scissors"}{Write-Host $rock_paper_scissors_losing_text; Start-Sleep -Seconds 1; $rock_paper_scissors_computer_score = $rock_paper_scissors_computer_score + 1}
            {$_ -eq "scissors" -and $rock_paper_scissors_computer_input -eq "rock"}{Write-Host $rock_paper_scissors_losing_text; Start-Sleep -Seconds 1; $rock_paper_scissors_computer_score = $rock_paper_scissors_computer_score + 1}
            {$_ -eq $rock_paper_scissors_computer_input}{Write-Host "It's a draw"; Start-Sleep -Seconds 1}
            default {Write-Host "You won, the computer chose $rock_paper_scissors_computer_input"; Start-Sleep -Seconds 1; $rock_paper_scissors_user_score = $rock_paper_scissors_user_score + 1}
        }

    }until ($rock_paper_scissors_user_score -ge 3 -or $rock_paper_scissors_computer_score -ge 3)

    #EVALUATE IF THE USER OR COMPUTER WON OVERALL
    if ($rock_paper_scissors_user_score -gt $rock_paper_scissors_computer_score){
        Clear-Host
        Write-Host "You won, you got $rock_paper_scissors_user_score. The computer scored $rock_paper_scissors_computer_score" -ForegroundColor Green
    }else{
        Clear-Host  
        Write-Host "You lost, the computer scored $rock_paper_scissors_computer_score. You scored $rock_paper_scissors_user_score" -ForegroundColor Red
    }

    #ASK USER IF THEY WANT TO REPLAY UNTIL INPUT IS A Y OR N
    do {$rock_paper_scissors_play_again = Read-Host "Do you want to play again? Y or N"} while (("y","n") -notcontains $rock_paper_scissors_play_again)

    #SWITCH TO EITHER PLAY AGAIN OR GO TO MAIN MENU
    switch ($rock_paper_scissors_play_again){
        "y" {rock_paper_scissors}
        "n" {main_menu}
        default {main_menu}
    }   
}

#QUICK FIRE COUNTING
function quick_fire_counting{

    #CHANGING NAME OF WINDOW
    $pshost = Get-Host
    $pswindow = $pshost.UI.RawUI
    $pswindow.WindowTitle = "Quick Fire Counting"    

    #GETTING VERBS FOR THE GAME
    $quick_fire_counting_get_verbs = Get-Verb | Select-Object -ExpandProperty verb

    #RESETTING SCORE
    [int]$quick_fire_counting_score = 0

    #RESETTING DISPLAY TIMER
    [int]$quick_fire_counting_display_timer = 1000

    #DO THIS (PLAY THE GAME) UNTIL THE INPUT DOESNT MATCH THE VERB'S LENGTH
    do {
        #PICKING A RANDOM VERB AND PUTTING THE LENGTH INTO A VARIABLE
        $quick_fire_counting_random_verb = Get-Random $quick_fire_counting_get_verbs
        $quick_fire_counting_random_verb_lenth = $quick_fire_counting_random_verb.Length

        #INCREASING THE SCORE BY ONE
        $quick_fire_counting_score = $quick_fire_counting_score + 1

        #DECREASING THE DISPLAY TIMER
        if ($quick_fire_counting_display_timer -ge 80){
            $quick_fire_counting_display_timer = $quick_fire_counting_display_timer - 30
        }else{}
        
        Clear-Host

        #OUTPUT VERB TO CONSOLE THEN CLEAR AFTER DISPLAY TIME IS OVER, CLEAR HOST AND THEN ASK FOR INPUT
        Write-Host $quick_fire_counting_random_verb
        Start-Sleep -Milliseconds $quick_fire_counting_display_timer
        Clear-Host
        $quick_fire_counting_input = Read-Host "How many letters did that word have?"   
    }until ($quick_fire_counting_input -ne $quick_fire_counting_random_verb_lenth)

    #WRITING RESULTS TO CONSOLE ONCE THE USER AS FAILED
    Clear-Host  
    Write-Host "Incorrect! there were $quick_fire_counting_random_verb_lenth letters"
    Write-Host "Your score was $quick_fire_counting_score"
    Start-Sleep -Seconds 2
    
    #ASK USER IF THEY WANT TO REPLAY UNTIL INPUT IS A Y OR N
    do {$quick_fire_counting_play_again = Read-Host "Do you want to play again? Y or N"} while (("y","n") -notcontains $quick_fire_counting_play_again)

    #SWITCH TO EITHER PLAY AGAIN OR GO TO MAIN MENU
    switch ($quick_fire_counting_play_again){
        "y" {quick_fire_counting}
        "n" {main_menu}
        default {main_menu}
    }
}

#BLACKJACK
function blackjack{

    #CHANGING NAME OF WINDOW
    $pshost = Get-Host
    $pswindow = $pshost.UI.RawUI
    $pswindow.WindowTitle = "Blackjack" 

    #RESETTING GAME OVER VARIABLE
    $blackjack_game_over = $false

    #GENERATING A RANDOM TOTAL FOR THE DEALER
    $blackjack_dealer_total = Get-Random -Minimum 14 -Maximum 22

    #CREATING AN ARRAY FOR THE USERS CARD NUMBERS
    $blackjack_user_card_array = [System.Collections.ArrayList]::new("")

    #GENERATING A RANDOM NUMBER FOR THE USERS FIRST CARD
    $blackjack_user_first_card = Get-Random -Minimum 1 -Maximum 11

    #ADDING USERS FIRST CARD TO ARRAY
    $blackjack_user_card_array.Add($blackjack_user_first_card)

    #CREATING A VARIABLE TO COUNT USERS TOTAL
    $blackjack_user_total = $blackjack_user_first_card

    Clear-Host

    Write-Host "Your first card is $blackjack_user_first_card"

    #DO THIS (PLAY GAME) UNTIL THE GAMEOVER VARIABLE IS TRUE
    do {
        #GET USER INPUT
        do {$blackjack_input = Read-Host "Take another card? (Y or N)"}while (("y","n") -notcontains $blackjack_input)

        #IF USER INPUT IS VALID AND ISN'T BUST AND WANTS ANOTHER CARD
        if ($blackjack_input -eq "y" -and $blackjack_user_total -le 21){

            #GENERATE A NEW CARD FOR THE USER
            $blackjack_user_new_card = Get-Random -Minimum 1 -Maximum 11

            #ADD NEW CARD TO CARD ARRAY
            $blackjack_user_card_array.Add($blackjack_user_new_card)

            #ADD NEW CARD TO CARD TOTAL
            $blackjack_user_total = $blackjack_user_total + $blackjack_user_new_card

            Clear-Host

            Write-Host "You have $blackjack_user_card_array"

            #IF THE USER IS BUST
            if ($blackjack_user_total -gt 21){
                Write-Host "You went bust! The dealer won with " -ForegroundColor Red -NoNewline
                Write-Host $blackjack_dealer_total 
                $blackjack_game_over = $true
            }
            
        #IF THE USER DOESNT WANT ANOTHER CARD
        }else{

            Clear-Host
            
            #OUTPUTTING THE FINAL SCORE
            #Write-Host "You had $blackjack_user_total and the dealer had $blackjack_dealer_total"

            #SWITCH TO SEE WHO WON
            switch ($blackjack_user_total){
                {$_ -gt 21}{Write-Host "You went bust! The dealer won with " -ForegroundColor Red -NoNewline; Write-Host $blackjack_dealer_total; $blackjack_game_over = $true; break}
                {$_ -eq $blackjack_dealer_total}{Write-Host "It's a draw, the dealer also had $blackjack_dealer_total"; $blackjack_game_over = $true; break}
                {$_ -gt $blackjack_dealer_total}{Write-Host "You win! The dealer only had " -ForegroundColor Green -NoNewline; Write-Host $blackjack_dealer_total; $blackjack_game_over = $true; break}
                {$_ -lt $blackjack_dealer_total}{Write-Host "You lose! The dealer won with " -ForegroundColor Red -NoNewline; Write-Host $blackjack_dealer_total; $blackjack_game_over = $true; break}
                default {Write-Host "Something happeneds that wasn't accounted for!" -ForegroundColor Red; break}
            }
        }
    }until ($blackjack_game_over)

    #ASK USER IF THEY WANT TO REPLAY UNTIL INPUT IS A Y OR N
    do {$blackjack_play_again = Read-Host "Do you want to play again? Y or N"} while (("y","n") -notcontains $blackjack_play_again)

    #SWITCH TO EITHER PLAY AGAIN OR GO TO MAIN MENU
    switch ($blackjack_play_again){
        "y" {blackjack}
        "n" {main_menu}
        default {main_menu}
    }   
}

#    _    _ ______ _      _____     _____ ______ _____ _______ _____ ____  _   _ 
#   | |  | |  ____| |    |  __ \   / ____|  ____/ ____|__   __|_   _/ __ \| \ | |
#   | |__| | |__  | |    | |__) | | (___ | |__ | |       | |    | || |  | |  \| |
#   |  __  |  __| | |    |  ___/   \___ \|  __|| |       | |    | || |  | | . ` |
#   | |  | | |____| |____| |       ____) | |___| |____   | |   _| || |__| | |\  |
#   |_|  |_|______|______|_|      |_____/|______\_____|  |_|  |_____\____/|_| \_|
#

#MAIN HELP MENU
function main_help_menu{

    #CHANGING NAME AND SIZE OF WINDOW
    $pshost = get-host
    $pswindow = $pshost.ui.rawui
    $newsize = $pswindow.buffersize
    $newsize.height = 57
    $newsize.width = 85
    $pswindow.buffersize = $newsize
    $newsize = $pswindow.windowsize
    $newsize.height = 50
    $newsize.width = 85
    $pswindow.windowsize = $newsize

    
    Clear-Host

    Write-Host @"
============================ Games Galore! - Help Menu ==============================


"@
    Write-Host "Guessing Games" -ForegroundColor Cyan
    Write-Host @"

    There are two different versions of this game. The first, generates
    a random number between 1 and 10. You then have one chance to guess
    this number. If you get it right, you win. If you get it wrong, you
    lose.

    The second version generates a number between 1 and 100. You then
    have unlimited guesses in order to find the number. The game 
    provides information for if your guess was too high or low. 
    The game outputs how many attempts it too you to finish at the end. 
    Count this as a score, I guess.

"@
    Write-Host "Math Games" -ForegroundColor Cyan
    Write-Host @"

    There are three versions of this game, ranging in difficulty. 
    The games multiply two numbers together but only show you one
    of these as you need to work the other one out. 

    The game will display you with a sum, with the second number 
    missing. You have to solve the sum and solve the missing number
    represented by X.

"@
    Write-Host "Quick Fire Counting" -ForegroundColor Cyan
    Write-Host @"

    YOU HAVE TO BE QUICK, WORDS DISPLAY AND DISAPPEAR VERY QUICKLY!

    The aim of this game is to quickly look at the word that flashes on screen
    and then as fast as you are possible, enter the number of characters
    in that word.
    
    The game counts the amount of times that you enter the correct number, if you 
    enter the incorrect amount of characters in a word then you lose your score
    and have a choice to play again or not.
    
    The text displays for 1 second and then decreases 30ms for every correct answer,
    but this stops at 80ms so that the game doesn't become impossible.

    The words come from the verbs in PowerShell incase you were interested. 
"@
    Write-Host "Verb Guessing" -ForegroundColor Cyan
    Write-Host @"
    
    In this game, you have the objective of trying to guess a verb with all of
    the vowels removed.
    
    An example of this would be "nstll". For those familiar to PowerShell,
    you can easily recognise this as install.
    
    Not all of the verbs will be this easy. For verbs like use, the only
    clue you will get is "s". And finally, I know that the verb sync doesn't
    have any vowels and yet still pops up. Count it as an easter egg :) 

"@
    Write-Host @"

Press any key to return to the Main Menu...
"@ -BackgroundColor Blue
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    main_menu
}

main_menu