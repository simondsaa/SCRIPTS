strComputer = "." ' Local Computer
strUser = "ignore"

Set objUser = GetObject("WinNT://" & strComputer & "/" & strUser)
objUser.AccountDisabled = True
objUser.SetInfo