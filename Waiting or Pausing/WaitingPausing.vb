' Pause 5 seconds:
Application.Wait (Now + TimeValue("0:00:06"))

' Pause until a specific time ie. this will pause a macro until 9:00am
Application.Wait "09:00:00"

'Wait doesn't accept delays less than 1 second
