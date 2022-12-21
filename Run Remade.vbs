Set WshShell = WScript.CreateObject("WScript.Shell")
Program = InputBox ("Please type the name of the program you would like to run","Run Remade","Type Program Here")
WshShell.Run "" & Program