Set wShell = CreateObject("WScript.Shell")

' WindowStyle (integer):
' 0: Hides the window and activates another window.
' 1: Activates and displays a window. If the window is minimized or maximized, the system restores it to its original size and position. An application should specify this flag when displaying the window for the first time.
' 2: Activates the window and displays it as a minimized window.
' 3: Activates the window and displays it as a maximized window.
' 4: Displays a window in its most recent size and position. The active window remains active.
' 5: Activates the window and displays it in its current size and position.
' 6: Minimizes the specified window and activates the next top-level window in the Z order.
' 7: Displays the window as a minimized window. The active window remains active.
' 8: Displays the window in its current state. The active window remains active.
' 9: Activates and displays the window. If the window is minimized or maximized, the system restores it to its original size and position. An application should specify this flag when restoring a minimized window.
' 10: Sets the show-state based on the state of the program that started the application.

' WaitOnReturn (boolean):
' False (default): the function returns 0 immediately after starting the program (not to be interpreted as an error code).
' True: script execution halts until the program finishes and the function returns any error code returned by the program.

return = wShell.Run("""" & WScript.Arguments(0) & """", 0, True)

' Exit with the errorlevel the App had
Set wShell = Nothing
Wscript.Quit return