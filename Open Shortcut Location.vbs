Set WshShell = CreateObject("WScript.Shell")

If WScript.Arguments.Count > 0 Then
    filePath = WScript.Arguments(0)
    
    ' If the item is a shortcut (.lnk), resolve its target
    If LCase(Right(filePath, 4)) = ".lnk" Then
        Set sc = WshShell.CreateShortcut(filePath)
        target = sc.TargetPath
        If target <> "" Then
            WshShell.Run "explorer.exe /select,""" & target & """"
            WScript.Quit
        End If
    End If
    
    ' If it's a regular file or target couldn't be resolved, highlight the file itself
    WshShell.Run "explorer.exe /select,""" & filePath & """"
End If
