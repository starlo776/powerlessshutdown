Set WshShell = CreateObject("WScript.Shell")
result = MsgBox("Do you want to restart the PC? (This won't update the computer!)",vbYesNo+32,"PowerlessShutdown")
Select Case result
       Case vbYes
            WshShell.Run "shutdown -r -t 0"
       Case vbNo
            
End Select
result = MsgBox("Do you want to restart into advanced boot options/recovery menu? (This won't update the computer!)",vbYesNo+32,"PowerlessShutdown")
Select Case result
       Case vbYes
            WshShell.Run "shutdown -r -o -t 0"
       Case vbNo
            WScript.Quit
End Select
            