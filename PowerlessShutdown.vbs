' This is the first version to be fully written in VSCode! I was just using Notepad before.
' I've removed the errors so the code is easier to read. They were not a good idea.

Set WshShell = CreateObject("WScript.Shell")
result = MsgBox("Do you want to power off the PC? (This won't update the computer)",vbYesNoCancel+32,"PowerlessShutdown v4.i5") ' 1
Select Case result
       Case vbYes
            WshShell.Run "shutdown /s /t 0",0,False ' 0,False = Silent command
            WScript.Quit
       Case vbNo
            result = MsgBox("Do you want to restart into advanced boot options/recovery menu? (This won't update the computer)",vbYesNo+32,"PowerlessShutdown v4.i5") ' 2
            Select Case result
                   Case vbYes
                        WshShell.Run "shutdown /r /o /t 0",0,False
                        WScript.Quit
                   Case vbNo
                        result = MsgBox("Do you want to log off (if you're unable to log off)?",vbYesNo+32,"PowerlessShutdown v4.i5") ' 3
                        Select Case result
                               Case vbYes
                                    WshShell.Run "shutdown /l",0,False ' /l flag doesn't do anything with /t 0, completely useless
                                    WScript.Quit
                               Case vbNo
                                    result = MsgBox("Do you want to put the PC into sleep mode? (This disables Hibernation)",vbYesNo+32,"PowerlessShutdown v4.i5") ' 4
                                    Select Case result
                                           Case vbYes
                                                WshShell.Run "powercfg -h off",0,False
                                                WshShell.Run "rundll32.exe powrprof.dll,SetSuspendState 0,1,0",0,False ' Why is it so hard to just put a pc to hibernate/sleep?????
                                           Case vbNo
                                                result = MsgBox("Do you want to hibernate the PC? (This enables Hibernation)",vbYesNo+32,"PowerlessShutdown v4.i5") ' 5
                                                Select Case result
                                                       Case vbYes
                                                            WshShell.Run "powercfg -h on",0,False
                                                            WshShell.Run "rundll32.exe powrprof.dll,SetSuspendState 0,1,0",0,False
                                                       Case vbNo
                                                            WScript.Quit
                                                End Select
                                    End Select                                         
                        End Select                  
            End Select
       Case vbCancel
            MsgBox "Quitting, if the script still continues use Task Manager to quit.",64,"Quitting!"
            WScript.Quit
End Select

' Internal version: 4-i5. 3-i3 was skipped due to having bugs that needed fixing
' Made by PXbass2.