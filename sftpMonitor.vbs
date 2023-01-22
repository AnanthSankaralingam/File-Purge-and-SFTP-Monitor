Option Explicit
Dim ObjExec,objShell,strFromProc,intButton,Port
Port = "22"
Set objShell = CreateObject("WScript.Shell")
Set ObjExec = objShell.Exec("%comspec% /c netstat -a | find "& DblQuote("ESTABLISHED") & "| find " & DblQuote(Port) &"")
strFromProc = ObjExec.StdOut.ReadAll
MsgBox(strFromProc)
'If Instr(strFromProc,"ESTABLISHED") > 0 Then
 '   intButton = objShell.Popup(strFromProc,3,"Connection Established @ Port "& Port &"",vbInformation)
'Else    
 '   intButton = objShell.Popup("Connection Not Established @ Port "& Port,3,"Connection Not Established @ Port "& Port &"",vbExclamation)
'End If  
'****************************************************************
Function DblQuote(Str)
    DblQuote = Chr(34) & Str & Chr(34)
End Function