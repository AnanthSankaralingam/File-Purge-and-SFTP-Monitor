Option Explicit
Dim str
Const intDaysOld = 100 
Function removeImage(ByRef destination)
    Dim obj : Set obj = createobject("Scripting.FileSystemObject") 
    Dim objFolder : Set objFolder = obj.GetFolder(destination) 'gets folder
    Dim objSubFile, objSubFolder

    For Each objSubFile In objFolder.Files 'deletes old files in folder, excluding subfolders
        If DateValue(objSubFile.DateLastModified) < DateValue(Now() - intDaysOld) Then 'checks if file is old enough for deletion
            str = str & ", " &objSubFile.GetFileName 
            objSubFile.Delete 
        End If
    Next

    For Each objSubFolder In objFolder.SubFolders 'recursisvely checks and deletes content in subfolder
        removeImage objSubFolder.Path
    Next
    Set obj=Nothing 'reset
End Function

removeImage("C:\Users\shank\vbs file functions")
MsgBox(str)



