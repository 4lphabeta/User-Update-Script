'Point this to the Release folder and it will copy all files (Don’t forget the backslash at the end)
Const SourceDir = "F:\Freshware_SQL\Vortex\Release\"
 
'Point this to User folders Parent Directory (Don’t forget the backslash at the end)
Const TargetDir  = "F:\Freshware_SQL\Vortex\Users\"
 
Result = MsgBox ("Are you sure you want to update?", vbYesNo, "Update Confirmation")
 
Select Case Result
Case vbYes
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFSO.GetFolder(TargetDir)
    Set colSubfolders = objFolder.Subfolders
 
    Set dosyalarklasor = objFSO.GetFolder(SourceDir)
    Set dosyalar = dosyalarklasor.Files
 
 
    For Each objSubfolder in colSubfolders
        if not instr(objSubfolder.Name,".git") > 0 then
            For Each dosya in dosyalar
            objFSO.CopyFile dosya, TargetDir & objSubfolder.Name & "\"
            Next
        end if 
    Next
    MsgBox "Done!"
Case vbNo
    'Don't update
End Select