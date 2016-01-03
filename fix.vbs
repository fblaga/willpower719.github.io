'=============================================
'File Name: DELETE_TEMP_FILES.VBS
'Comment: This script will delete all temporary files and folders
'=============================================

On Error Resume Next

'Declare variables
Dim fso 
Dim oFolder1
Dim oFolder2
Dim oFolder3
Dim oSubFolder1
Dim oSubFolder2
Dim oSubFolder3
Dim colSubfolders1
Dim colSubfolders2
Dim colSubfolders3
Dim oFile
Dim userProfile
Dim Windir
Dim folder

'Set up environment
Set WSHShell = CreateObject("WScript.Shell")
Set fso = createobject("Scripting.FileSystemObject")
userProfile = WSHShell.ExpandEnvironmentStrings("%userprofile%")
Windir = WSHShell.ExpandEnvironmentStrings("%windir%") 
folder = WSHShell.ExpandEnvironmentStrings("RARsfx*") 


'start deleting files
Set oFolder1 = fso.GetFolder(userProfile & "\AppData\Local\Temp\")
 For Each oFile In oFolder1.files
 On error resume next    
oFile.Delete, True 
 Next
'Delete folders and subfolders
Set colSubfolders1 = oFolder1.Subfolders
On Error Resume Next
For Each oSubfolder in colSubfolders1
    fso.DeleteFolder(oSubFolder), True
Next


Set colSubfolders3 = oFolder1.Subfolders
For Each oSubfolder in colSubfolders3
    fso.DeleteFolder(oSubFolder)

    fso.DeleteFolder(oFolder1), True
Next
'Clear memory
Set fso = Nothing
Set oFolder1 = Nothing
Set oFolder2 = Nothing
Set oFolder3 = Nothing
Set oSubFolder1 = Nothing
Set oSubFolder2 = Nothing
Set oSubFolder3 = Nothing
Set colSubfolders1 = Nothing
Set colSubfolders2 = Nothing
Set colSubfolders3 = Nothing
Set oFile = Nothing
Set userProfile = Nothing
Set Windir = Nothing

WScript.Quit
