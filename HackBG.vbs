ON ERROR RESUME NEXT
'100% By @0znzw on scratch or SurvExE_Pc on github
'Read license.txt
'Config is in config.txt
'And read README.txt
Dim oShell,FSO
'Objects used in app.
Set oShell = CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")

'Enviroment data!
appdata=oShell.ExpandEnvironmentStrings("%APPDATA%")
scriptdir = FSO.GetAbsolutePathName(".")

'Themes.
themes=appdata+"\Microsoft\Windows\Themes\"
paper=themes+"TranscodedWallpaper"
cach=themes+"\CachedFiles"

'Backup :|
Function CopyFiles(FiletoCopy,DestinationFolder)
   Dim fso
                Dim Filepath,WarFileLocation
                Set fso = CreateObject("Scripting.FileSystemObject")
                If  Right(DestinationFolder,1) <>"\"Then
                    DestinationFolder=DestinationFolder&"\"
                End If
    fso.CopyFile FiletoCopy,DestinationFolder,True
                FiletoCopy = Split(FiletoCopy,"\")

End Function

CopyFiles paper,scriptdir

'Get background..
Set wShell=oShell
Set oExec=wShell.Exec("mshta.exe ""about:<input type=file id=FILE><script>FILE.click();new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);</script>""")
sFileSelected = oExec.StdOut.ReadLine
sFileSelected = "" & sFileSelected & ""
FileName = FSO.GetFile(sFileSelected).Name
CopyFiles sFileSelected,scriptdir

'wscript.echo FileName

'Cleaning themes directory :\
If Not FSO.FolderExists(cach) then      
    FSO.CreateFolder(cach)
End If
Set objFolder = FSO.GetFolder(cach)
Set allFiles = objFolder.Files
For Each objFile in allFiles
    objFile.Delete
Next
Set objFolder = nothing
Set allFiles = nothing
FSO.DeleteFolder cach, True

'File string setups :D
paperreplace = "" & scriptdir+FileName & ""
paperreplace = "" & paperreplace & ""
paperTheme = "" & themes+FileName & ""
paperTheme = "" & paperTheme & ""
'wscript.echo paperreplace
'wscript.echo paperTheme

'Empty the themes folder :O
Set objFolder = FSO.GetFolder(themes)
Set allFiles = objFolder.Files
For Each objFile in allFiles
    objFile.Delete
Next
Set objFolder = nothing
Set allFiles = nothing

'Finally replace(rename to remake) the wallpaper file!
CopyFiles paperreplace,themes
FSO.GetFile(paperTheme).Name="TranscodedWallpaper"

'Restart File explorer to finish the job!!!
oShell.Run "taskkill /f /im explorer.exe"
wscript.sleep 2500
oShell.Run "explorer.exe"

'Cleanup current directory based on config
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
Set Config = FSO.OpenTextFile(scriptdir+"Config.txt", ForReading)
Config.ReadLine
KeepCopys = Config.ReadLine
Config.ReadLine
KeepTranscodedWp = Config.ReadLine
Config.Close

If KeepCopys = 0 Then
    FSO.GetFile(scriptdir+FileName).Delete
End If
If KeepTranscodedWp = 0 Then
    FSO.GetFile(scriptdir+"TranscodedWallpaper").Delete
End If

'Delete all stuff
appdata = nothing
FileName = nothing
paper = nothing
sFileSelected = nothing
Set wShell = nothing
Set oExec = nothing
scriptDir = nothing
themes = nothing
cach = nothing
paperreplace = nothing
paperTheme = nothing
Set FSO = nothing
Set oShell = nothing