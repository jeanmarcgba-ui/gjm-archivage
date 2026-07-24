' GJM TECNOLOGIE - Lanceur silencieux (aucune fenetre noire)
Dim objShell, objFSO, strDir, strPython, strScript

Set objShell = CreateObject("WScript.Shell")
Set objFSO   = CreateObject("Scripting.FileSystemObject")

strDir    = objFSO.GetParentFolderName(WScript.ScriptFullName)
strScript = strDir & "\gjm_archivage.py"

' Chercher Python dans tous les emplacements possibles
Dim arrPy(9)
arrPy(0) = "py"
arrPy(1) = "python"
arrPy(2) = "python3"
arrPy(3) = objShell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Programs\Python\Python314\python.exe"
arrPy(4) = objShell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Programs\Python\Python313\python.exe"
arrPy(5) = objShell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Programs\Python\Python312\python.exe"
arrPy(6) = objShell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Programs\Python\Python311\python.exe"
arrPy(7) = objShell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Programs\Python\Python310\python.exe"
arrPy(8) = "C:\Python314\python.exe"
arrPy(9) = "C:\Python311\python.exe"

strPython = ""
Dim i
For i = 0 To 9
    On Error Resume Next
    Dim ret
    ret = objShell.Run("cmd /c " & arrPy(i) & " --version >nul 2>&1", 0, True)
    If Err.Number = 0 And ret = 0 Then
        strPython = arrPy(i)
        Exit For
    End If
    Err.Clear
    On Error GoTo 0
Next

If strPython = "" Then
    MsgBox "Python n'est pas installe." & vbCrLf & vbCrLf & _
           "Installez Python depuis :" & vbCrLf & _
           "https://www.python.org/downloads/" & vbCrLf & vbCrLf & _
           "Cochez 'Add Python to PATH'.", _
           vbCritical, "GJM Archivage"
    WScript.Quit 1
End If

' Installer les bibliotheques en arriere-plan (completement silencieux)
Dim strLog
strLog = strDir & "\gjm_install.log"
objShell.Run "cmd /c """ & strPython & """ -m pip install openpyxl reportlab python-docx Pillow --user --quiet --no-warn-script-location > """ & strLog & """ 2>&1", 0, True

' Lancer l'application - SANS fenetre noire (parametre 0 = cache, False = non bloquant)
objShell.CurrentDirectory = strDir
objShell.Run """" & strPython & """ """ & strScript & """", 0, False

Set objShell = Nothing
Set objFSO   = Nothing
