
' MONSAM — Lanzador sin ventana de consola
' Doble clic sobre este archivo para abrir la app sin terminal visible

Dim objShell, objFSO, strDir, strCmd

Set objShell = CreateObject("WScript.Shell")
Set objFSO   = CreateObject("Scripting.FileSystemObject")

' Directorio donde vive este script
strDir = objFSO.GetParentFolderName(WScript.ScriptFullName)

' Cambiar al directorio de la app
objShell.CurrentDirectory = strDir

' Si no existe node_modules, instalar dependencias primero (en background silencioso)
If Not objFSO.FolderExists(strDir & "\node_modules") Then
    objShell.Run "cmd /c npm install --silent", 0, True
End If

' Lanzar Electron — el 0 es clave: oculta completamente la ventana de consola
objShell.Run "cmd /c npx electron .", 0, False

Set objShell = Nothing
Set objFSO   = Nothing
