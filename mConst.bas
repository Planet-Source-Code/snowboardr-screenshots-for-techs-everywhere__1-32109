Attribute VB_Name = "mConst"

Public SetNum       As Integer 'settings number
Public Color        As String 'Background color
Public currentURL   As String
Public NewUrl       As String
Public AutoCopy     As String
Public MainFolder   As String

Public IChecked     As Integer  'For saving a checkbox value in frmGenerate



Public Function GotoFolder(Folder As String, Area As String)
Dim ServerAddress           As String
'# Option 1

ServerAddress = "http://localhost/ss/"
frmMain.wb1.Navigate ServerAddress & "screenshot/getscreenshot.asp?Folder=" & Folder & "&" & "Area=" & Area & "&" & "color=" & Color & "&settings=" & SetNum

'# Option 2
'frmMain.wb1.Navigate GetMainFolder & "screenshot/getscreenshot.asp?Folder=" & Folder & "&" & "Area=" & Area & "&" & "color=" & Color & "&settings=" & SetNum

End Function


