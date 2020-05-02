Attribute VB_Name = "data"
'API for puuting on top of all other applications
Declare Function SetWindowPos Lib "User32" _
(ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
ByVal cy As Long, ByVal wFlags As Long) As Long

'API for putting in the system tray
Public Declare Function Shell_NotifyIcon Lib _
"shell32.dll" Alias "Shell_NotifyIconA" _
(ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Public Type NOTIFYICONDATA
   cbSize As Long
   hwnd As Long
   uID As Long
   uFlags As Long
   uCallbackMessage As Long
   hIcon As Long
   szTip As String * 64
End Type

Global NotifyIcon As NOTIFYICONDATA

Public db As database
Public rs As Recordset
Public ws As Workspace
Public rmcode As String
Public rmcodepop As String

Global Const NIM_ADD = &H0
Global Const NIM_MODIFY = &H1
Global Const NIM_DELETE = &H2
Global Const NIF_MESSAGE = &H1
Global Const NIF_ICON = &H2
Global Const NIF_TIP = &H4
Global Const WM_MOUSEMOVE = &H200

Public Function initdb()
On Error GoTo error
Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path & "\monvola.mdb")
Set rs = db.OpenRecordset("monvola", dbOpenTable)
rs.Index = "alm_done"
Exit Function

error:
MsgBox "Either database 'monvola.mdb' or a table 'EMinderTable' in this database or both missing." _
& vbCrLf & "If yes the make sure this database is in the path from where application is running. " _
& vbCrLf & "Or there is some unexpected error while opening the database.", _
vbCritical + vbOKOnly, "Monvola - Database Error"
MsgBox "Monvola can not be started.", vbCritical + vbOKOnly, "Monvola - Opening Error"
End
End Function

Public Function savealarm(alcode As String)
On Error GoTo notadded
' add new record in the database
rs.AddNew
' Populate all the fields of the database
rs("type") = rmcode
rs("sub_or_name") = main.subnam.Text
rs("com_or_desc") = main.comments.Text
rs("multi_info") = main.multi.Text
rs("alm_day") = Mid(main.aldate.Text, 4, 2)
rs("alm_month") = Mid(main.aldate.Text, 1, 2)
rs("alm_year") = Mid(main.aldate.Text, 7, 4)
rs("alm_hour") = Mid(main.altime.Text, 1, 2)
rs("alm_minute") = Mid(main.altime.Text, 4, 2)
rs("alm_done") = False
rs.Update
MsgBox "Reminder entry added sucessfully added to database.", _
vbInformation + vbOKOnly, "Monvola - Reminder Added"
Exit Function

notadded:
MsgBox "Some unexpected error occurred while saving into database." _
& vbCrLf & "Reminder entry not saved to database.", _
vbCritical + vbOKOnly, "Monvola - Database Error"
End Function

Public Function systemtray()
'We set up traypict (a picture) to accept callback data
'in it's MouseMove procedure.

NotifyIcon.cbSize = Len(NotifyIcon)
NotifyIcon.hwnd = main.traypict.hwnd
NotifyIcon.uID = 1&
NotifyIcon.uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
NotifyIcon.uCallbackMessage = WM_MOUSEMOVE

'Now, we set up the icon and tool tip message
NotifyIcon.hIcon = main.traypict.Picture
NotifyIcon.szTip = "Monvola" & vbCrLf & "Left Click to show." & vbCrLf & "Right Click to Close" & Chr$(0)

'Lastly, we add the icon
Shell_NotifyIcon NIM_DELETE, NotifyIcon
Shell_NotifyIcon NIM_ADD, NotifyIcon
End Function

