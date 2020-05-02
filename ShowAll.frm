VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form ShowAll 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Monvola - All Reminders"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8400
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "ShowAll.frx":0000
   ScaleHeight     =   6795
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ShowAll.frx":1351D
            Key             =   "birthday"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ShowAll.frx":1396F
            Key             =   "call"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ShowAll.frx":13DC1
            Key             =   "mail"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ShowAll.frx":14213
            Key             =   "misc"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ShowAll.frx":14665
            Key             =   "meeting"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmddel 
      Caption         =   "Delete ""Done"""
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   6
      ToolTipText     =   "Click to refresh list"
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton cmdchange 
      Caption         =   "Change"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   9
      ToolTipText     =   "Click to change date/time"
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7200
      TabIndex        =   10
      ToolTipText     =   "Click to cancel date/ time change"
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmddttm 
      Caption         =   "Change Date/Time"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      ToolTipText     =   "Click to change date/time"
      Top             =   5520
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   0
      Top             =   3840
   End
   Begin VB.CommandButton cmdrefresh 
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   1
      ToolTipText     =   "Click to refresh list"
      Top             =   5520
      Width           =   1335
   End
   Begin MSComctlLib.ListView rmlist 
      Height          =   4575
      Left            =   480
      TabIndex        =   3
      ToolTipText     =   "All Reminders"
      Top             =   720
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   8070
      SortKey         =   1
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483630
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Rem. Type"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Date"
         Object.Width           =   2116
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Time"
         Object.Width           =   1236
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Sub./Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Other Info"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Status"
         Object.Width           =   1588
      EndProperty
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "Go Back"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   0
      ToolTipText     =   "Click to go back to main"
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton cmdpending 
      Caption         =   "Mark as Pending"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      ToolTipText     =   "Click to mark as pending"
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton cmddone 
      Caption         =   "Mark as Done"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   "Click to mark as done."
      Top             =   5520
      Width           =   1335
   End
   Begin MSMask.MaskEdBox aldate 
      Height          =   285
      Left            =   1200
      TabIndex        =   7
      ToolTipText     =   "Date in MM/DD/YYYY format"
      Top             =   340
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox altime 
      Height          =   285
      Left            =   4320
      TabIndex        =   8
      ToolTipText     =   "Time in 24 hour format"
      Top             =   340
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##:##"
      PromptChar      =   "_"
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "NOTE : To sort any column, please click on the column heading."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   6480
      Width           =   7935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Total 'Done' entries = 0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   16
      Top             =   6120
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total 'Pending' entries = 0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   15
      Top             =   6120
      Width           =   3135
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "HH:MM"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   14
      Top             =   360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "MM/DD/YYYY"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   13
      Top             =   360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   12
      Top             =   360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   360
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "ShowAll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmddel_Click()

If rs.RecordCount = 0 Then ' Read all record if record count not zero
    MsgBox "No records are there in the database to delete.", vbOKOnly + vbInformation, "Monvola - No records"
    Exit Sub
End If

If (MsgBox("This will delete all the 'Done' entries." & vbCrLf & _
    "Do you want to continue ?", vbYesNo + vbQuestion, "Monvola - Warning")) = vbYes Then
    'Delete record whcih are having status as "Done"
    rs.MoveFirst
    For i = 1 To rs.RecordCount
        If rs("alm_done") = True Then rs.Delete
        rs.MoveNext
    Next i
    rmlist.ListItems.Clear
    fillwindow ' Refill the list
End If
End Sub


Private Sub cmdback_Click()
main.Show
Unload Me
End Sub


Private Sub cmdcancel_Click()
hidebuttons ' hide the date time buttons
End Sub

Private Sub cmdchange_Click()

aldate.BackColor = &H80000005
altime.BackColor = &H80000005

'Date validations
If aldate.Text = "" Then
    MsgBox "You have not entered the date.", vbOKOnly + vbCritical, "Monvola - Error"
    aldate.BackColor = &HFFFF&
    aldate.SetFocus
    Exit Sub
Else
    dateerror = ""
    checkdate (aldate.Text)
    If dateerror <> "" Then
    MsgBox dateerror, vbOKOnly + vbCritical, "Monvola - Error"
    aldate.BackColor = &HFFFF&
    aldate.SetFocus
    Exit Sub
    End If
End If

'Time validations
If altime.Text = "" Then
    MsgBox "You have not entered the time.", vbOKOnly + vbCritical, "Monvola - Error"
    altime.BackColor = &HFFFF&
    altime.SetFocus
    Exit Sub
Else
    timeerror = ""
    checktime (altime.Text)
    If timeerror <> "" Then
        MsgBox timeerror, vbOKOnly + vbCritical, "Monvola - Error"
        altime.BackColor = &HFFFF&
        altime.SetFocus
        Exit Sub
    End If
End If

If findrecord = False Then ' Call findrecord function
    MsgBox "Record is not found in the database.", vbOKOnly + vbInformation, "Monvola - Error"
    Exit Sub
Else ' if record is found then update the record with new date and time
    rs.Edit
    rs("alm_day") = Mid(aldate.Text, 4, 2)
    rs("alm_month") = Mid(aldate.Text, 1, 2)
    rs("alm_year") = Mid(aldate.Text, 7, 4)
    rs("alm_hour") = Mid(altime.Text, 1, 2)
    rs("alm_minute") = Mid(altime.Text, 4, 2)
    rs.Update
    MsgBox "Both time and date successfully changed.", vbOKOnly + vbInformation, "Monvola"
    rmlist.ListItems.Clear
    fillwindow ' refresh windowlist
    hidebuttons ' hide the date time buttons
End If
End Sub

Private Sub cmddone_Click()

If rmlist.ListItems.Count = 0 Then
    MsgBox "No entry to mark as 'Done'", vbOKOnly + vbCritical, "Monvola - Error"
    Exit Sub
End If

If rmlist.SelectedItem.SubItems(5) = "Done" Then ' Check if alraedy "Done"
    MsgBox "Status of this entry is already marked as 'Done'", vbOKOnly + vbInformation, "Monvola - Error"
    Exit Sub
End If
                                                  'Confirm change
If (MsgBox("Are you sure to mark this entry as 'Done' ?", _
    vbYesNo + vbQuestion, "Monvola - Confirm")) = vbNo Then Exit Sub

If findrecord = False Then ' find the record by calling findrecord function
    MsgBox "Record is not found in the database.", vbOKOnly + vbInformation, "Monvola - Error"
    Exit Sub
Else        ' if found then update the record
    rs.Edit
    rs("alm_done") = True
    rs.Update
    MsgBox "Successfully marked as 'Done'.", vbOKOnly + vbInformation, "Monvola"
    rmlist.ListItems.Clear
    fillwindow
End If
End Sub



Private Sub cmdpending_Click()

If rmlist.ListItems.Count = 0 Then
    MsgBox "No entry to mark as 'Pending'", vbOKOnly + vbCritical, "Monvola - Error"
    Exit Sub
End If

If rmlist.SelectedItem.SubItems(5) = "Pending" Then ' Check if already pending
    MsgBox "Status of this entry is already marked as 'Pending'", vbOKOnly + vbCritical, "Monvola - Error"
    Exit Sub
End If

                                                     'Confirm Change
If (MsgBox("Are you sure to mark this entry as 'Pending' ?", _
    vbYesNo + vbQuestion, "Monvola - Confirm")) = vbNo Then Exit Sub

If findrecord = False Then ' find the record by calling findrecord function
    MsgBox "Record is not found in the database.", vbOKOnly + vbInformation, "Monvola - Error"
    Exit Sub
Else                       ' if found then update record
    rs.Edit
    rs("alm_done") = False
    rs.Update
    MsgBox "Successfully marked as 'Pending'.", vbOKOnly + vbInformation, "Monvola"
    rmlist.ListItems.Clear
    fillwindow
End If

End Sub

Private Sub cmdrefresh_Click()
rmlist.ListItems.Clear ' clear windowlist
fillwindow
End Sub

Private Sub cmddttm_Click()

If rmlist.ListItems.Count = 0 Then
    MsgBox "No entry to change date and time.", vbOKOnly + vbCritical, "Monvola - Error"
    Exit Sub
End If
                                                ' Confirm change
If (MsgBox("Do you really want to change date and time of this entry ?", _
    vbYesNo + vbQuestion, "Monvola - Confirm")) = vbNo Then Exit Sub

showbuttons
End Sub

Private Sub Form_Load()
rmlist.ListItems.Clear ' Clear the windowlist
rmlist.Picture = LoadPicture() ' Initialize the image list
fillwindow   ' Call function to fill the window
End Sub
Private Function fillwindow()
Dim shtype As String
Dim shdate As String
Dim shtime As String
Dim shsubnam As String
Dim shmulti As String
Dim shstatus As String
Dim temp As String
Dim temp1 As String
Dim pndcount As Integer
Dim donecount As Integer
Dim pictindex As Integer

pndcount = 0
donecount = 0

Label1.Caption = "Total 'Pending' entries = 0"
Label2.Caption = "Total 'Done' entries = 0"

If rs.RecordCount = 0 Then Exit Function ' check if no record
rs.MoveFirst

For i = 1 To rs.RecordCount

shtype = rs("type")
Select Case shtype
Case "BR":  shtype = "Birthday"
            pictindex = 1 ' Picture index of the Imagelist
Case "CL":  shtype = "Call"
            pictindex = 2
Case "ML":  shtype = "Mail"
            pictindex = 3
Case "MS":  shtype = "Misc."
            pictindex = 4
Case "MT":  shtype = "Meeting"
            pictindex = 5
End Select

'Format date and then fill
temp = rs("alm_month")
temp1 = rs("alm_day")
If Len(temp) = 1 Then temp = "0" & temp
If Len(temp1) = 1 Then temp1 = "0" & temp1
shdate = temp & "/" & temp1 & "/" & rs("alm_year")

'Format time and then fill
temp = rs("alm_hour")
temp1 = rs("alm_minute")
If Len(temp) = 1 Then temp = "0" & temp
If Len(temp1) = 1 Then temp1 = "0" & temp1
shtime = temp & ":" & temp1

shsubnam = rs("sub_or_name") ' Name or subject
shmulti = rs("multi_info")   ' Other information

If rs("alm_done") = True Then ' Done or Pending
    shstatus = "Done"
    donecount = donecount + 1
Else
    shstatus = "Pending"
    pndcount = pndcount + 1
End If
' Fill all the intems of the ListView object
Set itmX = rmlist.ListItems.Add(, , shtype, , pictindex) ' Main item of ListView Object

itmX.SubItems(1) = shdate     ' first subitem of ListView Object
itmX.SubItems(2) = shtime     ' second subitem of ListView Object
itmX.SubItems(3) = shsubnam   ' third subitem of ListView Object
itmX.SubItems(4) = shmulti    ' fourth subitem of ListView Object
itmX.SubItems(5) = shstatus   ' fifth subitem of ListView Object

rs.MoveNext   ' read next record
Next i

Label1.Caption = "Total 'Pending' entries = " & pndcount
Label2.Caption = "Total 'Done' entries = " & donecount

End Function

Private Sub rmlist_ColumnClick(ByVal ColHeader As ColumnHeader)
  ' Sort according to the column heading pressed
    rmlist.SortKey = ColHeader.Index - 1
End Sub

Private Function findrecord() As Boolean
Dim fndtype As String
Dim fndcomment As String
Dim fndsubnam As String
Dim fndmulti As String
Dim fndyear As Integer
Dim fndmonth As Integer
Dim fndday As Integer
Dim fndhour As Integer
Dim fndminute As Integer

findrecord = False
If rs.RecordCount = 0 Then Exit Function ' Check no record condition
rs.MoveFirst

Select Case rmlist.SelectedItem.Text
    Case "Birthday": fndtype = "BR"
    Case "Call": fndtype = "CL"
    Case "Mail": fndtype = "ML"
    Case "Misc.": fndtype = "MS"
    Case "Meeting": fndtype = "MT"
End Select
' read all the subitems of the selected record
fndyear = Mid(rmlist.SelectedItem.SubItems(1), 7, 4)
fndmonth = Mid(rmlist.SelectedItem.SubItems(1), 1, 2)
fndday = Mid(rmlist.SelectedItem.SubItems(1), 4, 2)
fndhour = Mid(rmlist.SelectedItem.SubItems(2), 1, 2)
fndminute = Mid(rmlist.SelectedItem.SubItems(2), 4, 2)
fndsubnam = rmlist.SelectedItem.SubItems(3)
fndmulti = rmlist.SelectedItem.SubItems(4)

For i = 1 To rs.RecordCount
If fndtype = rs("type") And fndsubnam = rs("sub_or_name") And fndyear = rs("alm_year") And _
   fndmonth = rs("alm_month") And fndday = rs("alm_day") And fndhour = rs("alm_hour") And _
   fndminute = rs("alm_minute") And fndmulti = rs("multi_info") Then
   findrecord = True ' if record is found then findrecors is true
   Exit Function
Else
rs.MoveNext ' read next record from database
End If
Next i
End Function

Private Function showbuttons()
'Show all the date and time buttons for entering new date and time
cmddone.Enabled = False
cmdpending.Enabled = False
cmdrefresh.Enabled = False
cmddttm.Enabled = False
cmdcancel.Visible = True
cmdchange.Visible = True
aldate.Visible = True
altime.Visible = True
Label3.Visible = True
Label4.Visible = True
Label8.Visible = True
Label9.Visible = True
rmlist.Enabled = False
aldate.SetFocus
cmdchange.Default = True
End Function
Private Function hidebuttons()
'Hide all the date and time buttons for entering new date and time
cmddone.Enabled = True
cmdpending.Enabled = True
cmdrefresh.Enabled = True
cmddttm.Enabled = True
cmdcancel.Visible = False
cmdchange.Visible = False
aldate.Visible = False
altime.Visible = False
Label3.Visible = False
Label4.Visible = False
Label8.Visible = False
Label9.Visible = False
rmlist.Enabled = True
cmdback.Default = True
End Function
