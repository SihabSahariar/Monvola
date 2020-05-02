VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form main 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Monvola 1.0"
   ClientHeight    =   7935
   ClientLeft      =   0
   ClientTop       =   -60
   ClientWidth     =   9600
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "main.frx":0ECA
   ScaleHeight     =   7935
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin prj_monvola.N_Image N_Image5 
      Height          =   285
      Left            =   8880
      TabIndex        =   30
      Top             =   0
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   503
      Picture         =   "main.frx":1D5B0E
      PictureHover    =   "main.frx":1D5EFE
      PictureDown     =   "main.frx":1D62D2
   End
   Begin VB.PictureBox abt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   1200
      Picture         =   "main.frx":1D66C2
      ScaleHeight     =   3105
      ScaleWidth      =   6435
      TabIndex        =   25
      Top             =   3600
      Visible         =   0   'False
      Width           =   6465
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "BACK"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   5280
         MouseIcon       =   "main.frx":281884
         MousePointer    =   99  'Custom
         TabIndex        =   29
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Developed as personal project. So if there is any damage I'm not responsible. "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   480
         TabIndex        =   28
         Top             =   1680
         Width           =   5295
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Monvola Version: 1.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   27
         Top             =   960
         Width           =   2895
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Programmer: Sihab Sahariar Sizan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1200
         TabIndex        =   26
         Top             =   600
         Width           =   4335
      End
   End
   Begin prj_monvola.N_Image N_Image1 
      Height          =   930
      Left            =   4440
      TabIndex        =   21
      Top             =   6840
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   1640
      Picture         =   "main.frx":2819D6
      PictureHover    =   "main.frx":2862D0
      PictureDown     =   "main.frx":28ABCA
   End
   Begin MSMask.MaskEdBox altime 
      Height          =   285
      Left            =   5760
      TabIndex        =   4
      ToolTipText     =   "Time in 24 hour format"
      Top             =   4920
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   503
      _Version        =   393216
      BorderStyle     =   0
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
   Begin VB.PictureBox traypict 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1605
      Left            =   11040
      Picture         =   "main.frx":28F4C4
      ScaleHeight     =   1605
      ScaleMode       =   0  'User
      ScaleWidth      =   1245
      TabIndex        =   20
      Top             =   240
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   5760
      Top             =   0
   End
   Begin VB.PictureBox showpict 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   530
      Left            =   9840
      ScaleHeight     =   525
      ScaleMode       =   0  'User
      ScaleWidth      =   525
      TabIndex        =   17
      Top             =   1440
      Visible         =   0   'False
      Width           =   530
   End
   Begin VB.PictureBox miscpict 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   530
      Left            =   10680
      Picture         =   "main.frx":29038E
      ScaleHeight     =   525
      ScaleMode       =   0  'User
      ScaleWidth      =   525
      TabIndex        =   16
      Top             =   1080
      Visible         =   0   'False
      Width           =   530
   End
   Begin VB.PictureBox meetingpict 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   530
      Left            =   10320
      Picture         =   "main.frx":2907D0
      ScaleHeight     =   525
      ScaleMode       =   0  'User
      ScaleWidth      =   525
      TabIndex        =   15
      Top             =   360
      Visible         =   0   'False
      Width           =   530
   End
   Begin VB.PictureBox mailpict 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   530
      Left            =   10680
      Picture         =   "main.frx":290C12
      ScaleHeight     =   525
      ScaleMode       =   0  'User
      ScaleWidth      =   525
      TabIndex        =   14
      Top             =   1920
      Visible         =   0   'False
      Width           =   530
   End
   Begin VB.PictureBox callpict 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   530
      Left            =   10920
      Picture         =   "main.frx":291054
      ScaleHeight     =   525
      ScaleMode       =   0  'User
      ScaleWidth      =   525
      TabIndex        =   13
      Top             =   2400
      Visible         =   0   'False
      Width           =   530
   End
   Begin VB.PictureBox birthdaypict 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   530
      Left            =   10680
      Picture         =   "main.frx":291496
      ScaleHeight     =   525
      ScaleMode       =   0  'User
      ScaleWidth      =   525
      TabIndex        =   12
      Top             =   1080
      Visible         =   0   'False
      Width           =   530
   End
   Begin VB.TextBox comments 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   2640
      MaxLength       =   255
      MultiLine       =   -1  'True
      TabIndex        =   6
      ToolTipText     =   "Comments"
      Top             =   6000
      Width           =   4935
   End
   Begin VB.TextBox multi 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      MaxLength       =   100
      TabIndex        =   5
      ToolTipText     =   "Other information"
      Top             =   5400
      Width           =   4935
   End
   Begin VB.TextBox subnam 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      MaxLength       =   100
      TabIndex        =   2
      ToolTipText     =   "Name or Subject"
      Top             =   4440
      Width           =   4935
   End
   Begin VB.ComboBox comtypes 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2760
      TabIndex        =   0
      ToolTipText     =   "Reminder type"
      Top             =   3720
      Width           =   4815
   End
   Begin MSMask.MaskEdBox aldate 
      Height          =   285
      Left            =   2640
      TabIndex        =   3
      ToolTipText     =   "Date in MM/DD/YYYY format"
      Top             =   4920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      BorderStyle     =   0
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
   Begin prj_monvola.N_Image N_Image2 
      Height          =   930
      Left            =   2880
      TabIndex        =   22
      Top             =   6840
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   1640
      Picture         =   "main.frx":2918D8
      PictureHover    =   "main.frx":2961D2
      PictureDown     =   "main.frx":29AACC
   End
   Begin prj_monvola.N_Image N_Image3 
      Height          =   930
      Left            =   1320
      TabIndex        =   23
      Top             =   6840
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   1640
      Picture         =   "main.frx":29F3C6
      PictureHover    =   "main.frx":2A3CC0
      PictureDown     =   "main.frx":2A85BA
   End
   Begin prj_monvola.N_Image N_Image4 
      Height          =   930
      Left            =   6000
      TabIndex        =   24
      Top             =   6840
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   1640
      Picture         =   "main.frx":2ACEB4
      PictureHover    =   "main.frx":2B17AE
      PictureDown     =   "main.frx":2B60A8
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
      Left            =   6600
      TabIndex        =   19
      Top             =   4920
      Width           =   495
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
      Left            =   3960
      TabIndex        =   18
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Comments"
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
      Left            =   1320
      TabIndex        =   11
      Top             =   5880
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Relation"
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
      Left            =   1320
      TabIndex        =   10
      Top             =   5400
      Width           =   855
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
      Left            =   5160
      TabIndex        =   9
      Top             =   4920
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
      Left            =   1320
      TabIndex        =   8
      Top             =   4920
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
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
      Left            =   1320
      TabIndex        =   7
      Top             =   4440
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Reminder Type"
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
      Left            =   1320
      TabIndex        =   1
      Top             =   3840
      Width           =   1935
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
      
Private Declare Function ReleaseCapture Lib "User32" () As Long
Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1

Private Declare Function ShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" ( _
    ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long '1

Private Sub cmdclear_Click()


End Sub





Private Sub cmdexit_Click()

End Sub

Private Sub cmdsave_Click()


End Sub

Private Sub cmdshowall_Click()

End Sub

Private Sub comtypes_click()

'Set the  pictures and headings according to the reminder type
Label5.Visible = True
multi.Visible = True

Select Case comtypes.Text
Case "Birthday Reminder":
     showpict.Picture = birthdaypict.Picture
     Label5.Caption = "Relation"
     Label2.Caption = "Name"
     rmcode = "BR"
Case "Call Reminder":
     showpict.Picture = callpict.Picture
     Label5.Caption = "Number"
     Label2.Caption = "Name"
     rmcode = "CL"
Case "Mail Reminder":
     showpict.Picture = mailpict.Picture
     Label5.Caption = "E-mail"
     Label2.Caption = "Name"
     rmcode = "ML"
Case "Misc. Reminder":
     showpict.Picture = miscpict.Picture
     Label5.Visible = False
     multi.Visible = False
     Label2.Caption = "Subject"
     rmcode = "MS"
Case "Meeting Reminder":
     showpict.Picture = meetingpict.Picture
     Label5.Caption = "Location"
     Label2.Caption = "Subject"
     rmcode = "MT"

End Select
End Sub

Private Sub Form_Load()
'Give welcome message


' If already running then give message and end
If App.PrevInstance = True Then
MsgBox "Monvola is already running.", vbOKOnly + vbCritical, "Monvola - Already Running."
End
End If

Call initdb ' Initialize database
Timer1.Enabled = True
' Add items to the combobox and show the "Birthday Reminder" as default
comtypes.AddItem "Birthday Reminder"
comtypes.AddItem "Call Reminder"
comtypes.AddItem "Mail Reminder"
comtypes.AddItem "Misc. Reminder"
comtypes.AddItem "Meeting Reminder"
comtypes.Text = "Birthday Reminder"
rmcode = "BR"
' show the picture for "Birthday Reminder"
showpict.Picture = birthdaypict.Picture
App.TaskVisible = False

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
ReleaseCapture
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub
Private Sub Label12_Click()
abt.Visible = False
End Sub

Private Sub N_Image1_Click()
ShowAll.Show
Me.Hide
End Sub

Private Sub N_Image2_Click()
'clear all the fields and populate default values
subnam.BackColor = &H80000005
subnam.Text = ""
multi.BackColor = &H80000005
multi.Text = ""
aldate.BackColor = &H80000005
aldate.Mask = "##/##/####"
temp = "__/__/____"
aldate.Text = temp
altime.BackColor = &H80000005
temp = "__:__"
altime.Text = temp
comments.Text = ""
subnam.SetFocus
End Sub
Sub sizan()
subnam.BackColor = &H80000005
subnam.Text = ""
multi.BackColor = &H80000005
multi.Text = ""
aldate.BackColor = &H80000005
aldate.Mask = "##/##/####"
temp = "__/__/____"
aldate.Text = temp
altime.BackColor = &H80000005
temp = "__:__"
altime.Text = temp
comments.Text = ""
subnam.SetFocus
End Sub
Private Sub N_Image3_Click()
Dim temp1 As String
subnam.BackColor = &H80000005
multi.BackColor = &H80000005
aldate.BackColor = &H80000005
altime.BackColor = &H80000005

    
'Name or subject validation
If subnam.Text = "" Then
    If rmcode = "BR" Or rmcode = "ML" Or rmcode = "CL" Then
    MsgBox "You have not entered the name.", vbOKOnly + vbCritical, "Monvola - Error"
    Else
    MsgBox "You have not entered the subject.", vbOKOnly + vbCritical, "Monvola - Error"
    End If
    subnam.BackColor = &HFFFF&
    subnam.SetFocus
    Exit Sub
End If

'Date validations for all types of reminders
If aldate.Text = "" Then ' if date not entered
    MsgBox "You have not entered the date.", vbOKOnly + vbCritical, "Monvola - Error"
    aldate.BackColor = &HFFFF&
    aldate.SetFocus
    Exit Sub
Else
    dateerror = ""
    checkdate (aldate.Text)
    If dateerror <> "" Then 'if there is error in date the dateerror will not be blank
    MsgBox dateerror, vbOKOnly + vbCritical, "Monvola - Error"
    aldate.BackColor = &HFFFF&
    aldate.SetFocus
    Exit Sub
    End If
End If

'Time validations for all types of reminders
If altime.Text = "" Then ' if  time not entered
    MsgBox "You have not entered the time.", vbOKOnly + vbCritical, "Monvola - Error"
    altime.BackColor = &HFFFF&
    altime.SetFocus
    Exit Sub
Else
    timeerror = ""
    checktime (altime.Text)
    If timeerror <> "" Then 'if there is error in time the timeerror will not be blank
    MsgBox timeerror, vbOKOnly + vbCritical, "Monvola - Error"
    altime.BackColor = &HFFFF&
    altime.SetFocus
    Exit Sub
    End If
End If
    
'Now validations according to the reminder types

Select Case rmcode

Case "BR":
        If multi.Text = "" Then
            MsgBox "You have not entered your relation with - " & subnam, vbOKOnly + vbCritical, "Monvola - Error"
            multi.BackColor = &HFFFF&
            multi.SetFocus
            Exit Sub
        End If
Case "CL":
         If multi.Text = "" Then   'check if number field is blank
            MsgBox "You have not entered number to call to  - " & subnam, vbOKOnly + vbCritical, "Monvola - Error"
            multi.BackColor = &HFFFF&
            multi.SetFocus
            Exit Sub
        Else
            ' check if it is non numeric, this can be done by the IsNumeric Function,
            ' but that will allow "." to be present
            For i = 1 To Len(multi.Text)
            temp1 = Mid(multi.Text, i, 1)
            If temp1 <> "0" And temp1 <> "1" And temp1 <> "2" And temp1 <> "3" _
               And temp1 <> "4" And temp1 <> "5" And temp1 <> "6" And temp1 <> "7" _
               And temp1 <> "8" And temp1 <> "9" Then
                MsgBox "Only integer values allowed in the number.", vbOKOnly + vbCritical, "Monvola - Error"
                multi.BackColor = &HFFFF&
                multi.SetFocus
                Exit Sub
            End If
            Next i
        End If

Case "ML":
        'Mail Validations only for "Mail Reminder"
        If multi.Text = "" Then ' if blank
            MsgBox "You have not entered E-mail address.", vbOKOnly + vbCritical, "Monvola - Error"
            multi.BackColor = &HFFFF&
            multi.SetFocus
            Exit Sub
        Else
            mailerror = checkMailVal(multi.Text) ' call mail validation function
            If mailerror <> "" Then ' if invalid mail entered
                MsgBox mailerror, vbOKOnly + vbCritical, "Monvola - Error"
                multi.BackColor = &HFFFF&
                multi.SetFocus
                Exit Sub
            End If
        End If
        
Case "MS": ' In this case this field is not visible
Case "MT":
            ' check if the location field is blank
        If multi.Text = "" Then
            MsgBox "You have not enterd the location of your meeting.", vbOKOnly + vbCritical, "Monvola - Error"
            multi.BackColor = &HFFFF&
            multi.SetFocus
            Exit Sub
        End If
End Select

'save to database
If (MsgBox("Save this reminder entry ?", vbYesNo + vbInformation, "Monvola - Confirm")) = vbYes Then
savealarm (rmcode)
sizan
End If
End Sub

Private Sub N_Image4_Click()
abt.Visible = True
End Sub

Private Sub N_Image5_Click()
Cancel = True
Call systemtray ' Put in the system tray
main.Hide
End Sub

Private Sub Timer1_Timer()
Dim tmptype As String
Dim tmpsubnam As String
Dim tmpcomments As String
Dim tmpmulti As String
Dim tmpday As Integer
Dim tmpmonth As Integer
Dim tmpyear As Integer
Dim tmphour As Integer
Dim tmpminute As Integer
Dim tmpdone As Boolean
Dim AMorPM As String
Dim crnthour As Integer
Dim crntminute As Integer
Dim temptime As String
Dim i As Integer

'This timer checks for any remineder is there to show or not.
'It checks for all the reminders in the "Pending" state only


If rs.RecordCount = o Then Exit Sub

rs.MoveFirst

For i = 1 To rs.RecordCount ' Read all the records from database
tmptype = rs("type")
rmcodepop = tmptype
tmpsubnam = rs("sub_or_name")
tmpcomments = rs("com_or_desc")
tmpmulti = rs("multi_info")
tmpday = rs("alm_day")
tmpmonth = rs("alm_month")
tmpyear = rs("alm_year")
tmphour = rs("alm_hour")
tmpminute = rs("alm_minute")
tmpdone = rs("alm_done")

If tmpdone = False Then ' if reminder is in the "Pending" state
    
'get current system time and break in hour and minute as 24 hour format

temptime = Time
'this is done because if time can be of length 11 or 10
' e.g. "11:35:25 AM" and  "1:20:50 PM"
'so in both the cases we have to get the hour and minute differently
If Len(temptime) = 11 Then
    crnthour = Mid(temptime, 1, 2)
    crntminute = Mid(temptime, 4, 2)
    AMorPM = Mid(temptime, 10, 2)
Else
    crnthour = Mid(temptime, 1, 1)
    crntminute = Mid(temptime, 3, 2)
    AMorPM = Mid(temptime, 9, 2)
End If

'Convert to 24 hour format
If AMorPM = "PM" And crnthour <> 12 Then crnthour = crnthour + 12
If AMorPM = "AM" And crnthour = 12 Then crnthour = 0

' Check if this record is due for popup or not i.e.
' Its time and date is less than the current time and date or not
If (Year(Now) > tmpyear) Or ((Year(Now) = tmpyear) And (Month(Now) > tmpmonth)) Or _
   ((Year(Now) = tmpyear) And (Month(Now) = tmpmonth) And (Day(Now) > tmpday)) Or _
   ((Year(Now) = tmpyear) And (Month(Now) = tmpmonth) And (Day(Now) = tmpday) And (crnthour > tmphour)) Or _
   ((Year(Now) = tmpyear) And (Month(Now) = tmpmonth) And (Day(Now) = tmpday) And (crnthour = tmphour) And (crntminute >= tmpminute)) Then
       'Timer1.Enabled = False
       fillpopup ' fill the popup form
       PopUP.Show ' show popup form
       
       Exit Sub
    End If
End If
rs.MoveNext ' read next record
Next i


End Sub

Private Function fillpopup()
Dim temp As String
Dim temp1 As String

'this  function fills the popup form with the proper entries from the data base
PopUP.Label5.Visible = True
PopUP.multi.Visible = True

temp = rs("alm_month")
temp1 = rs("alm_day")
If Len(temp) = 1 Then temp = "0" & temp
If Len(temp1) = 1 Then temp1 = "0" & temp1
PopUP.aldate.Caption = temp & "/" & temp1 & "/" & rs("alm_year") ' date

temp = rs("alm_minute")
temp1 = rs("alm_hour")
If Len(temp) = 1 Then temp = "0" & temp
If Len(temp1) = 1 Then temp1 = "0" & temp1
PopUP.altime.Caption = temp1 & ":" & temp ' time

PopUP.subnam.Caption = rs("sub_or_name") ' subject or name
PopUP.multi.Caption = rs("multi_info")   ' other information
PopUP.comments.Caption = rs("com_or_desc") ' comments

' fill the heading as per the reminder type
Select Case rmcodepop
Case "BR"
    PopUP.type.Caption = "Birthday Reminder"
    PopUP.Label2.Caption = "Name"
    PopUP.Label5.Caption = "Relation"
    PopUP.showpict.Picture = birthdaypict.Picture
Case "CL"
    PopUP.type.Caption = "Call Reminder"
    PopUP.Label2.Caption = "Name"
    PopUP.Label5.Caption = "Number"
    PopUP.showpict.Picture = callpict.Picture
Case "ML"
    PopUP.type.Caption = " Mail Reminder"
    PopUP.Label2.Caption = "Name"
    PopUP.Label5.Caption = "E-mail"
    PopUP.showpict.Picture = mailpict.Picture
Case "MS"
    PopUP.type.Caption = "Misc. Reminder"
    PopUP.Label2.Caption = "Subject"
    PopUP.Label5.Visible = False
    PopUP.multi.Visible = False
    PopUP.showpict.Picture = miscpict.Picture
Case "MT"
    PopUP.type.Caption = "Meeting Reminder"
    PopUP.Label2.Caption = "Subject"
    PopUP.Label5.Caption = "Location"
    PopUP.showpict.Picture = meetingpict.Picture
End Select

End Function
Private Sub traypict_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If Hex(x) = "1E3C" Then ' if right click
If (MsgBox("Are to sure to exit Monvola ?", vbYesNo + vbInformation, "Monvola - Exit")) = vbYes Then End
End If

If Hex(x) = "1E0F" Then ' if left click
main.Show
End If

End Sub
