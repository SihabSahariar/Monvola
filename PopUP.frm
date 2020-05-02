VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form PopUP 
   BorderStyle     =   0  'None
   Caption         =   "E - Minder -  Alert"
   ClientHeight    =   5550
   ClientLeft      =   0
   ClientTop       =   -60
   ClientWidth     =   7770
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "PopUP.frx":0000
   ScaleHeight     =   5550
   ScaleWidth      =   7770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox showpict 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   530
      Left            =   6720
      ScaleHeight     =   525
      ScaleMode       =   0  'User
      ScaleWidth      =   525
      TabIndex        =   16
      Top             =   3840
      Visible         =   0   'False
      Width           =   530
   End
   Begin VB.ComboBox comremind 
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
      Left            =   4680
      TabIndex        =   13
      ToolTipText     =   "Select time to remind after"
      Top             =   4920
      Width           =   615
   End
   Begin VB.CommandButton cmdremind 
      Caption         =   "Remind Me"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      TabIndex        =   12
      ToolTipText     =   "Click to remind again"
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmddone 
      Caption         =   " OK"
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
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      ToolTipText     =   "Click to mark as done."
      Top             =   4800
      Width           =   1095
   End
   Begin WMPLibCtl.WindowsMediaPlayer w 
      Height          =   855
      Left            =   9000
      TabIndex        =   17
      Top             =   2400
      Width           =   1455
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   2566
      _cy             =   1508
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "after"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4080
      TabIndex        =   15
      Top             =   4920
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Minutes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5520
      TabIndex        =   14
      Top             =   4920
      Width           =   735
   End
   Begin VB.Label comments 
      BackStyle       =   0  'Transparent
      Caption         =   "comments"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   1560
      TabIndex        =   11
      Top             =   3600
      Width           =   5055
   End
   Begin VB.Label multi 
      BackStyle       =   0  'Transparent
      Caption         =   "multi"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2280
      TabIndex        =   10
      Top             =   2160
      Width           =   3735
   End
   Begin VB.Label altime 
      BackStyle       =   0  'Transparent
      Caption         =   "altime"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5640
      TabIndex        =   9
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label aldate 
      BackStyle       =   0  'Transparent
      Caption         =   "Aldate"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2280
      TabIndex        =   8
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label subnam 
      BackStyle       =   0  'Transparent
      Caption         =   "Sizan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2280
      TabIndex        =   7
      Top             =   1080
      Width           =   3735
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
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3240
      TabIndex        =   6
      Top             =   3120
      Width           =   1095
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
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   2160
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
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4920
      TabIndex        =   4
      Top             =   1680
      Width           =   615
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
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   1680
      Width           =   615
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
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label type 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Birthday Reminder"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   -480
      TabIndex        =   0
      Top             =   120
      Width           =   8175
   End
End
Attribute VB_Name = "PopUP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmddone_Click()
'Mark this entry as done
rs.Edit
rs("alm_done") = True
rs.Update
Unload Me
End Sub

Private Sub cmdremind_Click()

If comremind.Text = "" Then
MsgBox "Select a value from the list ? ", vbOKOnly + vbCritical, "Monvola - Error"
Exit Sub
End If

If remind = False Then  'call function "remind"
rs.Edit
rs("alm_done") = True
Else
rs("alm_done") = False
End If

'update the database
rs.Update
Unload Me
End Sub

Private Sub comremind_click()
'Enable the "Remind" button only when user selects an entry from combo box
If comremind.Text <> "" Then
cmdremind.Enabled = True
Else
cmdremind.Enabled = False
End If

End Sub

Private Sub Form_Load()
'Add items in the combo box
comremind.AddItem 10
comremind.AddItem 20
comremind.AddItem 30
comremind.AddItem 40
comremind.AddItem 50
comremind.AddItem 60
cmdremind.Enabled = False
main.Timer1.Enabled = False
SetWindowPos PopUP.hwnd, -1, 0, 0, 0, 0, &H1 Or &H2
w.URL = App.Path & "/m.wav"
End Sub

Private Function remind() As Boolean

Dim rmyr As Integer
Dim rmmon As Integer
Dim rmday As Integer
Dim rmmin As Integer
Dim rmhour As Integer
Dim rmAMorPM As String
Dim increment As Integer
Dim temp As Integer
Dim temptime As String

If comremind.Text = "" Then
remind = False
Exit Function
End If

'get current system time and break in hour and minute as 24 hour format
temptime = Time
'this is done because if time can be of length 11 or 10
' e.g. "11:35:25 AM" and  "1:20:50 PM"
'so in both the cases we have to get the hour and minute differently
If Len(temptime) = 11 Then
    rmhour = Mid(temptime, 1, 2)
    rmmin = Mid(temptime, 4, 2)
    rmAMorPM = Mid(temptime, 10, 2)
Else
    rmhour = Mid(temptime, 1, 1)
    rmmin = Mid(temptime, 3, 2)
    rmAMorPM = Mid(temptime, 9, 2)
End If

'Convert to 24 hour format
If rmAMorPM = "PM" And rmhour <> 12 Then rmhour = rmhour + 12
If rmAMorPM = "AM" And rmhour = 12 Then rmhour = 0

'get system date
rmyr = Year(Now)
rmmon = Month(Now)
rmday = Day(Now)
increment = comremind.Text

temp = rmmin
rmmin = rmmin + increment

If rmmin > 59 Then ' if minutes greater than 59 then increment hour
    rmhour = rmhour + 1
    rmmin = (increment - (60 - temp)) ' increment minutes accordingly
    If rmhour > 23 Then ' if hour more than 23
        rmday = rmday + 1   ' increment day and hour will be zero
        rmhour = 0
        Select Case rmmon   ' check that after increment in the day it is proper or not
        Case 1 Or 3 Or 5 Or 7 Or 8 Or 10:
            If rmday > 31 Then
                rmmon = rmmon + 1
                rmday = 1
            End If
        Case 4 Or 6 Or 9 Or 11:
            If rmday > 30 Then
                rmmon = rmmon + 1
                rmday = 1
            End If
            
        Case 2:
            If (rmyr Mod 400) = 0 Then ' if divisible by 400 then max days = 29
                If rmday > 29 Then
                    rmmon = rmmon + 1
                    rmday = 1
                End If
            Else
                If (rmyr Mod 100) = 0 Then 'if divisible by 100 and not by 400 then = 28
                    If rmday > 28 Then
                        rmmon = rmmon + 1
                        rmday = 1
                    End If
                Else
                    If (rmyr Mod 4) = 0 Then ' if by 4 but not by 100 and 400 then = 29
                        If rmday > 29 Then
                            rmmon = rmmon + 1
                            rmday = 1
                        End If
                    Else
                        If rmday > 28 Then ' if not by 4 and 100 and 400 then = 28
                            rmmon = rmmon + 1
                            rmday = 1
                        End If
                    End If
                End If
            End If
            
        Case 12: ' if december month then increment year
            If rmday > 31 Then
                rmyr = rmyr + 1
                rmmon = 1
                rmday = 1
            End If
        End Select
    End If
End If
' populate database fields
rs.Edit
rs("alm_minute") = rmmin
rs("alm_hour") = rmhour
rs("alm_day") = rmday
rs("alm_month") = rmmon
rs("alm_year") = rmyr
remind = True
End Function

Private Sub Form_Unload(Cancel As Integer)
main.Timer1.Enabled = True
SetWindowPos PopUP.hwnd, -2, 0, 0, 0, 0, &H1 Or &H2
End Sub
