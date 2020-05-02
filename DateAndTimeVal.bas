Attribute VB_Name = "timesettings"
Public dateerror As String
Public timeerror As String
Public today As Boolean
Public Function checkdate(indate As String)
Dim temp1, tmpmonth, tmpday, tmpyear As String
Dim tm As Integer
Dim ty As Integer
Dim td As Integer

today = False
'check if all the characters entered or not
For i = 1 To Len(indate)
    If Mid(indate, i, 1) = "_" Then
        dateerror = "Character no. " & i & " not filled. Full date should be filled."
        Exit Function
    End If
Next i
' Split the input date into year, month and day
tmpmonth = Mid(indate, 1, 2)
tmpday = Mid(indate, 4, 2)
tmpyear = Mid(indate, 7, 4)

'check month
If tmpmonth = 0 Or tmpmonth > 12 Then
    dateerror = "Month can not be zero or more than 12."
    Exit Function
End If

'check day for zero
If tmpday = 0 Then
    dateerror = "Day can not be zero."
    Exit Function
End If

'move in the integer fields
tm = tmpmonth
td = tmpday
ty = tmpyear

'check days according to month

If tm = 1 Or tm = 3 Or tm = 5 Or tm = 7 Or tm = 8 Or tm = 10 Or tm = 12 Then
    If tmpday > 31 Then
        dateerror = "Number of days can not be greater than 31 in this month."
        Exit Function
    End If
End If

If tm = 4 Or tm = 6 Or tm = 9 Or tm = 11 Then
    If tmpday > 30 Then
        dateerror = "Number of days can not be greater than 30 in this month."
        Exit Function
    End If
End If

If tm = 2 Then

'Leap year validations
    If (tmpyear Mod 400) = 0 Then ' if divisible by 400 then max days = 29
        If tmpday > 29 Then
            dateerror = "For this year days in Feb. can not be more than 29."
            Exit Function
        End If
    Else
        If (tmpyear Mod 100) = 0 Then 'if divisible by 100 and not by 400 then = 28
            If tmpday > 28 Then
                dateerror = "For this year days in Feb. can not be more than 28."
                Exit Function
            End If
        Else
            If (tmpyear Mod 4) = 0 Then ' if by 4 but not by 100 and 400 then = 29
                If tmpday > 29 Then
                    dateerror = "For this year days in Feb. can not be more than 29."
                    Exit Function
                End If
            Else
                If tmpday > 28 Then ' if not by 4 and 100 and 400 then = 28
                    dateerror = "For this year days in Feb. can not be more than 28."
                    Exit Function
                End If
            End If
        End If
    End If
    
End If

'check if date enterd is less then the system date
If Year(Now) > ty Then ' if enterd year less than current year
    dateerror = "You can not set alarm for back dates."
    Exit Function
Else
    If Year(Now) = ty Then ' if entered year equal to current year
        If Month(Now) > tm Then ' if entered month less than current month
            dateerror = "You can not set alarm for back dates."
            Exit Function
        Else
            If Month(Now) = tm Then 'if entered month equal to current month
                If Day(Now) > td Then ' if entered day less than current day
                    dateerror = "You can not set alarm for back dates."
                    Exit Function
                Else
                    ' if day also equal then it is today
                    If Day(Now) = td Then today = True
                End If
            End If
        End If
    End If
End If

End Function

Public Function checktime(intime As String)
Dim temp1, hour, minute As String
Dim th As Integer
Dim tm As Integer
Dim AMorPM As String
Dim crnthour As Integer
Dim crntminute As Integer
Dim temptime As String

'Check if all the character are filled in the time field or not
For i = 1 To Len(intime)
    If Mid(intime, i, 1) = "_" Then
        timeerror = "Character no. " & i & " not filled. Full time should be filled."
        Exit Function
    End If
Next i
' Split input time into hour and minute
th = Mid(intime, 1, 2)
tm = Mid(intime, 4, 2)

'check for invalid hour
If th > 23 Then
    timeerror = "Hour can not be greater than 23."
    Exit Function
End If
'Check for invalid minute
If tm > 59 Then
    timeerror = "Minutes can not be greater than 59."
    Exit Function
End If

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

If AMorPM = "PM" And crnthour <> 12 Then crnthour = crnthour + 12
If AMorPM = "AM" And crnthour = 12 Then crnthour = 0

'check if the entered time is less than current system time
'this need to checked if the date entered is today's date
If today = True Then
    If crnthour > th Then ' if entered hour less than current hour
        timeerror = "You can not set alarm for back time."
        Exit Function
    Else
        If crnthour = th Then ' if entered hour equal to current hour
            If crntminute > tm Then ' if entered minute less than current minute
                timeerror = "You can not set alarm for back time."
            Exit Function
            End If
        End If
    End If
End If
End Function
