Attribute VB_Name = "MailValidation"

Public Function checkMailVal(getPreVal) As String
        'first check for any invalid character
        Dim parseMail
        
        For parseMail = 1 To Len(getPreVal)
            
            If (chkAccChars(Mid(getPreVal, parseMail, 1))) + _
            (chkAlfaN(Mid(getPreVal, parseMail, 1))) = 0 Then
            
                checkMailVal = "Invalid character in the e-mail"
                Exit Function
            End If
        Next
        
        ' check if @ and . are present in the IDs
        If InStr(1, getPreVal, "@") = 0 Or InStr(1, getPreVal, ".") = 0 Then
            checkMailVal = "The @ and the DOT must be present in the email ID"
            Exit Function
        End If
        
        ' there should be only one @
        Dim atCtr, pos
        atCtr = 0
        For pos = 1 To Len(getPreVal)
                    
            If Mid(getPreVal, pos, 1) = "@" Then
                'increment the atCtr(ctr which
                'counts the @ in the text
                atCtr = atCtr + 1
            End If
            'not more than 1 @ is allowed in the
            'mail address!!
            If atCtr > 1 Then
                checkMailVal = "Not more than one @ is allowed!"
                Exit Function
            End If
        Next
        
        ' Now capture the position of @ and . for later use
        
        Dim atPosGlobal As Integer
        Dim dotPosGlobal() As Integer
        Dim arrCtr As Integer
        Dim parseCtr As Integer
        
        
        ' populate an array with the positions of the @
        For parseCtr = 1 To Len(getPreVal)
            If Mid(getPreVal, parseCtr, 1) = "." Then
                arrCtr = arrCtr + 1
                ReDim Preserve dotPosGlobal(arrCtr)
                    
                dotPosGlobal(arrCtr) = parseCtr
            End If
        Next
        
        ' get the @ position
        atPosGlobal = InStr(1, getPreVal, "@")
        
        
        
        If (chkAlfaN(Mid(getPreVal, 1, 1))) * _
        (chkAlfaN(Mid(getPreVal, Len(getPreVal), 1))) = 0 Then
               'invalid mail
               checkMailVal = "Invalid character at the start or end of the mail ID"
               Exit Function
        End If

        If (chkAlfaN(Mid(getPreVal, atPosGlobal - 1, 1))) * _
        (chkAlfaN(Mid(getPreVal, atPosGlobal + 1, 1))) = 0 Then
               'invalid mail
               checkMailVal = "The @ is placed before or after an invalid character"
               Exit Function
        End If


        'now ensure that the . doesn't repeat itself in sequence
        ' eg> ..(this is wrong) .com(this is right)

        Dim currDotPos, prevDotPos
        For currDotPos = 1 To Len(getPreVal)
            If Mid(getPreVal, currDotPos, 1) = "." Then
                If currDotPos - prevDotPos = 1 Then
                    checkMailVal = "You cannot have the DOT placed in continuous sequence"
                    Exit Function
                Else
                    prevDotPos = currDotPos
                End If
            End If
        Next
        
                
        'the @ should lie next to an alphanumeric character
         
        ' Now check for the @ pos
        ' the @ should typically lie like this x@x.xx
        ' which means the @ should neither lie in the end
        ' or the start of the mail id
                
                'the email is structurely right if it has reached so far
                
                ' next check if the "." lies within the last 4 chars
                               
                Dim revPos, valParse
                
                For valParse = 1 To Len(getPreVal)
                    
                    If Mid(getPreVal, Len(getPreVal) - (valParse - 1), 1) = "." Then
                        If valParse > 4 Then
                            checkMailVal = "The amount of characters placed after the last DOT is incorrect"
                            Exit Function
                        End If
                        Exit For
                    End If
                
                Next
             
                    
                
    
    
    'all is well
    'now return back an empty string
    checkMailVal = ""
    Exit Function
    
End Function

Private Function chkAlfaN(getChar) As Boolean
    Select Case Asc(getChar)
        Case 97 To 122, 65 To 90, 48 To 57
                '97-122(a-z), 65-90(A-Z)
                '48-59(0-9)
            chkAlfaN = True
        Case Else
            chkAlfaN = False
    End Select
    
    
End Function


Private Function chkAccChars(getChar) As Boolean
    Select Case Asc(getChar)
        Case 45 To 47, 64, 95, 126
                '-./@_~
            chkAccChars = True
        Case Else
            chkAccChars = False
    End Select
    
    
End Function

