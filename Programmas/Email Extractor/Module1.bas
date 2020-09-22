Attribute VB_Name = "Module1"
Option Explicit

Public Sub AddUnique(ByVal StringToAdd As String, lst As ListBox)
    lst.Text = StringToAdd
    If lst.ListIndex = -1 Then
        'it does not exist, so add it..
        lst.AddItem StringToAdd
    End If
End Sub

Public Function ParseString(SubStrs() As String, SrcStr As String) As Long
 ' Dimension variables:
      ReDim SubStrs(0) As String
      Dim CurPos As Long
      Dim NextPos As Long
      Dim DelLen As Integer
      Dim nCount As Long
      Dim TStr As String
      Dim Delimiter As String
        
      'Set the delimeter value
      Delimiter = " "
               
      ' Add delimiters to start and end of string to make loop simpler:
      SrcStr = Delimiter & SrcStr & Delimiter
      ' Calculate the delimiter length only once:
      DelLen = Len(Delimiter)
      ' Initialize the count and position:
      nCount = 0
      CurPos = 1
      NextPos = InStr(CurPos + DelLen, SrcStr, Delimiter)
      
      ' Loop searching for delimiters:
      Do Until NextPos = 0
         ' Extract a sub-string:
         TStr = Mid$(SrcStr, CurPos + DelLen, NextPos - CurPos - DelLen)
         
         If TStr <> Delimiter Then
         ' Increment the sub string counter:
         nCount = nCount + 1
         ' Add room for the new sub-string in the array:
         ReDim Preserve SubStrs(nCount) As String
         ' Put the sub-string in the array:
           SubStrs(nCount) = TStr
         End If
         ' Position to the last found delimiter:
         CurPos = NextPos
         ' Find the next delimiter:
         NextPos = InStr(CurPos + DelLen, SrcStr, Delimiter)
       Loop
      
      ' Return the number of sub-strings found:
      
      ParseString = nCount
End Function


Public Function GetAddress(ByVal s As String) As String
    Dim amppos As Integer
    Dim p1 As Integer
    Dim p2 As Integer
    Dim n As Integer
    Dim temp As String
    amppos = InStr(s, "@")
    p1 = 1
    p2 = Len(s)
    GetAddress = ""
    If amppos = 0 Then Exit Function
    
    ' find start of address
    For n = (amppos - 1) To 1 Step -1
        temp = Mid(s, n, 1)
        If (temp = " ") Or (temp = "<") Or (temp = "(") Or (temp = ":") Or (temp = ",") Or (temp = "[") Then
            p1 = n + 1
            Exit For
        End If
    Next
    
   ' find end of address
    For n = (amppos + 1) To Len(s)
        temp = Mid(s, n, 1)
        If (temp = " ") Or (temp = ">") Or (temp = ")") Or (temp = ":") Or (temp = ",") Or (temp = "]") Then
            p2 = n - 1
            Exit For
        End If
    Next
        
    ' make address
    GetAddress = Mid(s, p1, (p2 - p1) + 1)
End Function

Public Function checkIfEmail(ByVal email As String) As Boolean
    Dim i As Integer
    Dim char As String
    Dim c() As String
    'checks if the string has the standard e
    '     mail pattern:
    If Not email Like "*@*.*" Then
        checkIfEmail = False
        Exit Function
    End If
    'not starting with @
    If Left(email, 1) = "@" Then
        checkIfEmail = False
        Exit Function
    End If
    'splits the email-string with a "." deli
    '     meter and returns the subtring in the c-
    '     string array
    c = Split(email, ".", -1, vbBinaryCompare)
    'checks if the last substring has a leng
    '     th of either 2 or 3
    If Not Len(c(UBound(c))) = 3 And Not Len(c(UBound(c))) = 2 Then
        checkIfEmail = False
        Exit Function
    End If
    'steps through the last substring to see
    '     if it contains anything else unless char
    '     acters from a to z


    For i = 1 To Len(c(UBound(c))) Step 1
        char = Mid(c(UBound(c)), i, 1)


        If Not (LCase(char) <= Chr(122)) Or Not (LCase(char) >= Chr(97)) Then
            checkIfEmail = False
            Exit Function
        End If
    Next i
    'steps through the whole email string to
    '     see if it contains any special character
    '     s:


    For i = 1 To Len(email) Step 1
        char = Mid(email, i, 1)
        If (LCase(char) <= Chr(122) And LCase(char) >= Chr(97)) _
        Or (char >= Chr(48) And char <= Chr(57)) _
        Or (char = ".") _
        Or (char = "@") _
        Or (char = "-") _
        Or (char = "_") Then
        checkIfEmail = True
    Else
        checkIfEmail = False
        Exit Function
    End If
Next i
End Function

