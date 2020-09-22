VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmExtractor 
   Caption         =   "Email - Extractor"
   ClientHeight    =   10920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10695
   Icon            =   "frmExtractor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10920
   ScaleWidth      =   10695
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   10695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10485
      _ExtentX        =   18494
      _ExtentY        =   18865
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Email"
      TabPicture(0)   =   "frmExtractor.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblTotal"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "RichTextBox1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdExtract"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "List1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Drive1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Dir1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "File1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdSaveList"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "chkSorted"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "chkNumbers"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "List2"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "chkAppend"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "chkFileToList"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cmdClear"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "Ignore List"
      TabPicture(1)   =   "frmExtractor.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstIgnore"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdDelete"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdAdd"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdSaveIgnore"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label4"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   495
         Left            =   360
         TabIndex        =   5
         ToolTipText     =   "Clearing the source."
         Top             =   9600
         Width           =   1095
      End
      Begin VB.CheckBox chkFileToList 
         Caption         =   "File To List"
         Height          =   255
         Left            =   8880
         TabIndex        =   4
         ToolTipText     =   $"frmExtractor.frx":047A
         Top             =   3330
         Width           =   1215
      End
      Begin VB.CheckBox chkAppend 
         Caption         =   "Append"
         Height          =   255
         Left            =   7080
         TabIndex        =   9
         ToolTipText     =   "Extract other files to the E-mail list with or without clearing this E-mail list first."
         Top             =   10170
         Width           =   1215
      End
      Begin VB.ListBox lstIgnore 
         Height          =   6495
         Left            =   -72480
         Sorted          =   -1  'True
         TabIndex        =   11
         Top             =   960
         Width           =   5055
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   495
         Left            =   -70560
         TabIndex        =   13
         ToolTipText     =   "Delete E-mail address from Ignore list."
         Top             =   8160
         Width           =   1215
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   495
         Left            =   -72480
         TabIndex        =   12
         ToolTipText     =   "Add E-mail address to Ignore list."
         Top             =   8160
         Width           =   1215
      End
      Begin VB.CommandButton cmdSaveIgnore 
         Caption         =   "Save"
         Height          =   495
         Left            =   -68640
         TabIndex        =   14
         ToolTipText     =   "Save the Ignore list."
         Top             =   8160
         Width           =   1215
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         Height          =   5490
         Left            =   7080
         Sorted          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   3720
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.CheckBox chkNumbers 
         Caption         =   "Numbers"
         Height          =   255
         Left            =   7080
         TabIndex        =   8
         ToolTipText     =   "Save the E-mail list with or without numbers."
         Top             =   9810
         Width           =   1095
      End
      Begin VB.CheckBox chkSorted 
         Caption         =   "Sorted"
         Height          =   255
         Left            =   7080
         TabIndex        =   7
         ToolTipText     =   "View/Save : Sorted or unsorted E-mail list."
         Top             =   9450
         Width           =   975
      End
      Begin VB.CommandButton cmdSaveList 
         Caption         =   "Save"
         Height          =   495
         Left            =   8880
         TabIndex        =   10
         ToolTipText     =   "Save the E-mail list"
         Top             =   9600
         Width           =   1215
      End
      Begin VB.FileListBox File1 
         Appearance      =   0  'Flat
         Height          =   2565
         Left            =   3600
         Pattern         =   "*.txt;*.rtf;*.htm;*.html;*.asp"
         TabIndex        =   3
         Top             =   600
         Width           =   6495
      End
      Begin VB.DirListBox Dir1 
         Height          =   2115
         Left            =   360
         TabIndex        =   2
         Top             =   1080
         Width           =   3015
      End
      Begin VB.DriveListBox Drive1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   360
         TabIndex        =   1
         Top             =   600
         Width           =   3015
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   5490
         Left            =   7080
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   3720
         Width           =   3015
      End
      Begin VB.CommandButton cmdExtract 
         Caption         =   "Extract ..."
         Height          =   495
         Left            =   1680
         TabIndex        =   6
         ToolTipText     =   "Extract all the E-mail addresses from the source into the E-mail list."
         Top             =   9600
         Width           =   1095
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   5535
         Left            =   360
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Load a file or paste text into this box where you wanna get the E-mail addresses from."
         Top             =   3720
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   9763
         _Version        =   393217
         ScrollBars      =   3
         Appearance      =   0
         TextRTF         =   $"frmExtractor.frx":0531
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "E-mail addresses that not may extracted from source."
         Height          =   195
         Left            =   -72480
         TabIndex        =   22
         Top             =   600
         Width           =   3720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "E-mail list:"
         Height          =   195
         Left            =   7080
         TabIndex        =   21
         Top             =   3360
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Source:"
         Height          =   195
         Left            =   360
         TabIndex        =   20
         Top             =   3360
         Width           =   555
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   3480
         TabIndex        =   19
         Top             =   9750
         Width           =   90
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total :"
         Height          =   195
         Left            =   3000
         TabIndex        =   18
         Top             =   9750
         Width           =   450
      End
   End
End
Attribute VB_Name = "frmExtractor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim blnIgnoreChanged As Boolean

Private Sub chkSorted_Click()
    On Error Resume Next
    List1.ListIndex = 0
    List2.ListIndex = 0
    If chkSorted.Value = 1 Then
        List1.Visible = False
        List2.Visible = True
        List2.SetFocus
    Else
        List2.Visible = False
        List1.Visible = True
        List1.SetFocus
    End If
End Sub

Private Sub cmdAdd_Click()
    Dim strItem As String
    strItem = InputBox("Add To List", "Email - Extractor")
    If strItem <> "" Then
        If checkIfEmail(strItem) = True Then
            lstIgnore.AddItem strItem
            blnIgnoreChanged = True
        End If
    End If
End Sub

Private Sub cmdClear_Click()
    RichTextBox1.Text = ""
End Sub

Private Sub cmdDelete_Click()
    If MsgBox("Delete " & lstIgnore.List(lstIgnore.ListIndex), vbOKCancel + vbQuestion, "Email - Extractor") = vbOK Then
        lstIgnore.RemoveItem lstIgnore.ListIndex
        blnIgnoreChanged = True
    End If
End Sub

Private Sub cmdExtract_Click()
    If RichTextBox1.Text = "" Then
        MsgBox "Load a source-file", vbOKOnly + vbExclamation, "Email - Extractor"
        Exit Sub
    End If
    Me.MousePointer = vbHourglass
    Dim DataString As String
    ReDim SubStr(0) As String
    Dim SubStrCount As Long
    Dim l, lEnter As Long
    Dim myEmail As String
    If chkAppend.Value <> 1 Then
        List1.Clear
        List2.Clear
        lblTotal.Caption = 0
    End If
    ' Create a comma-delimited string:
    DataString = RichTextBox1.Text
    For l = 1 To Len(DataString)
        If Mid(DataString, l, 2) = vbCrLf Then
            lEnter = lEnter + 1
        End If
    Next
    For l = 1 To Len(DataString) - lEnter
        If Mid(DataString, l, 2) = vbCrLf Then
            DataString = Left(DataString, l - 1) & " " & Right(DataString, Len(DataString) - l - 1)
        End If
    Next
    ' Parse the string into sub-strings:
    SubStrCount = ParseString(SubStr(), DataString)
    ' Display the sub-strings:
    Dim blnIgnore As Boolean
    Dim lCheck As Long
    Dim lNewMails As Long
    lNewMails = List1.ListCount
    For l = 1 To SubStrCount
        myEmail = GetAddress(SubStr(l))
        blnIgnore = False
        For lCheck = 0 To lstIgnore.ListCount - 1
            If myEmail = lstIgnore.List(lCheck) Then
                blnIgnore = True
            End If
        Next
        If blnIgnore = False Then
            If checkIfEmail(myEmail) Then
                AddUnique myEmail, List1
                AddUnique myEmail, List2
                lblTotal.Caption = List1.ListCount
                Me.Refresh
                DoEvents
            End If
        End If
    Next
    On Error Resume Next
    List1.ListIndex = 0
    List2.ListIndex = 0
    If chkSorted.Value = 1 Then
        List2.SetFocus
    Else
        List1.SetFocus
    End If
    Me.MousePointer = vbDefault
    lNewMails = List1.ListCount - lNewMails
    MsgBox lNewMails & " new E-mail adresses extracted", vbOKOnly + vbInformation, "Email - Extractor"
End Sub

Private Sub cmdSaveIgnore_Click()
    If lstIgnore.ListCount = 0 Then
        Exit Sub
    End If
    Dim FileNum As Long
    Dim sPath As String
    Dim i As Long
    FileNum = FreeFile()
    sPath = App.Path & "\Ignore.txt"
    Open (sPath) For Output As #FileNum
    For i = 0 To lstIgnore.ListCount - 1
        Print #FileNum, lstIgnore.List(i)
    Next
    Close #FileNum
    blnIgnoreChanged = False
    File1.Refresh
End Sub

Private Sub cmdSaveList_Click()
    If List1.ListCount = 0 Then
        MsgBox "Extract the E-MAILS first", vbOKOnly + vbCritical, "Email - Extractor"
        Exit Sub
    End If
    Dim FileNum As Long
    Dim sPath As String
    Dim i As Long
    FileNum = FreeFile()
    sPath = App.Path & "\Emails Extracted.txt"
    Me.MousePointer = vbHourglass
    Open (sPath) For Output As #FileNum
    If chkNumbers.Value = 1 Then
        If chkSorted.Value = 1 Then
            For i = 0 To List2.ListCount - 1
                If i < 9 Then
                    Print #FileNum, "0000" & i + 1 & "  " & (List2.List(i))
                ElseIf i < 99 And i > 8 Then
                    Print #FileNum, "000" & i + 1 & "  " & (List2.List(i))
                ElseIf i < 999 And i > 98 Then
                    Print #FileNum, "00" & i + 1 & "  " & (List2.List(i))
                ElseIf i < 9999 And i > 998 Then
                    Print #FileNum, "0" & i + 1 & "  " & (List2.List(i))
                Else
                    Print #FileNum, i + 1 & "  " & (List2.List(i))
                End If
            Next
        Else
            For i = 0 To List1.ListCount - 1
                If i < 9 Then
                    Print #FileNum, "0000" & i + 1 & "  " & (List1.List(i))
                ElseIf i < 99 And i > 8 Then
                    Print #FileNum, "000" & i + 1 & "  " & (List1.List(i))
                ElseIf i < 999 And i > 98 Then
                    Print #FileNum, "00" & i + 1 & "  " & (List1.List(i))
                ElseIf i < 9999 And i > 998 Then
                    Print #FileNum, "0" & i + 1 & "  " & (List2.List(i))
                Else
                    Print #FileNum, i + 1 & "  " & (List1.List(i))
                End If
            Next
        End If
    Else
        If chkSorted.Value = 1 Then
            For i = 0 To List2.ListCount - 1
                Print #FileNum, List2.List(i)
            Next
        Else
            For i = 0 To List1.ListCount - 1
                Print #FileNum, List1.List(i)
            Next
        End If
    End If
    Close #FileNum
    Me.MousePointer = vbDefault
    File1.Refresh
    MsgBox "E-MAILS saved into : " & sPath, vbOKOnly + vbInformation, "Email - Extractor"
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
    'Debug.Print File1.Path
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
    If chkAppend.Value <> 1 Then
        List1.Clear
        List2.Clear
        lblTotal.Caption = 0
    End If
    If chkFileToList.Value <> 1 Then
        RichTextBox1.LoadFile (Dir1.Path & "\" & File1.FileName)
    Else
        Dim FileNum As Long
        Dim strData, sPath As String
        Dim blnFirstLine, blnNumbers, blnFourZero As Boolean
        blnFirstLine = True
        If File1.FileName = "Ignore.txt" Then
            Exit Sub
        End If
        sPath = Dir1.Path & "\" & File1.FileName
        FileNum = FreeFile()
        Me.MousePointer = vbHourglass
        Open (sPath) For Input As #FileNum
            Do Until EOF(FileNum)
                Line Input #FileNum, strData
                If blnFirstLine = True Then
                    If Left(strData, 5) = "0001 " Or Left(strData, 6) = "00001 " Then
                        If Left(strData, 5) = "0001 " Then
                            blnFourZero = True
                        End If
                        blnNumbers = True
                    End If
                    blnFirstLine = False
                End If
                If blnNumbers = True Then
                    If blnFourZero = False Then
                        If Left(strData, 5) > 99999 Then
                            strData = Mid(strData, 9)
                        Else
                            strData = Mid(strData, 8)
                        End If
                    Else
                        If Left(strData, 5) > 9999 Then
                            strData = Mid(strData, 8)
                        Else
                            strData = Mid(strData, 7)
                        End If
                    End If
                Else
                    If checkIfEmail(strData) = False Then
                        Close #FileNum
                        Me.MousePointer = vbDefault
                        Exit Sub
                    End If
                End If
                List1.AddItem strData
                List2.AddItem strData
                lblTotal.Caption = List1.ListCount
                Me.Refresh
                DoEvents
            Loop
        Close #FileNum
        Me.MousePointer = vbDefault
    End If
End Sub

Private Sub File1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        If File1.FileName = "Ignore.txt" Then
            If MsgBox(File1.FileName & " holds you're ignored e-mails, are you sure you wanna delete it", vbOKCancel + vbCritical + vbDefaultButton2, "Email Extractor") = vbOK Then
                Kill Dir1.Path & "\" & File1.FileName
                File1.Refresh
                RichTextBox1.Text = ""
            End If
        Else
            If MsgBox("Delete " & File1.FileName, vbOKCancel + vbQuestion, "Email Extractor") = vbOK Then
                Kill Dir1.Path & "\" & File1.FileName
                File1.Refresh
                RichTextBox1.Text = ""
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo errHandle
    Dim FileNum As Long
    Dim strData, sPath As String
    FileNum = FreeFile()
    sPath = App.Path & "\Ignore.txt"
    Open (sPath) For Input As #FileNum
        Do Until EOF(FileNum)
            Line Input #FileNum, strData
            lstIgnore.AddItem strData
        Loop
    Close #FileNum
    Exit Sub
errHandle:
    Open (sPath) For Output As #FileNum
        Print #FileNum, "Ignored Emails"
    Close #FileNum
    File1.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If blnIgnoreChanged = True Then
        If MsgBox("Ignore List has changed, save it", vbOKCancel + vbQuestion, "Email Extractor") = vbOK Then
            cmdSaveIgnore_Click
        End If
    End If
End Sub

