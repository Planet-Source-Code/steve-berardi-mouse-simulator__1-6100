VERSION 4.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mouse Simulator"
   ClientHeight    =   4020
   ClientLeft      =   1170
   ClientTop       =   1575
   ClientWidth     =   6135
   Height          =   4425
   Icon            =   "frmMain.frx":0000
   Left            =   1110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   6135
   Top             =   1230
   Width           =   6255
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   4440
      TabIndex        =   17
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   375
      Left            =   2640
      TabIndex        =   16
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton cmdRunC 
      Caption         =   "Run Cursor Spy"
      Height          =   375
      Left            =   2640
      TabIndex        =   15
      Top             =   3000
      Width           =   1575
   End
   Begin VB.ListBox lstCode 
      Height          =   1425
      Left            =   2760
      TabIndex        =   14
      Top             =   4440
      Width           =   1935
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save List"
      Height          =   375
      Left            =   1440
      TabIndex        =   13
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load List"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmdClearList 
      Caption         =   "Clear List"
      Height          =   375
      Left            =   4440
      TabIndex        =   11
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start Loop"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Frame fraActions 
      Caption         =   "Action List:"
      Height          =   2775
      Left            =   2640
      TabIndex        =   8
      Top             =   120
      Width           =   3375
      Begin VB.ListBox lstActions 
         Height          =   2400
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame fraSettings 
      Caption         =   "Settings:"
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
      Begin VB.CommandButton cmdAction 
         Caption         =   "Add Action"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   2160
         Width           =   2175
      End
      Begin VB.ComboBox cboActions 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1560
         Width           =   1935
      End
      Begin VB.TextBox txtYpos 
         Height          =   285
         Left            =   960
         MaxLength       =   4
         TabIndex        =   4
         Text            =   "0"
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtXpos 
         Height          =   285
         Left            =   960
         MaxLength       =   4
         TabIndex        =   3
         Text            =   "0"
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblAction 
         Caption         =   "Action:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label lblY 
         Caption         =   "Y Pos :"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblX 
         Caption         =   "X Pos :"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   615
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1080
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_Creatable = False
Attribute VB_Exposed = False


Private Sub cmdAbout_Click()
    frmAbout.Show
End Sub


Private Sub cmdAction_Click()
Dim Xis, Yis As String
Dim code, endCode As String
Dim duration As String

    duration = "0"
    
    Xis = txtXpos.Text
    Yis = txtYpos.Text
    
Select Case Len(Xis)
    Case 1: Xis = "000" & Xis
    Case 2: Xis = "00" & Xis
    Case 3: Xis = "0" & Xis
End Select
    
Select Case Len(Yis)
    Case 1: Yis = "000" & Yis
    Case 2: Yis = "00" & Yis
    Case 3: Yis = "0" & Yis
End Select
    
        
Select Case cboActions.Text
    
    Case "< Left Click >"
        code = "10"
        
    Case "< Right Click >"
        code = "01"
    
    Case "[ Pause Action ]"
        code = "00"
        duration = InputBox$("Please enter the length of time (in seconds) you want to pause the actions (max of three digits, including a decimal): ", "Enter Pause Time")
        Select Case Len(duration)
            Case 1: duration = "00" & duration
            Case 2: duration = "0" & duration
        End Select
        
    Case "[ End Actions  ]"
        code = "11"
        
    Case "-----------------------------------"
        MsgBox "Please select a valid action.", 48, "Actions"
        Exit Sub
        
    Case ""
        MsgBox "Please select a valid action.", 48, "Actions"
        Exit Sub
        
End Select

If duration = "0" Then duration = "000"

endCode = code & "-" & Xis & "-" & Yis & "-" & duration
    
'add to list boxes:
If duration = "000" Then
lstActions.AddItem cboActions.Text & " at (" & txtXpos.Text & "," & txtYpos.Text & ")"
Else:
lstActions.AddItem cboActions.Text & " for " & duration & " seconds " & " at (" & txtXpos.Text & "," & txtYpos.Text & ")"
End If

lstCode.AddItem endCode
        
End Sub

Private Sub cmdClearList_Click()
    lstActions.Clear
    lstCode.Clear
End Sub



Private Sub cmdExit_Click()
    Unload Me
    End
End Sub

Private Sub cmdLoad_Click()
'Open file in list:
    CommonDialog1.Filter = "Mouse Simulator Action List (*.sbl)|*.sbl|All Files (*.*)|*.*"
    CommonDialog1.ShowOpen
    
    fname$ = CommonDialog1.filename
    
    If fname$ = "" Then Exit Sub

Dim ActualCode As Boolean

ActualCode = False
lstActions.Clear
lstCode.Clear

op:
On Error GoTo er
Open fname$ For Input As #1
Input #1, A$
lstActions.AddItem A$
Do Until (EOF(1) = True)
    Input #1, A$
    If A$ = "-" Then
        ActualCode = True
        GoTo act
    End If
    Select Case ActualCode
        Case True: lstCode.AddItem A$
        Case False: lstActions.AddItem A$
    End Select

act:
Loop
Close #1
Exit Sub


er:
MsgBox "File Not Found.", 48, "Error":
Exit Sub

End Sub

Private Sub cmdRunC_Click()
    Dim junkStuff As Long
    
    junkStuff = Shell(App.Path & "\cspy.exe", 1)
    
End Sub

Private Sub cmdSave_Click()
'Save File in list:
Dim TotalItems, n As Long

    CommonDialog1.Filter = "Mouse Simulator Action List (*.sbl)|*.sbl|All Files (*.*)|*.*"
    CommonDialog1.ShowSave
    
    fname$ = CommonDialog1.filename
    If fname$ = "" Then Exit Sub

TotalItems = (lstActions.ListCount - 1)

'save file:
On Error GoTo er
Open fname$ For Output As #1

For n = 0 To TotalItems Step 1
    Write #1, lstActions.List(n)
Next n

Write #1, "-"

For n = 0 To TotalItems Step 1
    Write #1, lstCode.List(n)
Next n

Close #1
Exit Sub


er:
MsgBox "File Not Found.", 22, "Error":
Exit Sub
End Sub

Private Sub cmdStart_Click()
    Dim NextAction As String
    Dim code As String
    Dim Xvalue, Yvalue As Long
    Dim pTime As Integer
    Dim n, HowManyTimes As Long
    Dim ItemIndex, NextItem As Long
    
    HowManyTimes = InputBox("Please enter how many times you want these actions to loop (from 1 to 2,000,000,000): ", "Enter Times To Loop")
    
'Begin Loop:
For n = 0 To HowManyTimes Step 1

ItemIndex = (lstCode.ListCount - 1)

'Read the code list box:
For NextItem = 0 To ItemIndex Step 1

    NextAction = lstCode.List(NextItem) 'gets next action as whole code
    
    'get values of each code:
    code = Left$(NextAction, 2) 'action to perform
    Xvalue = Val(Mid(NextAction, 4, 4)) 'X value
    Yvalue = Val(Mid(NextAction, 9, 4)) 'Y value
    pTime = Val(Right(NextAction, 3)) 'duration value
    
Select Case code
    
    Case "10"
        LeftClickM (Xvalue), (Yvalue)
        
    Case "01"
        RightClickM (Xvalue), (Yvalue)
    
    Case "00"
        PauseAction (pTime)
    
    Case "11"
        EndActions
        
End Select

Next NextItem
    
Next n

End Sub

Private Sub Form_Load()

    cboActions.AddItem "< Left Click >"
    cboActions.AddItem "-----------------------------------"
    cboActions.AddItem "< Right Click >"
    cboActions.AddItem "-----------------------------------"
    cboActions.AddItem "[ Pause Action ]"
    cboActions.AddItem "[ End Actions  ]"
    cboActions.AddItem "-----------------------------------"
    
End Sub


