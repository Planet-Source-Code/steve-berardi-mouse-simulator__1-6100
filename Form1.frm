VERSION 4.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "CursorSpy"
   ClientHeight    =   1335
   ClientLeft      =   2340
   ClientTop       =   2160
   ClientWidth     =   1950
   Height          =   1740
   Icon            =   "Form1.frx":0000
   Left            =   2280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   1950
   ShowInTaskbar   =   0   'False
   Top             =   1815
   Width           =   2070
   Begin VB.Frame Frame1 
      Caption         =   "Cursor Position"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
      Begin VB.Label lblYpos 
         Caption         =   "0"
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblXpos 
         Caption         =   "0"
         Height          =   255
         Left            =   600
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblY 
         Caption         =   "Y -"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblX 
         Caption         =   "X -"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   960
      Top             =   2400
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub


Private Sub Form_Load()
    FormOnTop Form1, True
End Sub


Private Sub Timer1_Timer()
    Dim dl&
    Dim pt As PointAPI
    
    dl& = GetCursorPos(pt)

    lblXpos.Caption = pt.x
    lblYpos.Caption = pt.y
    
End Sub


