VERSION 4.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Mouse Simulator"
   ClientHeight    =   2250
   ClientLeft      =   4380
   ClientTop       =   2745
   ClientWidth     =   5010
   ControlBox      =   0   'False
   Height          =   2655
   Left            =   4320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   Top             =   2400
   Width           =   5130
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "e-mail:  sberardi@smblib.8m.com"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   "website:  smblib.8m.com"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Programmed by Steve Berardi"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Version 1.0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   600
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmAbout.frx":0000
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Mouse Simulator"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

