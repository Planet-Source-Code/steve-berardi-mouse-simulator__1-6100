Attribute VB_Name = "Module1"
'Mouse Simulator (version 1.0)
'By Steve Berardi
'Copyright (c) 2000 by Steve Berardi
'***************************************
'All code by Steve Berardi
'Compiled: February  11, 2000
'***************************************

Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)


Type POINTAPI
        x As Long
        y As Long
End Type


Public Const MOUSEEVENTF_ABSOLUTE = &H8000
Public Const MOUSEEVENTF_LEFTDOWN = &H2
Public Const MOUSEEVENTF_LEFTUP = &H4
Public Const MOUSEEVENTF_MIDDLEDOWN = &H20
Public Const MOUSEEVENTF_MIDDLEUP = &H40
Public Const MOUSEEVENTF_MOVE = &H1
Public Const MOUSEEVENTF_RIGHTDOWN = &H8
Public Const MOUSEEVENTF_RIGHTUP = &H10

Sub EndActions()
    
    Pause 1
    
End Sub

Sub LeftClickM(xP As Long, yP As Long)
Dim junk As Long

    junk = SetCursorPos(xP, yP)

    Pause 0.2
    
    mouse_event MOUSEEVENTF_LEFTDOWN, xP, yP, 0, 0
    mouse_event MOUSEEVENTF_LEFTUP, xP, yP, 0, 0
    
End Sub


Sub LeftDownM(xP As Long, yP As Long)
Dim junk As Long

    junk = SetCursorPos(xP, yP)

    Pause 0.2
    
    mouse_event MOUSEEVENTF_LEFTDOWN, xP, yP, 0, 0

End Sub

Sub LeftUpM(xP As Long, yP As Long)
Dim junk As Long

    junk = SetCursorPos(xP, yP)

    Pause 0.2
    
    mouse_event MOUSEEVENTF_LEFTUP, xP, yP, 0, 0
    
End Sub


Sub Pause(ByVal nSecond As Single)
'creates a pause in program, in seconds
'works well with timers and slowing a program down
'Note: you may pass a decimal through the interval argument
   
   Dim tOut As Single
   tOut = Timer
   
   Do While Timer - tOut < nSecond
   
      Dim junk As Integer
      
      junk = DoEvents()
      
      If Timer < tOut Then
         tOut = tOut - CLng(24) * CLng(60) * CLng(60)
      End If
      
   Loop
   
End Sub
Sub PauseAction(interval As Integer)

    Pause interval
    
End Sub

Sub RightClickM(xP As Long, yP As Long)
Dim junk As Long

    junk = SetCursorPos(xP, yP)

    Pause 0.2
    
    mouse_event MOUSEEVENTF_RIGHTDOWN, xP, yP, 0, 0
    mouse_event MOUSEEVENTF_RIGHTUP, xP, yP, 0, 0
    
End Sub


Sub RightDownM(xP As Long, yP As Long)
Dim junk As Long

    junk = SetCursorPos(xP, yP)

    Pause 0.2
    
    mouse_event MOUSEEVENTF_RIGHTDOWN, xP, yP, 0, 0
  
End Sub


Sub RightUpM(xP As Long, yP As Long)
Dim junk As Long

    junk = SetCursorPos(xP, yP)

    Pause 0.2
    
    mouse_event MOUSEEVENTF_RIGHTUP, xP, yP, 0, 0
  
End Sub



