VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   660
      TabIndex        =   0
      Top             =   180
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements iTimer

Private msglStart As Single

Private Sub Form_Initialize()
    SetATimer Me, 3000, 3
    SetATimer Me, 5000, 5
    SetATimer Me, 7000, 7
    SetATimer Me, 15000, 15
    SetATimer Me, 35000, 35
    msglStart = Timer
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    List1.Move 0, 0, Width - 100, Height - 450
End Sub

Private Sub Form_Terminate()
    KillTimers Me
End Sub

Private Sub iTimer_Fire(ByVal Tag As Long)
    List1.AddItem "Timer w/ Interval of " & Tag & "secs. Fired!  " & "Total Time: " & Timer - msglStart
End Sub
