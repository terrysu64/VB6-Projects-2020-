VERSION 5.00
Begin VB.Form frmSandbox 
   Caption         =   "Sandbox - Supporting VB6 Dots"
   ClientHeight    =   6165
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9180
   LinkTopic       =   "Form1"
   ScaleHeight     =   411
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   612
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   5520
      Width           =   8895
   End
   Begin VB.PictureBox picTouchPad 
      Height          =   5295
      Left            =   120
      ScaleHeight     =   349
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   357
      TabIndex        =   1
      Top             =   120
      Width           =   5415
   End
   Begin VB.ListBox lstOutput 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5280
      Left            =   5760
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "frmSandbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Mr. Smith
'Date: May 19, 20220
'Purpose: Sandbox to demonstrate Pset and MouseDown
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
Option Explicit

'GLOBAL
Dim intCountClicks As Integer

Dim intX1 As Integer
Dim intY1 As Integer
Dim intX2 As Integer
Dim intY2 As Integer

Private Sub cmdClear_Click()
picTouchPad.Cls
lstOutput.Clear

picTouchPad.Print "First Click was at " & intX1 & " and " & intY1 '
picTouchPad.Print "Second Click was at " & intX2 & " and " & intY2

End Sub



Private Sub Form_Load()
intCountClicks = 0
intX1 = 0
intY1 = 0
intX2 = 0
intY2 = 0
End Sub

Private Sub picTouchPad_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

intCountClicks = intCountClicks + 1

If intCountClicks <= 4 Then

    If intCountClicks = 1 Then
        intX1 = X
        intY1 = Y
    ElseIf intCountClicks = 2 Then
        intX2 = X
        intY2 = Y
    End If
    
    lstOutput.AddItem "X=" & X & "   " & "Y=" & Y

    picTouchPad.PSet (X, Y), vbRed
    picTouchPad.Circle (X, Y), 3, vbBlue
Else
    MsgBox ("You've clicked 4 times, please stop")
End If

End Sub
