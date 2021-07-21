VERSION 5.00
Begin VB.Form frmSierpinskiTriangles 
   Caption         =   "Sierpinski’s Triangles"
   ClientHeight    =   7080
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9180
   LinkTopic       =   "Form1"
   ScaleHeight     =   472
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
      Left            =   4680
      TabIndex        =   2
      Top             =   6360
      Width           =   4335
   End
   Begin VB.CommandButton cmdGo 
      Appearance      =   0  'Flat
      Caption         =   "GO!!!"
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
      TabIndex        =   1
      Top             =   6360
      Width           =   4335
   End
   Begin VB.PictureBox picTouchPad 
      Height          =   5295
      Left            =   120
      ScaleHeight     =   349
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   589
      TabIndex        =   0
      Top             =   840
      Width           =   8895
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Please select 4 points and prepare to be amazed!"
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
      TabIndex        =   3
      Top             =   120
      Width           =   8895
   End
End
Attribute VB_Name = "frmSierpinskiTriangles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Terry Su
'Date: May 20, 2020
'Purpose: To generate Sierpinski’s Triangles using a simple 'for loop algorithm
Option Explicit

'GLOBAL VARIABLES
'declare global variables
Dim intCount As Integer
Dim intX1 As Integer
Dim intY1 As Integer
Dim intX2 As Integer
Dim intY2 As Integer
Dim intX3 As Integer
Dim intY3 As Integer
Dim intX4 As Integer
Dim intY4 As Integer

Private Sub cmdClear_Click()
picTouchPad.Cls
intCount = 0
cmdGo.Enabled = False
End Sub

Private Sub cmdGo_Click()
'Declare
Dim intRandom As Integer
Dim intCounter As Integer

'Initialize
intRandom = Int(Rnd * 3) + 1
intCounter = 0

'no input

'Process/Output
For intCounter = 1 To 5000
intRandom = Int(Rnd * 3) + 1
    If intRandom = 1 Then
        picTouchPad.PSet ((intX4 + intX1) / 2, (intY4 + intY1) / 2), RGB(200, 0, 300)
        intX4 = (intX4 + intX1) / 2
        intY4 = (intY4 + intY1) / 2
    ElseIf intRandom = 2 Then
        picTouchPad.PSet ((intX4 + intX2) / 2, (intY4 + intY2) / 2), RGB(100, 200, 300)
        intX4 = (intX4 + intX2) / 2
        intY4 = (intY4 + intY2) / 2
    ElseIf intRandom = 3 Then
        picTouchPad.PSet ((intX4 + intX3) / 2, (intY4 + intY3) / 2), RGB(20, 100, 150)
        intX4 = (intX4 + intX3) / 2
        intY4 = (intY4 + intY3) / 2
     End If
        
Next intCounter

End Sub

Private Sub Form_Activate()
cmdGo.Enabled = False
Randomize
'Initialize global variables
intCount = 0
intX1 = 0
intY1 = 0
intX2 = 0
intY2 = 0
intX3 = 0
intY3 = 0
intX4 = 0
intY4 = 0
End Sub

Private Sub picTouchPad_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

intCount = intCount + 1

If intCount = 1 Then
    picTouchPad.PSet (X, Y), vbRed
    picTouchPad.Circle (X, Y), 2, vbRed
    intX1 = X
    intY1 = Y
ElseIf intCount = 2 Then
    picTouchPad.PSet (X, Y), vbBlue
    picTouchPad.Circle (X, Y), 2, vbBlue
    intX2 = X
    intY2 = Y
ElseIf intCount = 3 Then
    picTouchPad.PSet (X, Y), vbGreen
    picTouchPad.Circle (X, Y), 2, vbGreen
    intX3 = X
    intY3 = Y
ElseIf intCount = 4 Then
    picTouchPad.PSet (X, Y), RGB(0, 100, 300)
    picTouchPad.Circle (X, Y), 2, RGB(0, 100, 300)
    intX4 = X
    intY4 = Y
    cmdGo.Enabled = True
Else
    MsgBox "You've clicked 4 times, please stop", vbExclamation, "Warning"
End If

End Sub
