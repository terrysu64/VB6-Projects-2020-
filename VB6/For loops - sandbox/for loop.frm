VERSION 5.00
Begin VB.Form frmForLoops 
   Caption         =   "Sandbox - For Loops"
   ClientHeight    =   5715
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5565
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   5565
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picOutput 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   2760
      ScaleHeight     =   3915
      ScaleWidth      =   2595
      TabIndex        =   5
      Top             =   1560
      Width           =   2655
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear / Reset"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2760
      TabIndex        =   4
      Top             =   720
      Width           =   2655
   End
   Begin VB.CommandButton cmdCount 
      Caption         =   "Count"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   2535
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
      Height          =   3975
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   2535
   End
   Begin VB.TextBox txtInput 
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
      Left            =   3720
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblInput 
      Caption         =   "Enter Stopping Value:"
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
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmForLoops"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Terry Su
'Date: May 11, 2020
'Purpose: Sandbox - Experimenting with FOR Loops
Option Explicit

Private Sub cmdClear_Click()
lstOutput.Clear
picOutput.Cls
End Sub

Private Sub cmdCount_Click()
'declare
Dim intCount As Integer
Dim intResult As Integer
Dim intStop As Integer
Dim intRows As Integer
Dim intColumns As Integer
'initialize
intCount = 0
intResult = 0
intRows = 0
intColumns = 0
'input
intStop = Val(txtInput.Text)
'process/output
For intCount = 1 To intStop
    intResult = intCount * intCount
    lstOutput.AddItem intCount & " : " & intResult
Next intCount

For intRows = 1 To intStop
    For intColumns = 1 To intStop
        picOutput.Print "*";
    Next intColumns
    picOutput.Print ""
Next intRows

End Sub




