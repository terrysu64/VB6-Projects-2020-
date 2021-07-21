VERSION 5.00
Begin VB.Form frmStructuredProgramming 
   Caption         =   "structured programming - sandbox"
   ClientHeight    =   3465
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   3465
   ScaleWidth      =   6585
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picOutput 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      ScaleHeight     =   915
      ScaleWidth      =   6195
      TabIndex        =   5
      Top             =   2280
      Width           =   6255
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "Process Your Input"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   6255
   End
   Begin VB.TextBox txtWord2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3720
      TabIndex        =   3
      Top             =   720
      Width           =   2655
   End
   Begin VB.TextBox txtWord1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3720
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label lblWord2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Enter Second Word"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   3375
   End
   Begin VB.Label lblWord1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Enter First Word:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
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
Attribute VB_Name = "frmStructuredProgramming"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Terry Su
'Date: April 14, 2020
'Purpose: example sandbox
Option Explicit
'======================================================

Private Sub cmdProcess_Click()
'Declare
Dim strWord1 As String
Dim strWord2 As String
'Intialize
strWord1 = ""
strWord2 = ""
'Input
strWord1 = Trim(txtWord1.Text)
strWord2 = Trim(txtWord2.Text)
'Calculations/Output
picOutput.Cls
If strWord1 = strWord2 Then
    picOutput.Print "The words are the same"
ElseIf Len(strWord1) > Len(strWord2) Then
    picOutput.Print "the first word is longer then the second"
Else
    picOutput.Print "the words are different"
End If

End Sub
