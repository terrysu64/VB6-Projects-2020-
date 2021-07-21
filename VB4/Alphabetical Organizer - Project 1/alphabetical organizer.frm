VERSION 5.00
Begin VB.Form frmAlphabeticalOrganizer 
   Caption         =   "Alphabetical Organizer"
   ClientHeight    =   3090
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   6450
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
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   6075
      TabIndex        =   5
      Top             =   2400
      Width           =   6135
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "Sort"
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
      Top             =   1560
      Width           =   6135
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
      Height          =   495
      Left            =   3360
      TabIndex        =   3
      Top             =   840
      Width           =   2895
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
      Height          =   495
      Left            =   3360
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label lblWord2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Enter Second Word:"
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
      Top             =   840
      Width           =   3015
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
      Width           =   3015
   End
End
Attribute VB_Name = "frmAlphabeticalOrganizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Terry Su
'Date: April 14, 2020
'Purpose: To organize 2 inputed strings in alphabetical order
Option Explicit
'=======================================================================================================

Private Sub cmdProcess_Click()
'Declare
Dim strWord1 As String
Dim strWord2 As String
Dim strWord3 As String
'Initialize
strWord1 = ""
strWord2 = ""
strWord3 = ""
'Input
strWord1 = LCase(Trim(txtWord1.Text))
strWord2 = LCase(Trim(txtWord2.Text))
'Process/Output
picOutput.Cls
If strWord1 > strWord2 Then
    picOutput.Print strWord2 & "," & strWord1
Else
    picOutput.Print strWord1 & "," & strWord2
End If
End Sub

Private Sub txtWord1_Click()
txtWord1.Text = ""
End Sub

Private Sub txtWord2_Click()
txtWord2.Text = ""
End Sub
