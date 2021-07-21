VERSION 5.00
Begin VB.Form frmDad 
   Caption         =   "Continuous disector"
   ClientHeight    =   4125
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   ScaleHeight     =   4125
   ScaleWidth      =   7410
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picOutput 
      Height          =   1215
      Left            =   240
      ScaleHeight     =   1155
      ScaleWidth      =   3195
      TabIndex        =   2
      Top             =   2280
      Width           =   3255
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "disect"
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   3375
   End
   Begin VB.TextBox txtInput 
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmDad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Terry Su
'Date: March 27, 2020
'Purpose: to continuously disect a string
Option Explicit

Private Sub cmdCalculate_Click()
'Declare
Dim strInput As String
Dim intComma As Integer
Dim strPart2 As String
'Input
strInput = Trim(txtInput.Text)
'Calculations
intComma = InStr(strInput, ",")
strPart2 = Mid(strInput, intComma + 1)
'Output
picOutput.Print (Mid(strInput, 1, intComma - 1))
picOutput.Print Mid(strPart2, 1, (InStr(strPart2, ",")) - 1)

End Sub


