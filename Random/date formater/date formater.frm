VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4725
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   6555
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picOut 
      Height          =   1815
      Left            =   480
      ScaleHeight     =   1755
      ScaleWidth      =   5475
      TabIndex        =   2
      Top             =   2400
      Width           =   5535
   End
   Begin VB.CommandButton cmdFormat 
      Caption         =   "format"
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   1560
      Width           =   2655
   End
   Begin VB.TextBox txtDate 
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Text            =   "month/day/year"
      Top             =   600
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdFormat_Click()
'Declare
Dim strInput As String
Dim strMonth As String
Dim strDay As String
Dim strYear As String
Dim intComma1 As Integer
Dim strPart2 As String
Dim intComma2 As Integer
'Initialization
strInput = ""
strMonth = ""
strDay = ""
strYear = ""
intComma1 = 0
strPart2 = ""
intComma2 = 0
'input
strInput = Trim(txtDate.Text)
'Calculations
intComma1 = InStr(strInput, "/")
strMonth = Left(strInput, intComma1 - 1)
strPart2 = Mid(strInput, intComma1 + 1)
intComma2 = InStr(strPart2, "/")
strDay = Left(strPart2, intComma2 - 1)
strYear = Mid(strPart2, intComma2 + 1)
'Output
picOut.Cls
picOut.Print strYear & "," & strMonth & "," & strDay
End Sub
