VERSION 5.00
Begin VB.Form frmAcronymGenerator 
   Caption         =   "Acronym Generator"
   ClientHeight    =   2865
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   ScaleHeight     =   2865
   ScaleWidth      =   7425
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstOutput 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   7095
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "Process"
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
      TabIndex        =   2
      Top             =   1440
      Width           =   7095
   End
   Begin VB.TextBox txtInput 
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
      Left            =   120
      TabIndex        =   1
      Text            =   "(enter here)"
      Top             =   840
      Width           =   7095
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Please enter a phrase in the text box below"
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
      Width           =   7095
   End
End
Attribute VB_Name = "frmAcronymGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Terry Su
'Date: April 30, 2020
'Purpose: To generate an acronym out of a user inputed phrase
Option Explicit

Private Sub cmdProcess_Click()
'declare
Dim strInput As String
Dim intSpace As Integer
Dim strAcro As String
'initialize
strInput = ""
strAcro = ""
intSpace = 0
'intput
strInput = Trim(txtInput.Text)
'Process/output
lstOutput.Clear
Do
    strAcro = strAcro + UCase(Left(strInput, 1))
    intSpace = InStr(strInput, " ")
    strInput = Trim(Mid(strInput, intSpace + 1))
Loop While intSpace <> 0
lstOutput.AddItem "Your acronym is " & "'" & strAcro & "'"
End Sub

Private Sub txtInput_Click()
txtInput.Text = ""
End Sub
