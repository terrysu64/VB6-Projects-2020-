VERSION 5.00
Begin VB.Form frmLoopsSandbox 
   Caption         =   "Loops Sandbox"
   ClientHeight    =   5100
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   ScaleHeight     =   5100
   ScaleWidth      =   6600
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear The List Box"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   4200
      Width           =   6255
   End
   Begin VB.ListBox lstOutput 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2670
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   6255
   End
   Begin VB.CommandButton cmdCount 
      Caption         =   "Click Here To Count"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "frmLoopsSandbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Terry Su
'Date: april 28, 2020
'Purpose: Loops Sandbox
Option Explicit
'===============================================

Private Sub cmdClear_Click()
lstOutput.Clear
End Sub

Private Sub cmdCount_Click()
'Declare
Dim intCount As Integer
Dim intSquare As Integer
'Intialize
intCount = 0
intSquare = 0
'Input
    'no input in this sandbox
'Process/Output
Do While intCount < 30
    'inside the loop must be indented
    intCount = intCount + 1
    intSquare = intCount * intCount
    lstOutput.AddItem intCount & " " & intSquare & "squared"
Loop

End Sub
