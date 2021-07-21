VERSION 5.00
Begin VB.Form frmForLoopsProject1 
   Caption         =   "For Loops Project 1"
   ClientHeight    =   4800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7725
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   7725
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstOutput2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3060
      Left            =   3840
      TabIndex        =   4
      Top             =   1560
      Width           =   3735
   End
   Begin VB.ListBox lstOutput 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3060
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   3615
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   2
      Top             =   840
      Width           =   3735
   End
   Begin VB.TextBox txtInput 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Text            =   "(enter here)"
      Top             =   840
      Width           =   3615
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Please enter a number greater than 0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
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
      Width           =   7455
   End
End
Attribute VB_Name = "frmForLoopsProject1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Terry Su
'Date: May 12, 2020
'Purpose: to generate a series of numbers using for loops based off a user input
Option Explicit

Private Sub cmdGo_Click()
'Declare
Dim int1 As Integer
Dim int2 As Integer
Dim intSum As Integer
Dim intCount As Integer
Dim intStep As Integer
Dim intCount2 As Integer
Dim intCount3 As Integer
Dim intCount4 As Integer
Dim intCount5 As Integer
Dim intCount6 As Integer
Dim intCount7 As Integer
Dim intCount8 As Integer
Dim intCount9 As Integer
Dim intCount10 As Integer
Dim intCount11 As Integer
Dim intCount12 As Integer
Dim intTotal As Integer
'Initialize
int1 = 0
int2 = 0
intSum = 0
intCount = 0
intCount2 = 0
intCount3 = 0
intCount4 = 0
intCount5 = 0
intCount6 = 0
intCount7 = 0
intCount8 = 0
intCount9 = 0
intCount10 = 0
intCount11 = 0
intCount12 = 0
intTotal = 0
'Input
intStep = Val(txtInput.Text)
'Process/Output
If intStep <= 0 Then
    MsgBox "entered number must be greater than 0", vbCritical, "Warning"
    End
End If

lstOutput.Clear
lstOutput2.Clear
For intCount = 1 To intStep
    int1 = Int(Rnd * 6) + 1
    int2 = Int(Rnd * 6) + 1
    intSum = int1 + int2
    If intSum = 2 Then
        intCount2 = intCount2 + 1
    ElseIf intSum = 3 Then
        intCount3 = intCount3 + 1
    ElseIf intSum = 4 Then
        intCount4 = intCount4 + 1
    ElseIf intSum = 5 Then
        intCount5 = intCount5 + 1
    ElseIf intSum = 6 Then
        intCount6 = intCount6 + 1
    ElseIf intSum = 7 Then
        intCount7 = intCount7 + 1
    ElseIf intSum = 8 Then
        intCount8 = intCount8 + 1
    ElseIf intSum = 9 Then
        intCount9 = intCount9 + 1
    ElseIf intSum = 10 Then
        intCount10 = intCount10 + 1
    ElseIf intSum = 11 Then
        intCount11 = intCount11 + 1
    ElseIf intSum = 12 Then
        intCount12 = intCount12 + 1
    End If
    lstOutput.AddItem "#1 = " & int1 & "," & " #2 = " & int2 & "," & " sum = " & intSum
Next intCount

intTotal = intCount2 + intCount3 + intCount4 + intCount5 + intCount6 + intCount7 + intCount8 + intCount9 + intCount10 + intCount11 + intCount12
lstOutput2.AddItem "count sum = 2: " & intCount2 & " (" & Format(intCount2 / intTotal * 100, "0") & "%)"
lstOutput2.AddItem "count sum = 3: " & intCount3 & " (" & Format(intCount3 / intTotal * 100, "0") & "%)"
lstOutput2.AddItem "count sum = 4: " & intCount4 & " (" & Format(intCount4 / intTotal * 100, "0") & "%)"
lstOutput2.AddItem "count sum = 5: " & intCount5 & " (" & Format(intCount5 / intTotal * 100, "0") & "%)"
lstOutput2.AddItem "count sum = 6: " & intCount6 & " (" & Format(intCount6 / intTotal * 100, "0") & "%)"
lstOutput2.AddItem "count sum = 7: " & intCount7 & " (" & Format(intCount7 / intTotal * 100, "0") & "%)"
lstOutput2.AddItem "count sum = 8: " & intCount8 & " (" & Format(intCount8 / intTotal * 100, "0") & "%)"
lstOutput2.AddItem "count sum = 9: " & intCount9 & " (" & Format(intCount9 / intTotal * 100, "0") & "%)"
lstOutput2.AddItem "count sum = 10: " & intCount10 & " (" & Format(intCount10 / intTotal * 100, "0") & "%)"
lstOutput2.AddItem "count sum = 11: " & intCount11 & " (" & Format(intCount11 / intTotal * 100, "0") & "%)"
lstOutput2.AddItem "count sum = 12: " & intCount12 & " (" & Format(intCount12 / intTotal * 100, "0") & "%)"
End Sub

Private Sub Form_Activate()
Randomize
End Sub

Private Sub txtInput_Click()
txtInput.Text = ""
End Sub
