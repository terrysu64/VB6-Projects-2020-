VERSION 5.00
Begin VB.Form frmVowelCounter 
   Caption         =   "Vowel Counter"
   ClientHeight    =   4530
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9180
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   9180
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdProcess 
      Caption         =   "Process"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   7200
      TabIndex        =   4
      Top             =   1560
      Width           =   1815
   End
   Begin VB.ListBox lstOutput 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2685
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   6855
   End
   Begin VB.CheckBox chkInclude 
      Caption         =   "include ""y/Y"" as vowel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox txtInput 
      Alignment       =   2  'Center
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
      TabIndex        =   1
      Text            =   "(Enter here)"
      Top             =   840
      Width           =   6975
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Please enter a word or a phrase in the text box below"
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
      Width           =   8895
   End
End
Attribute VB_Name = "frmVowelCounter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Terry Su
'Date: April 30, 2020
'Purpose: To determine the # of vowels, frequency of each vowel and the # of consonants in a user inputed string.
Option Explicit

Private Sub cmdProcess_Click()

'Declare
Dim strInput As String
Dim intVowel As Integer
Dim intConsonant As Integer
Dim intCount As Integer
Dim strLetter As String
Dim intA As Integer
Dim intE As Integer
Dim intI As Integer
Dim intO As Integer
Dim intU As Integer
Dim intY As Integer

'Initialize
strInput = ""
intVowel = 0
intConsonant = 0
intCount = 0
strLetter = ""
intA = 0
intE = 0
intI = 0
intO = 0
intU = 0

'Input
strInput = LCase(Trim(txtInput.Text))

'process
'Find the numbers of each vowel
Do While InStr(strInput, "a") <> 0
    intA = intA + 1
    intVowel = intVowel + 1
    strInput = Mid(strInput, InStr(strInput, "a") + 1)
Loop
strInput = LCase(Trim(txtInput.Text))

Do While InStr(strInput, "e") <> 0
    intE = intE + 1
    intVowel = intVowel + 1
    strInput = Mid(strInput, InStr(strInput, "e") + 1)
Loop
strInput = LCase(Trim(txtInput.Text))

Do While InStr(strInput, "i") <> 0
    intI = intI + 1
    intVowel = intVowel + 1
    strInput = Mid(strInput, InStr(strInput, "i") + 1)
Loop
strInput = LCase(Trim(txtInput.Text))

Do While InStr(strInput, "o") <> 0
    intO = intO + 1
    intVowel = intVowel + 1
    strInput = Mid(strInput, InStr(strInput, "o") + 1)
Loop
strInput = LCase(Trim(txtInput.Text))

Do While InStr(strInput, "u") <> 0
    intU = intU + 1
    intVowel = intVowel + 1
    strInput = Mid(strInput, InStr(strInput, "u") + 1)
Loop
strInput = LCase(Trim(txtInput.Text))

Do While InStr(strInput, "y") <> 0
    intY = intY + 1
    strInput = Mid(strInput, InStr(strInput, "y") + 1)
Loop
strInput = LCase(Trim(txtInput.Text))

'find the numbers of each consonant
If chkInclude.Value = 0 Then
    Do While intCount < Len(strInput)
        intCount = intCount + 1
        If InStr("bcdfghjklmnpqrstvwxyz", Mid(strInput, intCount, 1)) <> 0 Then
            intConsonant = intConsonant + 1
        End If
    Loop
    
Else
    'not including y
    Do While intCount < Len(strInput)
        intCount = intCount + 1
        If InStr("bcdfghjklmnpqrstvwxz", Mid(strInput, intCount, 1)) <> 0 Then
            intConsonant = intConsonant + 1
        End If
    Loop
End If

    

'Output
If chkInclude.Value = 0 Then
    lstOutput.Clear
    lstOutput.AddItem "There are " & intConsonant & " consonant(s)"
    lstOutput.AddItem "There are " & intVowel & " vowel(s)"
    lstOutput.AddItem "There are " & intA & " A/a(s)"
    lstOutput.AddItem "There are " & intE & " E/e(s)"
    lstOutput.AddItem "There are " & intI & " I/i(s)"
    lstOutput.AddItem "There are " & intO & " O/o(s)"
    lstOutput.AddItem "There are " & intU & " U/u(s)"

Else
    lstOutput.Clear
    lstOutput.AddItem "There are " & intConsonant & " consonant(s)"
    lstOutput.AddItem "There are " & intVowel + intY & " vowel(s)"
    lstOutput.AddItem "There are " & intA & " A/a(s)"
    lstOutput.AddItem "There are " & intE & " E/e(s)"
    lstOutput.AddItem "There are " & intI & " I/i(s)"
    lstOutput.AddItem "There are " & intO & " O/o(s)"
    lstOutput.AddItem "There are " & intU & " U/u(s)"
    lstOutput.AddItem "There are " & intY & " Y/y(s)"

End If

End Sub

Private Sub txtInput_click()
txtInput.Text = ""
End Sub
