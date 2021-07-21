VERSION 5.00
Begin VB.Form frmTemperatureConversion 
   Caption         =   "Temperature Conversion"
   ClientHeight    =   5850
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7620
   LinkTopic       =   "Form1"
   ScaleHeight     =   5850
   ScaleWidth      =   7620
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstOutput 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2670
      Left            =   3000
      TabIndex        =   8
      Top             =   1920
      Width           =   4335
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   4920
      Width           =   7335
   End
   Begin VB.TextBox txtEndingValue 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      TabIndex        =   6
      Text            =   "(Ending Value)"
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox txtStartingValue 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   5
      Text            =   "(Starting Value)"
      Top             =   960
      Width           =   2055
   End
   Begin VB.Frame fraOptions 
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   2655
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Left            =   0
         TabIndex        =   9
         Top             =   3720
         Width           =   135
      End
      Begin VB.OptionButton optKF 
         Caption         =   "Kelvin to Fahrenheit"
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
         Left            =   240
         TabIndex        =   4
         Top             =   2640
         Width           =   2175
      End
      Begin VB.OptionButton optCK 
         Caption         =   "Celsius to Kelvin"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   240
         TabIndex        =   3
         Top             =   1560
         Width           =   2295
      End
      Begin VB.OptionButton optCF 
         Caption         =   "Celsius to Fahrenheit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Temperature Conversion"
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
      TabIndex        =   0
      Top             =   120
      Width           =   7335
   End
End
Attribute VB_Name = "frmTemperatureConversion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Terry Su
'Date: April 28, 2020
'Purpose: To allow a user to convert between different temperature measurement units.
Option Explicit

Private Sub cmdCalculate_Click()
'Declare
Dim sglStarting As Single
Dim sglEnding As Single
'Initialize
sglStarting = 0
sglEnding = 0
'Input
sglStarting = Val(txtStartingValue.Text)
sglEnding = Val(txtEndingValue.Text)
'Process/Output
If Option1.Value = True Then
    MsgBox "select a conversion option!", vbExclamation, "Warning"
End If

If sglStarting > sglEnding Then
    MsgBox "Starting value must not exceed the ending value!", vbExclamation, "Warning"
End If

If txtStartingValue.Text = "(Starting Value)" Or txtStartingValue.Text = "" Then
    MsgBox "Enter starting value!", vbExclamation, "Warning"
End If

If txtEndingValue.Text = "(Ending Value)" Or txtEndingValue.Text = "" Then
    MsgBox "Enter ending value!", vbExclamation, "Warning"
End If

lstOutput.Clear
If optCF.Value = True Then
    lstOutput.AddItem "Celsius  " & "Farenheit"
    Do While sglStarting <= sglEnding
       lstOutput.AddItem "  " & Format(sglStarting, ".0") & "°C =      " & (sglStarting * 9 / 5) + 32 & "°F"
       sglStarting = sglStarting + 0.5
    Loop
    
ElseIf optCK.Value = True Then
    lstOutput.AddItem "Celsius      " & "Kelvin"
    Do While sglStarting <= sglEnding
       lstOutput.AddItem "  " & Format(sglStarting, ".0") & "°C =     " & sglStarting + 273.15 & "K"
       sglStarting = sglStarting + 0.5
    Loop
    
ElseIf optKF.Value = True Then
   lstOutput.AddItem "Farhenheit    " & "Kelvin"
   Do While sglStarting <= sglEnding
       lstOutput.AddItem "   " & Format(sglStarting, ".0") & "K =        " & (sglStarting * 1.8) - 459.67 & "°F"
       sglStarting = sglStarting + 0.5
    Loop
End If
End Sub

Private Sub Form_Load()
Option1.Value = True
End Sub

Private Sub txtEndingValue_Click()
txtEndingValue.Text = ""
End Sub

Private Sub txtStartingValue_Click()
txtStartingValue.Text = ""
End Sub
