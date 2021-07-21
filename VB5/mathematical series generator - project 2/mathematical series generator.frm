VERSION 5.00
Begin VB.Form frmMathematicalSeriesGenerator 
   Caption         =   "Mathematical Series Generator"
   ClientHeight    =   6105
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10800
   LinkTopic       =   "Form1"
   ScaleHeight     =   6105
   ScaleWidth      =   10800
   StartUpPosition =   3  'Windows Default
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
      TabIndex        =   9
      Top             =   5160
      Width           =   10455
   End
   Begin VB.Frame fraFibonacci 
      Caption         =   "Fibonacci Sequence"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   4200
      Width           =   7095
      Begin VB.TextBox txtNumberTermsF 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   2520
         TabIndex        =   17
         Text            =   "(#of terms)"
         Top             =   120
         Width           =   2055
      End
   End
   Begin VB.Frame fraGeometric 
      Caption         =   "Geometric"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   3120
      Width           =   7095
      Begin VB.TextBox txtNumberTermsG 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4920
         TabIndex        =   16
         Text            =   "(# of terms)"
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtCommonRatioNumberG 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         TabIndex        =   15
         Text            =   "(common ratio#)"
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtStartingNumberG 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   14
         Text            =   "(starting #)"
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame fraArithmetic 
      Caption         =   "Arithmetic"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   7095
      Begin VB.TextBox txtNumberTermsA 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4920
         TabIndex        =   13
         Text            =   "(# of terms)"
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtCommonDifferenceNumberA 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         TabIndex        =   12
         Text            =   "(common difference#)"
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtStartingNumberA 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   11
         Text            =   "(starting #)"
         Top             =   240
         Width           =   1935
      End
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
      Height          =   3540
      Left            =   7440
      TabIndex        =   5
      Top             =   1200
      Width           =   3135
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
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   7095
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Left            =   7080
         TabIndex        =   10
         Top             =   720
         Width           =   15
      End
      Begin VB.OptionButton optF 
         Caption         =   "Fibonacci Sequence"
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
         Left            =   4560
         TabIndex        =   4
         Top             =   240
         Width           =   2415
      End
      Begin VB.OptionButton optG 
         Caption         =   "Geometric"
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
         Left            =   2400
         TabIndex        =   3
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton optA 
         Caption         =   "Arithmetic"
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
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Select A Mathematical Series To Generate"
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
      Width           =   10455
   End
End
Attribute VB_Name = "frmMathematicalSeriesGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Terry Su
'Date: April 29, 2020
'Purpose: To allow a user to generate a variety of customizable mathematical sequences with starting and ending values.
Option Explicit

Private Sub cmdCalculate_Click()

If Option1.Value = True Then
    MsgBox "Select an option!", vbExclamation, "Warning"
End If

If optA.Value = True And (txtStartingNumberA.Text = "(starting #)" Or txtCommonDifferenceNumberA.Text = "(common difference #)" Or txtNumberTermsA.Text = "(# of terms)") Then
    MsgBox "Make sure all required information is filled!", vbExclamation, "Warning"
End If

If optG.Value = True And (txtStartingNumberG.Text = "(starting #)" Or txtCommonRatioNumberG.Text = "(common ratio #)" Or txtNumberTermsG.Text = "(# of terms)") Then
    MsgBox "Make sure all required information is filled!", vbExclamation, "Warning"
End If

If optF.Value = True And txtNumberTermsF.Text = "(# of terms)" Then
    MsgBox "Make sure all required information is filled!", vbExclamation, "Warning"
End If

lstOutput.Clear
If optA.Value = True Then
    'Declare
    Dim sglStartingNumberA As Single
    Dim sglCommonDifferenceNumberA As Single
    Dim intNumberTermsA As Integer
    Dim intCountA As Integer
    'Initialize
    sglStartingNumberA = 0
    sglCommonDifferenceNumberA = 0
    intNumberTermsA = 0
    intCountA = 0
    'Input
    sglStartingNumberA = Val(txtStartingNumberA.Text)
    sglCommonDifferenceNumberA = Val(txtCommonDifferenceNumberA.Text)
    intNumberTermsA = Val(txtNumberTermsA.Text)
    'Calculations/Output
     lstOutput.AddItem sglStartingNumberA & ","
    Do While intCountA + 1 < intNumberTermsA
        intCountA = intCountA + 1
        lstOutput.AddItem sglStartingNumberA + (sglCommonDifferenceNumberA * intCountA) & ","
    Loop

ElseIf optG.Value = True Then
    'Declare
    Dim sglStartingNumberG As Single
    Dim sglCommonRatioNumberG As Single
    Dim intNumberTermsG As Integer
    Dim intCountG As Integer
    'Initialize
    sglStartingNumberG = 0
    sglCommonRatioNumberG = 0
    intNumberTermsG = 0
    intCountG = 0
    'Input
    sglStartingNumberG = Val(txtStartingNumberG.Text)
    sglCommonRatioNumberG = Val(txtCommonRatioNumberG.Text)
    intNumberTermsG = Val(txtNumberTermsG.Text)
    'Calculations/Output
     lstOutput.AddItem sglStartingNumberG & ","
    Do While intCountG + 1 < intNumberTermsG
        intCountG = intCountG + 1
        lstOutput.AddItem sglStartingNumberG * (sglCommonRatioNumberG ^ intCountG) & ","
    Loop

ElseIf optF.Value = True Then
    'Declare
     Dim intNumberTermsF As Integer
     Dim intCountF As Integer
     Dim intF As Integer
     Dim Sqrt As Double
     'initialize
     intNumberTermsF = 0
     intCountF = 0
     intF = 1
     Sqrt = Math.Sqr(5)
     'input
     intNumberTermsF = Val(txtNumberTermsF.Text)
     'Process/Output
     lstOutput.AddItem intF & ","
     Do While intCountF + 1 < intNumberTermsF
        intCountF = intCountF + 1
        intF = (1 / Sqrt) * ((1 + Sqrt) / 2) ^ (intCountF + 1) - ((1 / Sqrt) * ((1 - Sqrt) / 2) ^ (intCountF + 1))
        lstOutput.AddItem intF & ","
     Loop
End If

End Sub

Private Sub Form_Load()
Option1.Value = True
fraArithmetic.Enabled = False
fraGeometric.Enabled = False
fraFibonacci.Enabled = False
End Sub


Private Sub optA_Click()
fraArithmetic.Enabled = True
fraGeometric.Enabled = False
fraFibonacci.Enabled = False
txtStartingNumberA.Text = "(starting #)"
txtCommonDifferenceNumberA.Text = "(common difference #)"
txtNumberTermsA.Text = "(# of terms)"
txtStartingNumberG.Text = "(starting #)"
txtCommonRatioNumberG.Text = "(common ratio #)"
txtNumberTermsG.Text = "(# of terms)"
txtNumberTermsF.Text = "(# of terms)"
End Sub

Private Sub optF_Click()
fraFibonacci.Enabled = True
fraArithmetic.Enabled = False
fraGeometric.Enabled = False
txtStartingNumberA.Text = "(starting #)"
txtCommonDifferenceNumberA.Text = "(common difference #)"
txtNumberTermsA.Text = "(# of terms)"
txtStartingNumberG.Text = "(starting #)"
txtCommonRatioNumberG.Text = "(common ratio #)"
txtNumberTermsG.Text = "(# of terms)"
txtNumberTermsF.Text = "(# of terms)"
End Sub

Private Sub optG_Click()
fraGeometric.Enabled = True
fraArithmetic.Enabled = False
fraFibonacci.Enabled = False
txtStartingNumberA.Text = "(starting #)"
txtCommonDifferenceNumberA.Text = "(common difference #)"
txtNumberTermsA.Text = "(# of terms)"
txtStartingNumberG.Text = "(starting #)"
txtCommonRatioNumberG.Text = "(common ratio #)"
txtNumberTermsG.Text = "(# of terms)"
txtNumberTermsF.Text = "(# of terms)"
End Sub

Private Sub txtCommonDifferenceNumberA_Click()
txtCommonDifferenceNumberA.Text = ""
End Sub

Private Sub txtCommonRatioNumberG_Click()
txtCommonRatioNumberG.Text = ""
End Sub

Private Sub txtNumberTermsA_Click()
txtNumberTermsA.Text = ""
End Sub

Private Sub txtNumberTermsF_Click()
txtNumberTermsF.Text = ""
End Sub

Private Sub txtNumberTermsG_Click()
txtNumberTermsG.Text = ""
End Sub

Private Sub txtStartingNumberA_Click()
txtStartingNumberA.Text = ""
End Sub

Private Sub txtStartingNumberG_Click()
txtStartingNumberG.Text = ""
End Sub
