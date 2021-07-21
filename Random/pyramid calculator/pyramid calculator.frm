VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4230
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   ScaleHeight     =   4230
   ScaleWidth      =   10695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPyramidCalc 
      Caption         =   "calculate"
      Height          =   735
      Left            =   6240
      TabIndex        =   4
      Top             =   1920
      Width           =   2895
   End
   Begin VB.PictureBox picOutput 
      Height          =   1095
      Left            =   240
      ScaleHeight     =   1035
      ScaleWidth      =   5115
      TabIndex        =   3
      Top             =   1920
      Width           =   5175
   End
   Begin VB.TextBox txtS 
      Height          =   405
      Left            =   4440
      TabIndex        =   2
      Text            =   "s"
      Top             =   600
      Width           =   1935
   End
   Begin VB.TextBox txtH 
      Height          =   525
      Left            =   2880
      TabIndex        =   1
      Text            =   "H"
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox txtB 
      Height          =   405
      Left            =   480
      TabIndex        =   0
      Text            =   "b"
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPyramidCalc_Click()
'delcare
Dim dblB As Double
Dim dblH As Double
Dim dblS As Double
Dim dblVolume As Double
Dim dblSurfaceArea As Double
'Initialize
dblB = 0
dblH = 0
dblS = 0
dblVolume = 0
dblSurfaceArea = 0
'Input
dblB = Val(txtB.Text)
dblH = Val(txtH.Text)
dblS = Val(txtS.Text)
'Calculations
dblVolume = (1 / 3) * (dblB ^ 2) * dblH
dblSurfaceArea = (2 * dblB * dblS) + (dblB ^ 2)
'Output
picOutput.Cls
picOutput.Print "The volume of the square based pyramid is " & dblVolume & " cube units"
picOutput.Print "The surface area of the square based pyramid is " & dblSurfaceArea & " square units"
End Sub

