VERSION 5.00
Begin VB.Form frmRockPaperScissors 
   Caption         =   "Rock Paper Scissors"
   ClientHeight    =   5010
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11100
   LinkTopic       =   "Form1"
   ScaleHeight     =   5010
   ScaleWidth      =   11100
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picComputer 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10080
      ScaleHeight     =   435
      ScaleWidth      =   795
      TabIndex        =   11
      Top             =   1080
      Width           =   855
   End
   Begin VB.PictureBox picYou 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      ScaleHeight     =   435
      ScaleWidth      =   795
      TabIndex        =   10
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5880
      TabIndex        =   7
      Top             =   4200
      Width           =   5055
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "GO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   4200
      Width           =   5535
   End
   Begin VB.Frame fraRockPaperScissors 
      Caption         =   "Select your choice"
      Height          =   3015
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   3015
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Left            =   0
         TabIndex        =   5
         Top             =   3000
         Width           =   255
      End
      Begin VB.OptionButton optScissors 
         Caption         =   "Scissors"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   4
         Top             =   2040
         Width           =   2295
      End
      Begin VB.OptionButton optPaper 
         Caption         =   "Paper"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   2175
      End
      Begin VB.OptionButton optRock 
         Caption         =   "Rock"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Image imgYou 
      BorderStyle     =   1  'Fixed Single
      Height          =   2175
      Left            =   3360
      Picture         =   "Rock Paper Scissors.frx":0000
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   3375
   End
   Begin VB.Image imgComputer 
      BorderStyle     =   1  'Fixed Single
      Height          =   2175
      Left            =   7200
      Picture         =   "Rock Paper Scissors.frx":2A38
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   3375
   End
   Begin VB.Label lblComputer 
      Alignment       =   2  'Center
      Caption         =   "Computer Score"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      TabIndex        =   9
      Top             =   1080
      Width           =   3015
   End
   Begin VB.Label lblYou 
      Alignment       =   2  'Center
      Caption         =   "Your score"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   8
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Make your selection and click GO to begin playing!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
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
      Width           =   10815
   End
End
Attribute VB_Name = "frmRockPaperScissors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Terry Su
'Date: April 17, 2020
'Purpose: to simulate a rock, paper, scissors game.
Option Explicit
'GLOBAL VARIABLES
Dim intYouScore As Integer
Dim intComputerScore As Integer

Private Sub cmdGo_Click()
'declare
Dim ComputerRandom As Integer
'initialization
ComputerRandom = Int(Rnd * 3) + 1
'Process/Output
    'rock = 1
    'paper = 2
    'scissors = 3
        'display
picYou.Cls
picComputer.Cls
If ComputerRandom = 1 Then
    imgComputer.Picture = LoadPicture(App.Path & "\rock.jpg")
ElseIf ComputerRandom = 2 Then
    imgComputer.Picture = LoadPicture(App.Path & "\paper.jpg")
Else
    imgComputer.Picture = LoadPicture(App.Path & "\scissors.jpg")
End If

        'compare the two results
If optRock.Value = True And ComputerRandom = 1 Then
    lblMessage.Caption = "It's a tie! Try again!"
    picYou.Print intYouScore
    picComputer.Print intComputerScore
    
ElseIf optRock.Value = True And ComputerRandom = 2 Then
    lblMessage.Caption = "You lose! Try again!"
    intComputerScore = intComputerScore + 1
    picYou.Print intYouScore
    picComputer.Print intComputerScore
    
ElseIf optRock.Value = True And ComputerRandom = 3 Then
    lblMessage.Caption = "You win!"
    intYouScore = intYouScore + 1
    picYou.Print intYouScore
    picComputer.Print intComputerScore

ElseIf optPaper.Value = True And ComputerRandom = 2 Then
    lblMessage.Caption = "It's a tie! Try again!"
    picYou.Print intYouScore
    picComputer.Print intComputerScore
    
ElseIf optPaper.Value = True And ComputerRandom = 3 Then
    lblMessage.Caption = "You lose! Try again!"
    intComputerScore = intComputerScore + 1
    picYou.Print intYouScore
    picComputer.Print intComputerScore
    
ElseIf optPaper.Value = True And ComputerRandom = 1 Then
    lblMessage.Caption = "You win!"
    intYouScore = intYouScore + 1
    picYou.Print intYouScore
    picComputer.Print intComputerScore

ElseIf optScissors.Value = True And ComputerRandom = 3 Then
    lblMessage.Caption = "It's a tie! Try again!"
    picYou.Print intYouScore
    picComputer.Print intComputerScore
    
ElseIf optScissors.Value = True And ComputerRandom = 1 Then
    lblMessage.Caption = "You lose! Try again!"
    intComputerScore = intComputerScore + 1
    picYou.Print intYouScore
    picComputer.Print intComputerScore
    
ElseIf optScissors.Value = True And ComputerRandom = 2 Then
    lblMessage.Caption = "You win!"
    intYouScore = intYouScore + 1
    picYou.Print intYouScore
    picComputer.Print intComputerScore
End If


Option1.Value = True
cmdGo.Enabled = False

End Sub

Private Sub cmdReset_Click()
picYou.Cls
picComputer.Cls
intYouScore = 0
intComputerScore = 0
picYou.Print intYouScore
picComputer.Print intComputerScore
imgComputer.Picture = LoadPicture(App.Path & "\computer.jpg")
imgYou.Picture = LoadPicture(App.Path & "\you.jpg")
End Sub

Private Sub Form_Activate()
'initialization (global variables)
intYouScore = 0
intComputerScore = 0
'output
picYou.Print intYouScore
picComputer.Print intComputerScore
Option1.Value = True
cmdGo.Enabled = False
Randomize
End Sub

Private Sub optPaper_Click()
imgYou.Picture = LoadPicture(App.Path & "\paper.jpg")
cmdGo.Enabled = True
End Sub

Private Sub optRock_Click()
imgYou.Picture = LoadPicture(App.Path & "\rock.jpg")
cmdGo.Enabled = True
End Sub

Private Sub optScissors_Click()
imgYou.Picture = LoadPicture(App.Path & "\scissors.jpg")
cmdGo.Enabled = True
End Sub
