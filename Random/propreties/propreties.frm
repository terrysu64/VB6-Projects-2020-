VERSION 5.00
Begin VB.Form frmPropreties 
   Caption         =   "Propreties"
   ClientHeight    =   4635
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   ScaleHeight     =   4635
   ScaleWidth      =   4800
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBottom 
      Caption         =   "Bottom"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2040
      TabIndex        =   8
      Top             =   1920
      Width           =   2535
   End
   Begin VB.CommandButton cmdTop 
      Caption         =   "Top"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2040
      TabIndex        =   7
      Top             =   840
      Width           =   2535
   End
   Begin VB.Frame fraFonts 
      Caption         =   "Fonts"
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Width           =   4455
      Begin VB.OptionButton OptArial 
         Caption         =   "Arial"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   6
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton optMSSansSerif 
         Caption         =   "MS Sans Serif"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame fraEnabling 
      Caption         =   "Enabling"
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1695
      Begin VB.OptionButton optDisable 
         Caption         =   "Disable"
         Height          =   615
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   1455
      End
      Begin VB.OptionButton optEnable 
         Caption         =   "Enable"
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Label lblPropreties 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Propreties"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
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
      Width           =   4455
   End
End
Attribute VB_Name = "frmPropreties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Terry Su
'Date: march 26, 2020
'Purpose: Unit 2 test - part B
Option Explicit

Private Sub cmdBottom_Click()
lblPropreties.Top = 3960
End Sub

Private Sub cmdTop_Click()
lblPropreties.Top = 120
End Sub

Private Sub Form_Load()
optDisable.Value = True
optMSSansSerif.Value = True
End Sub

Private Sub OptArial_Click()
lblPropreties.Font = "Arial"
End Sub

Private Sub optDisable_Click()
cmdTop.Enabled = False
cmdBottom.Enabled = False
End Sub

Private Sub optEnable_Click()
cmdTop.Enabled = True
cmdBottom.Enabled = True
End Sub

Private Sub optMSSansSerif_Click()
lblPropreties.Font = "MS Sans Serif"
End Sub
