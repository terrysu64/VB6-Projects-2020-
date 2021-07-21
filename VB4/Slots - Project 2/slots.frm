VERSION 5.00
Begin VB.Form frmSlotsGame 
   Caption         =   "Slots Game!"
   ClientHeight    =   5730
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8235
   LinkTopic       =   "Form1"
   ScaleHeight     =   5730
   ScaleWidth      =   8235
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBorrowMoney 
      Caption         =   "Borrow Money"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   6
      Top             =   4560
      Width           =   4815
   End
   Begin VB.PictureBox picTotalOwed 
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
      Left            =   2760
      ScaleHeight     =   555
      ScaleWidth      =   2115
      TabIndex        =   5
      Top             =   3720
      Width           =   2175
   End
   Begin VB.PictureBox picPurse 
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
      ScaleHeight     =   555
      ScaleWidth      =   2235
      TabIndex        =   4
      Top             =   3720
      Width           =   2295
   End
   Begin VB.CommandButton cmdPull 
      Caption         =   "Pull"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   5400
      TabIndex        =   1
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label lblTotalOwed 
      Alignment       =   2  'Center
      Caption         =   "Total Owed"
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
      Left            =   2760
      TabIndex        =   3
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label lblPurse 
      Alignment       =   2  'Center
      Caption         =   "Your Purse"
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
      Left            =   120
      TabIndex        =   2
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Image imgFruit3 
      BorderStyle     =   1  'Fixed Single
      Height          =   1815
      Left            =   5400
      Picture         =   "slots.frx":0000
      Stretch         =   -1  'True
      Top             =   960
      Width           =   2295
   End
   Begin VB.Image imgFruit2 
      BorderStyle     =   1  'Fixed Single
      Height          =   1815
      Left            =   2760
      Picture         =   "slots.frx":6998
      Stretch         =   -1  'True
      Top             =   960
      Width           =   2295
   End
   Begin VB.Image imgFruit1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1815
      Left            =   120
      Picture         =   "slots.frx":D330
      Stretch         =   -1  'True
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Welcome to Slots! Pull to begin your game!"
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
      Width           =   7695
   End
End
Attribute VB_Name = "frmSlotsGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Terry Su
'Date: April 16, 2020
'Purpose: To create a slots casino like game, while implementing the concepts of decisions, relational and logical operators.
Option Explicit
'GLOBAL VARIABLES
Dim sglPurse As Single
Dim sglTotalOwed As Single

Private Sub cmdBorrowMoney_Click()
picPurse.Cls
picTotalOwed.Cls
sglPurse = sglPurse + 20
sglTotalOwed = sglTotalOwed + 20
picPurse.Print FormatCurrency(sglPurse)
picTotalOwed.Print FormatCurrency(sglTotalOwed)
lblMessage.Caption = "Don't borrow too much!"
cmdBorrowMoney.Enabled = False
cmdPull.Enabled = True
End Sub

Private Sub cmdPull_Click()
'Declare
Dim intrandom1 As Integer
Dim intrandom2 As Integer
Dim intrandom3 As Integer
'Initialize
intrandom1 = Int(Rnd * 3) + 1
intrandom2 = Int(Rnd * 3) + 1
intrandom3 = Int(Rnd * 3) + 1
'Process/Output
    'cherry  = 1
    'lemon = 2
    'apple = 3
        'image display
If intrandom1 = 1 Then
    imgFruit1.Picture = LoadPicture(App.Path & "\cherry.jpg")
ElseIf intrandom1 = 2 Then
    imgFruit1.Picture = LoadPicture(App.Path & "\lemon.jpg")
Else
    imgFruit1.Picture = LoadPicture(App.Path & "\apple.jpg")
End If

If intrandom2 = 1 Then
    imgFruit2.Picture = LoadPicture(App.Path & "\cherry.jpg")
ElseIf intrandom2 = 2 Then
    imgFruit2.Picture = LoadPicture(App.Path & "\lemon.jpg")
Else
    imgFruit2.Picture = LoadPicture(App.Path & "\apple.jpg")
End If

If intrandom3 = 1 Then
    imgFruit3.Picture = LoadPicture(App.Path & "\cherry.jpg")
ElseIf intrandom3 = 2 Then
    imgFruit3.Picture = LoadPicture(App.Path & "\lemon.jpg")
Else
    imgFruit3.Picture = LoadPicture(App.Path & "\apple.jpg")
End If
        'Money and label display
picPurse.Cls
If intrandom1 = 1 And intrandom2 = 1 And intrandom3 = 1 Then
    sglPurse = sglPurse + 4
    picPurse.Print FormatCurrency(sglPurse)
    lblMessage.Caption = "you win $4.00!"
    
ElseIf intrandom1 = 2 And intrandom2 = 2 And intrandom3 = 2 Then
    sglPurse = sglPurse + 8
    picPurse.Print FormatCurrency(sglPurse)
    lblMessage.Caption = "you win $8.00!"
    
ElseIf intrandom1 = 3 And intrandom2 = 3 And intrandom3 = 3 Then
    sglPurse = sglPurse + 12
    picPurse.Print FormatCurrency(sglPurse)
    lblMessage.Caption = "you win $12.00!"
    
Else
    sglPurse = sglPurse - 1
    picPurse.Print FormatCurrency(sglPurse)
    lblMessage.Caption = "Try again"

End If
            'Borrowing money
If sglPurse = 0 Then
    cmdPull.Enabled = False
    cmdBorrowMoney.Enabled = True
    lblMessage.Caption = "You ran out of money!"
End If

End Sub

Private Sub Form_Activate()
'initialization (global variables)
sglPurse = 20
sglTotalOwed = 0
'Output
picPurse.Print FormatCurrency(sglPurse)
picTotalOwed.Print FormatCurrency(sglTotalOwed)
cmdBorrowMoney.Enabled = False
Randomize

End Sub


