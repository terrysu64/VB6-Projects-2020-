VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4470
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8130
   LinkTopic       =   "Form1"
   ScaleHeight     =   4470
   ScaleWidth      =   8130
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picTotalCost 
      Height          =   495
      Left            =   2280
      ScaleHeight     =   435
      ScaleWidth      =   2235
      TabIndex        =   5
      Top             =   2520
      Width           =   2295
   End
   Begin VB.PictureBox picChange 
      Height          =   615
      Left            =   5040
      ScaleHeight     =   555
      ScaleWidth      =   2835
      TabIndex        =   4
      Top             =   2520
      Width           =   2895
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "calculate"
      Height          =   855
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   5775
   End
   Begin VB.TextBox txtTicketCost 
      Height          =   615
      Left            =   3120
      TabIndex        =   1
      Text            =   "ticket cost"
      Top             =   480
      Width           =   2895
   End
   Begin VB.TextBox txtWallet 
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Text            =   "wallet"
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label lblTickets 
      Caption         =   "# tickets"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   2520
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCalculate_Click()
'Declare
Dim intWallet As Integer
Dim intTicketCost As Integer
Dim intTotalTickets As Integer
Dim intChange As Integer
Dim intTotalCost As Integer
'Initialization
intWallet = 0
intTicketCost = 0
intTotalTickets = 0
intChange = 0
intTotalCost = 0
'input
intWallet = Val(txtWallet.Text)
intTicketCost = Val(txtTicketCost.Text)
'Calculations
intTotalTickets = intWallet \ intTicketCost
intChange = intWallet Mod intTicketCost
intTotalCost = intTotalTickets * intTicketCost
'Output
picChange.Cls
picTotalCost.Cls
lblTickets.Caption = "you can buy " & intTotalTickets & " tickets"
picChange.Print "your total change is " & Format(intChange, "currency")
picTotalCost.Print "you total cost is " & Format(intTotalCost, "currency")


End Sub
