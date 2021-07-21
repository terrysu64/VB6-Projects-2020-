VERSION 5.00
Begin VB.Form frmShapeGenerator 
   Caption         =   "Shape generator"
   ClientHeight    =   8265
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12330
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8265
   ScaleWidth      =   12330
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picOutput 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   2280
      ScaleHeight     =   6435
      ScaleWidth      =   9795
      TabIndex        =   9
      Top             =   1560
      Width           =   9855
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   7
      Top             =   5880
      Width           =   1935
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
      Left            =   2280
      TabIndex        =   2
      Text            =   "(enter shape size)"
      Top             =   840
      Width           =   9855
   End
   Begin VB.Frame fraOptions 
      Caption         =   "Options"
      Height          =   4815
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1935
      Begin VB.OptionButton optDiamondH 
         Caption         =   "Diamond (hollow)"
         Height          =   615
         Left            =   120
         TabIndex        =   11
         Top             =   3960
         Width           =   1455
      End
      Begin VB.OptionButton optDiamond 
         Caption         =   "Diamond"
         Height          =   615
         Left            =   120
         TabIndex        =   10
         Top             =   3240
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1920
         TabIndex        =   8
         Top             =   3240
         Width           =   135
      End
      Begin VB.OptionButton optTriangleH 
         Caption         =   "Right triangle (hollow)"
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   2520
         Width           =   1575
      End
      Begin VB.OptionButton optTriangle 
         Caption         =   "Right triangle"
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   1800
         Width           =   1695
      End
      Begin VB.OptionButton optSquareH 
         Caption         =   "Square (hollow)"
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   1575
      End
      Begin VB.OptionButton optSquare 
         Caption         =   "Square"
         Height          =   615
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Please select a shape to generate"
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
      Width           =   12015
   End
End
Attribute VB_Name = "frmShapeGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Terry Su
'Date: May 13, 2020
'Purpose: to generate customizable shapes based off user inputs
Option Explicit

Private Sub cmdGo_Click()
'Declare
Dim intCount As Integer
Dim intRows As Integer
Dim intRows2 As Integer
Dim intColumns As Integer
Dim intColumns2 As Integer
Dim intStop As Integer
Dim intCount2 As Integer

'Initialize
intCount = 0
intCount2 = 0
intRows = 0
intRows2 = 0
intColumns = 0
intColumns2 = 0
intStop = 0

'Input
intStop = Val(txtInput.Text)

'Process/Output
'all shape sizes must be > 0
If intStop <= 0 Then
    MsgBox "shape size must be greater than 0", vbCritical, "Warning"
    End
End If

'clear everything before starting any code execution
picOutput.Cls

'Make sure an option is selected
If Option1.Value = True Then
    MsgBox "please select a shape option", vbExclamation, "Warning"

'full square
ElseIf optSquare.Value = True Then
    For intRows = 1 To intStop
        For intCount = 1 To intStop
            picOutput.Print "*  ";
        Next intCount
        picOutput.Print ""
    Next intRows

'hollow square
ElseIf optSquareH.Value = True Then
    'if size is 1
    If intStop = 1 Then
        picOutput.Print "*"
    Else
        'first row
        For intColumns = 1 To intStop
            picOutput.Print "*  ";
        Next intColumns
     picOutput.Print ""
        
        'all middle rows  are composed of a "*", a certain amount of " "s and another *
        For intCount = 1 To intStop - 2
            picOutput.Print "*";
            For intRows = 0 To 2 + (3 * (intStop - 2))
                picOutput.Print " ";
            Next intRows
            picOutput.Print "*"
        Next intCount

        'last row same as first row
        For intColumns = 1 To intStop
            picOutput.Print "*  ";
        Next intColumns
    End If

'full triangle
ElseIf optTriangle.Value = True Then
    
    'start with 1 "*" in row 1, 2 "*"s in row 2... n "*"s in row n
    For intRows = 1 To intStop
    intCount = intCount + 1
    intColumns = intStop - intCount
        Do While intColumns < intStop
            picOutput.Print "*  ";
            intColumns = intColumns + 1
        Loop
    picOutput.Print ""
    Next intRows

ElseIf optTriangleH.Value = True Then
    'if size is 1
    If intStop = 1 Then
        picOutput.Print "*"
    Else
    
        'first row
        picOutput.Print "*"
        
        'all middle rows composed of a "*", a certain amount of " "s and a last "*"
        For intRows = 1 To intStop - 2
            picOutput.Print "*";
            intCount = 2 + 3 * intCount2
            Do While intColumns < intCount
                picOutput.Print " ";
                intColumns = intColumns + 1
            Loop
            picOutput.Print "*"
            intColumns = 0
            intCount2 = intCount2 + 1
        Next intRows
        
        'last row
        For intColumns = 1 To intStop
        picOutput.Print "*  ";
        Next intColumns
    End If

ElseIf optDiamond.Value = True Then
    
    'all diamond sizes must be odd values
    If intStop Mod 2 = 0 Then
        MsgBox "diamond sizes must be an odd value", vbCritical, "Warning"
        End
    End If
    
    'all rows are composed of a certain amount of " "s, and a certain amount of "*  "s
    'first half of diamond
    For intRows = 1 To (intStop \ 2) + 1
        For intColumns = 1 To 3 * ((intStop \ 2) - intCount)
            picOutput.Print " ";
        Next intColumns
    
        For intRows2 = 1 To intCount2 + 1
            picOutput.Print "*  ";
        Next intRows2
            picOutput.Print ""
            
        intCount2 = intCount2 + 2
        intCount = intCount + 1
    Next intRows

    intCount = 1
    intCount2 = 0
    intRows = 0
    intRows2 = 0
    intColumns = 0
    
    'second half of the diamond
    'composed in the same way as the first row
    For intRows = 1 To intStop \ 2
        For intColumns = 1 To 3 * intCount
            picOutput.Print " ";
        Next intColumns
        For intRows2 = 1 To intStop - 2 * intCount
            picOutput.Print "*  ";
        Next intRows2
        picOutput.Print ""
        intCount = intCount + 1
        
    Next intRows

'hollow diamond
ElseIf optDiamondH.Value = True Then
    'size values must be odd again
    If intStop Mod 2 = 0 Then
        MsgBox "diamond sizes must be an odd value", vbCritical, "Warning"
        End
    End If
    
    intCount2 = 1
    
    'first row
    For intColumns = 1 To 3 * (intStop \ 2)
        picOutput.Print " ";
    Next intColumns
    picOutput.Print "*"
    
    'rows from 2 to intstop - 1
    For intRows = 1 To (intStop \ 2) - 1
        For intColumns = 1 To 3 * ((intStop \ 2) - intCount - 1)
            picOutput.Print " ";
        Next intColumns
        picOutput.Print "*";
        
        For intColumns2 = 1 To 2 + (3 * intCount2)
            picOutput.Print " ";
        Next intColumns2
        picOutput.Print "*"
            
        intCount2 = intCount2 + 2
        intCount = intCount + 1
    Next intRows

    intCount = 0
    intCount2 = 0
    intRows = 0
    intRows2 = 0
    intColumns = 0
    intColumns2 = 0
    
    'middle row
    picOutput.Print "*";
    For intColumns = 1 To 4 + (3 * (intStop - 3))
        picOutput.Print " ";
    Next intColumns
    picOutput.Print "*"
    
    intColumns = 0
    intCount = 1
    
    'second half except last row
    For intRows = 1 To (intStop \ 2) - 1
        
        For intColumns = 1 To 3 * intCount
            picOutput.Print " ";
        Next intColumns
        picOutput.Print "*";
        
        For intColumns2 = 1 To 2 + 3 * (intStop - 4 - intCount2)
            picOutput.Print " ";
        Next intColumns2
        picOutput.Print "*"
    
    intCount = intCount + 1
    intCount2 = intCount2 + 2
    
    Next intRows
    'last row
    For intColumns = 1 To 3 * (intStop \ 2)
        picOutput.Print " ";
    Next intColumns
    picOutput.Print "*"

End If
            
End Sub

Private Sub Form_Load()
Option1.Value = True
End Sub

Private Sub optDiamond_Click()
lblMessage.Caption = "For diamonds the shape size must be odd!"
End Sub

Private Sub optDiamondH_Click()
lblMessage.Caption = "For diamonds the shape size must be odd!"
End Sub

Private Sub optSquare_Click()
lblMessage.Caption = "Please select a shape to generate"
End Sub

Private Sub optSquareH_Click()
lblMessage.Caption = "Please select a shape to generate"
End Sub

Private Sub optTriangle_Click()
lblMessage.Caption = "Please select a shape to generate"
End Sub

Private Sub optTriangleH_Click()
lblMessage.Caption = "Please select a shape to generate"
End Sub

Private Sub txtInput_Click()
txtInput.Text = ""
End Sub
