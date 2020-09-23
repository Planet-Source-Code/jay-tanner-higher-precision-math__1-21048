VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " High-Precision Functions Test Form"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   4470
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton TanX_Button 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tan X"
      Height          =   330
      Left            =   1305
      TabIndex        =   3
      Top             =   900
      Width           =   645
   End
   Begin VB.TextBox Work 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   45
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1845
      Width           =   4380
   End
   Begin VB.CommandButton Cu_Root_Button 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cube Root"
      Height          =   330
      Left            =   3465
      TabIndex        =   5
      Top             =   900
      Width           =   960
   End
   Begin VB.TextBox Arg1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   45
      MaxLength       =   30
      TabIndex        =   0
      Text            =   "30"
      Top             =   495
      Width           =   4380
   End
   Begin VB.CommandButton SqRoot_Button 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Square Root"
      Height          =   330
      Left            =   2340
      TabIndex        =   4
      Top             =   900
      Width           =   1140
   End
   Begin VB.CommandButton CosX_Button 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cos X"
      Height          =   330
      Left            =   675
      TabIndex        =   2
      Top             =   900
      Width           =   645
   End
   Begin VB.CommandButton SinX_Button 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Sin X"
      Height          =   330
      Left            =   45
      TabIndex        =   1
      Top             =   900
      Width           =   645
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Arguments  in  Degrees"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   45
      TabIndex        =   9
      Top             =   1260
      Width           =   1905
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   45
      X2              =   4410
      Y1              =   1575
      Y2              =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Input  Argument"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   45
      TabIndex        =   7
      Top             =   270
      Width           =   4380
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Computed  Output"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   45
      TabIndex        =   6
      Top             =   1620
      Width           =   4380
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

  Option Explicit

' This is a program that shows how it is possible to compute
' certain basic math functions to a higher than 16 digits
' level of precision.
'
' Written by Jay Tanner - Jay@NeoProgrammics.com
'
' The functions developed so far are:
' Sine, Cosine, Tangent, Square Root and Cube Root
'
' Building on the methods used here, more high-precision
' math functions are possible.

' -------------------------------
  Private Sub SinX_Button_Click()

  Dim Q As Variant
  Dim A As Variant

' Check for non-numeric input argument
  If IsNumeric(Arg1) = False Then
     Work = "ERROR: Invalid input argument"
     Beep
     Exit Sub
  End If

' Convert degrees to radians
  Q = CDec(Arg1) * Pi / 180

' Compute sine of argument
  A = SinRad(Q)
  
  Work = A
  
  End Sub

' -------------------------------
  Private Sub CosX_Button_Click()

  Dim Q As Variant
  Dim A As Variant

' Check for non-numeric input argument
  If IsNumeric(Arg1) = False Then
     Work = "ERROR: Invalid input argument"
     Beep
     Exit Sub
  End If

' Convert degrees to radians
  Q = CDec(Arg1) * Pi / 180

' Compute sine of argument
  A = CosRad(Q)
  
  Work = A
  
  End Sub

' -------------------------------
  Private Sub TanX_Button_Click()

  Dim Q As Variant
  Dim A As Variant

' Check for non-numeric input argument
  If IsNumeric(Arg1) = False Then
     Work = "ERROR: Invalid input argument"
     Beep
     Exit Sub
  End If

' Check for infinite result
  If Abs(Arg1) = 90 Or Abs(Arg1) = 270 Then
     Work = "Infinite result"
     Exit Sub
  End If

' Convert degrees to radians
  Q = CDec(Arg1) * Pi / 180

' Compute tangent of argument
  A = TanRad(Q)
  
  Work = A

  End Sub

' ---------------------------------
  Private Sub SqRoot_Button_Click()

  Dim Q As Variant
  Dim A As Variant

' Check for non-numeric input argument
  If IsNumeric(Arg1) = False Then
     Work = "ERROR: Invalid input argument"
     Beep
     Exit Sub
  End If

' Compute square root of input argument
  Q = CDec(Arg1)
  A = Square_Root_Of(Q)
  
  Work = A
  
  End Sub

' ---------------------------------
  Private Sub Cu_Root_Button_Click()

  Dim Q As Variant
  Dim A As Variant

' Check for non-numeric input argument
  If IsNumeric(Arg1) = False Then
     Work = "ERROR: Invalid input argument"
     Beep
     Exit Sub
  End If

' Compute cube root of input argument
  Q = CDec(Arg1)
  A = Cube_Root_Of(Q)
  
  Work = A
 
  End Sub


