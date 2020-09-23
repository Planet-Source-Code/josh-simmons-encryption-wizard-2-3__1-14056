VERSION 5.00
Begin VB.Form ProgressDec 
   BackColor       =   &H80000007&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Decrypting..."
   ClientHeight    =   495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8055
   Icon            =   "ProgressDec.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   495
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox T 
      Height          =   285
      Left            =   2280
      TabIndex        =   2
      Text            =   "0"
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Timer ti 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8040
      Top             =   120
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   4800
      X2              =   4800
      Y1              =   120
      Y2              =   480
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   3000
      X2              =   3000
      Y1              =   120
      Y2              =   480
   End
   Begin VB.Label Tme 
      BackColor       =   &H00808080&
      Caption         =   " Elapsed:"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3000
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7815
   End
End
Attribute VB_Name = "ProgressDec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ti_Timer()
TT = Val(T.Text) + 1
T.Text = TT
End Sub
