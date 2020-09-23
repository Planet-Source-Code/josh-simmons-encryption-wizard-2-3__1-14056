VERSION 5.00
Begin VB.Form FileDat 
   BackColor       =   &H80000012&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Data"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5835
   Icon            =   "FileDat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   5835
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton K 
      Caption         =   "View"
      Height          =   270
      Left            =   4920
      TabIndex        =   5
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox Num 
      Height          =   255
      Left            =   4200
      TabIndex        =   4
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label5 
      BackColor       =   &H00404040&
      Caption         =   "F i l e    P  a  t  h:"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label FPath 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   9
      Top             =   960
      Width           =   3735
   End
   Begin VB.Label EncN 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2640
      TabIndex        =   8
      Top             =   1320
      Width           =   3015
   End
   Begin VB.Label FName 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   7
      Top             =   240
      Width           =   3735
   End
   Begin VB.Label FSize 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   6
      Top             =   600
      Width           =   3735
   End
   Begin VB.Label Label4 
      BackColor       =   &H00404040&
      Caption         =   "View File Encrytion If Encryption Number Is...."
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   3855
   End
   Begin VB.Label Label3 
      BackColor       =   &H00404040&
      Caption         =   "File Encryption Number:"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404040&
      Caption         =   "F i l e    S  i  z  e:"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      Caption         =   "F i l e   N a m e:"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "FileDat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
FSize.Caption = Crypter.Sze.Caption
FSize.ToolTipText = Crypter.Sze.Caption

FNam = Split(Crypter.Filelbl.Caption, "\")     'Get the file name from the path
FNam2 = FNam(UBound(FNam))                     'and display....
FName.Caption = FNam2
FName.ToolTipText = FNam2

For i = 0 To UBound(FNam) - 1
    FPth = FPth & FNam(i) & "\"                'Get the path and remove the file
Next i                                         'name from the string to display
FPath.Caption = FPth                           'to the user....
FPath.ToolTipText = FPth


EncN.Caption = Crypter.RN.Text
End Sub

Private Sub K_Click()
If Val(Num.Text) > 127 Then Exit Sub
If Val(Num.Text) < 1 Then Exit Sub
ViewEnc.Show
End Sub
