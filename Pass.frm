VERSION 5.00
Begin VB.Form Pass 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password Required"
   ClientHeight    =   1185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3225
   Icon            =   "Pass.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1185
   ScaleWidth      =   3225
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton K 
      Caption         =   "OK"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   3015
   End
   Begin VB.TextBox Pass 
      Height          =   255
      Left            =   120
      MaxLength       =   8
      TabIndex        =   0
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      Caption         =   "Password to open program with:"
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
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "Pass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub K_Click()
On Error GoTo err
Open App.Path & "\enc.dll" For Input As #1
Line Input #1, St1
Line Input #1, St2
Close #1

For i = Len(St2) To 1 Step -1
    Strn = Strn & Mid(St2, i, 1)
Next i
St2 = Strn

For i = 0 To 127
    St1 = Replace(St1, Chr(i + 127), Chr(i))
    St2 = Replace(St2, Chr(i + 127), Chr(i))
Next i

If "pass = " & Pass.Text = St2 Then
    Crypter.Men1(0).Visible = True
    Crypter.Men2(0).Visible = True
    Unload Me
Else
Beep
End If
err:
Exit Sub
End Sub
