VERSION 5.00
Begin VB.Form SetEnc 
   BackColor       =   &H80000008&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Set File Encryption Number"
   ClientHeight    =   510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6840
   Icon            =   "SetEnc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   510
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Encrypt 
      Caption         =   "Encrypt"
      Height          =   285
      Left            =   5760
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox EncNum 
      Height          =   285
      Left            =   2640
      MaxLength       =   3
      TabIndex        =   1
      Top             =   120
      Width           =   3015
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
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "SetEnc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Encrypt_Click()
If Val(EncNum.Text) < 1 Then EncNum.Text = "1"
If Val(EncNum.Text) > 127 Then EncNum.Text = "127"

RL = Crypter.RLine.Text

For i = 0 To 127
        RL = Replace(RL, Chr(i), Chr(i + 127))
Next i
Ran1 = Val(EncNum.Text)
For i = 128 To 255
        RL = Replace(RL, Chr(i), Chr(i - 127 + Ran1))
Next i

Crypter.RN.Text = ""
Crypter.RN.Text = Ran1

RQ = Ran1
If Ran1 > 99 Then
    RQ = RQ & "3"
End If
If Ran1 > 9 And Ran1 < 100 Then
    RQ = RQ & "2"
End If
If Ran1 > 0 And Ran1 < 10 Then
    RQ = RQ & "1"
End If

RQ = Replace(RQ, "1", "!")
RQ = Replace(RQ, "2", "@")
RQ = Replace(RQ, "3", "#")
RQ = Replace(RQ, "4", "$")
RQ = Replace(RQ, "5", "%")
RQ = Replace(RQ, "6", "^")
RQ = Replace(RQ, "7", "&")
RQ = Replace(RQ, "8", "*")
RQ = Replace(RQ, "9", "(")
RQ = Replace(RQ, "0", ")")

Crypter.DLine.Text = RL & RQ

 Crypter.CMD.ShowSave
 On Error GoTo Errt
 Open Crypter.CMD.FileName For Output As #1
 Print #1, DLine.Text
 Close #1
 
 Filelbl.Caption = CMD.FileName
 
Errt:
Exit Sub

End Sub

Private Sub EncNum_Change()
If Val(EncNum.Text) < 1 Then EncNum.Text = "1"
If Val(EncNum.Text) > 127 Then EncNum.Text = "127"
End Sub
