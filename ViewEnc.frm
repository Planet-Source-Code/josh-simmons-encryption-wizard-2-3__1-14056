VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form ViewEnc 
   BackColor       =   &H80000012&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View File (Encrypted)"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8310
   Icon            =   "ViewEnc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   8310
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox RLine 
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   10186
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"ViewEnc.frx":000C
   End
End
Attribute VB_Name = "ViewEnc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
RL = Crypter.RLine.Text

For i = 0 To 127
        RL = Replace(RL, Chr(i), Chr(i + 127))
Next i
Ran1 = Val(FileDat.Num.Text)
For i = 128 To 255
        RL = Replace(RL, Chr(i), Chr(i - 127 + Ran1))
Next i

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

RLine.Text = RL & RQ
End Sub
