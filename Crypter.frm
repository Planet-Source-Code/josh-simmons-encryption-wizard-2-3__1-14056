VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Crypter 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Encrypter"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8325
   Icon            =   "Crypter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Crypter.frx":000C
   ScaleHeight     =   5325
   ScaleWidth      =   8325
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox QuickCrypt 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "  QuickCrypt"
      Top             =   4920
      Width           =   1335
   End
   Begin VB.TextBox Men2 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   6
      Left            =   4800
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   17
      TabStop         =   0   'False
      Text            =   "Unpass. Protect"
      Top             =   4560
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Men2 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   5
      Left            =   4800
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   16
      TabStop         =   0   'False
      Text            =   "Pass. Protect"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Men2 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   4
      Left            =   4800
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   15
      TabStop         =   0   'False
      Text            =   "Show Ela. Time"
      Top             =   4080
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Men2 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   3
      Left            =   3480
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   14
      TabStop         =   0   'False
      Text            =   "Clear All"
      Top             =   4080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Men2 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   2
      Left            =   3480
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   13
      TabStop         =   0   'False
      Text            =   "File Data"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Men2 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   1
      Left            =   3480
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   12
      TabStop         =   0   'False
      Text            =   "Set Encrypt #"
      Top             =   4560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Men2 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   3480
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   "   O p t i o n s"
      Top             =   4920
      Width           =   1335
   End
   Begin VB.TextBox Men1 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   3
      Left            =   1920
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "Open File"
      Top             =   4080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Men1 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   2
      Left            =   1920
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "Save/Encrypt"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Men1 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   1920
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "Open/Decrypt"
      Top             =   4560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Men1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   1920
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "          F i l e"
      Top             =   4920
      Width           =   1335
   End
   Begin VB.TextBox RN 
      Height          =   285
      Left            =   4800
      TabIndex        =   4
      Top             =   5880
      Width           =   3015
   End
   Begin VB.TextBox DLine 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   5400
      Width           =   4215
   End
   Begin MSComDlg.CommonDialog CMD 
      Left            =   240
      Top             =   6720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RLine 
      Height          =   4695
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "This displays the data of the file that is currently open or the file that you are creating."
      Top             =   120
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   8281
      _Version        =   393217
      BorderStyle     =   0
      ScrollBars      =   2
      TextRTF         =   $"Crypter.frx":597B
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      BorderWidth     =   5
      X1              =   480
      X2              =   480
      Y1              =   4920
      Y2              =   5280
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   5
      X1              =   0
      X2              =   8400
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      BorderWidth     =   5
      X1              =   8280
      X2              =   8280
      Y1              =   5280
      Y2              =   0
   End
   Begin VB.Line Men1L 
      BorderWidth     =   2
      Index           =   5
      X1              =   1920
      X2              =   3240
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Line Line12 
      BorderColor     =   &H80000009&
      BorderWidth     =   4
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   5280
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000009&
      BorderWidth     =   5
      X1              =   8280
      X2              =   0
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Line Line9 
      BorderColor     =   &H80000009&
      BorderWidth     =   5
      X1              =   3480
      X2              =   3480
      Y1              =   4920
      Y2              =   5280
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000009&
      BorderWidth     =   5
      X1              =   4920
      X2              =   4920
      Y1              =   5280
      Y2              =   4920
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      BorderWidth     =   5
      X1              =   1920
      X2              =   1920
      Y1              =   4920
      Y2              =   5280
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      BorderWidth     =   5
      X1              =   3360
      X2              =   3360
      Y1              =   5280
      Y2              =   4920
   End
   Begin VB.Label Filelbl 
      BackColor       =   &H00404040&
      Caption         =   "Untitled.txt"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   6240
      Width           =   6615
   End
   Begin VB.Label Sze 
      BackColor       =   &H00000000&
      Caption         =   "0 Bytes"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5040
      TabIndex        =   2
      Top             =   6600
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808080&
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   4920
      Width           =   135
   End
   Begin VB.Label DteTme 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      Top             =   4920
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Height          =   375
      Left            =   0
      TabIndex        =   18
      Top             =   4920
      Width           =   1935
   End
End
Attribute VB_Name = "Crypter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub CreateNew_Click()
RLine.Text = ""
Filelbl.Caption = "Untitled.txt"
End Sub

Private Sub DteTme_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
For i = 1 To Men2.UBound
    Men2(i).Visible = False
Next i
For i = 1 To Men1.UBound
    Men1(i).Visible = False
Next i
Men1(0).ForeColor = &H0&
Men1(0).BackColor = &H404040
Men2(0).ForeColor = &H0&
Men2(0).BackColor = &H404040
End Sub

Private Sub Form_Load()
On Error GoTo err
Open App.Path & "\enc.dll" For Input As #1
Line Input #1, St1
Line Input #1, St2
Close #1
For i = 0 To Men1.UBound
    Men1(i).Visible = False
Next i
For i = 0 To Men2.UBound
    Men2(i).Visible = False
Next i
Pass.Show
err:
End Sub

Private Sub Men1_Click(Index As Integer)
Men1(Index).SelStart = 0
If Index = 0 Then
    For i = 1 To Men1.UBound
        Men1(i).Visible = True
    Next i
End If
If Index = 1 Then
    Call GlobDecrypt
End If
If Index = 2 Then
    Call GlobEncrypt
End If
If Index = 3 Then
    End
End If
If Index = 4 Then
    Call RegOpen
End If
End Sub

Private Sub Men1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
For i = 1 To Men2.UBound
    Men2(i).Visible = False
Next i
Men1(0).ForeColor = &H0&
Men1(0).BackColor = &H404040
Men2(0).ForeColor = &H0&
Men2(0).BackColor = &H404040
For i = 1 To Men1.UBound
    Men1(i).BackColor = &H80FF&
    Men1(i).ForeColor = &HFFFFFF
Next i
If Index > 0 Then Men1(Index).ForeColor = &HFFFFFF
If Index > 0 Then Men1(Index).BackColor = &HFF0000
End Sub

Private Sub Men2_Click(Index As Integer)
Men2(Index).SelStart = 0
If Index = 0 Then
    For i = 1 To Men2.UBound
        Men2(i).Visible = True
    Next i
End If
If Index = 1 Then
    SetEnc.Show
End If
If Index = 2 Then
    FileDat.Show
End If
If Index = 3 Then
    Filelbl.Caption = "Untitled.txt"
    Filelbl.ToolTipText = Filelbl.Caption
    RLine.Text = ""
    RN.Text = "Not encrypted yet!"
    Sze.Caption = "0 Bytes"
End If
If Index = 4 Then
    If Men2(4).Text = "Hide Ela. Time" Then
        Men2(4).Text = "Show Ela. Time"
        Exit Sub
    End If
    If Men2(4).Text = "Show Ela. Time" Then
        Men2(4).Text = "Hide Ela. Time"
        Exit Sub
    End If
End If
If Index = 5 Then
    PassProt.Show
End If
If Index = 6 Then
    On Error Resume Next
    Kill App.Path & "\enc.dll"
End If
End Sub

Private Sub Men2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
For i = 1 To Men1.UBound
    Men1(i).Visible = False
Next i
Men1(0).ForeColor = &H0&
Men1(0).BackColor = &H404040
Men2(0).ForeColor = &H0&
Men2(0).BackColor = &H404040
For i = 1 To Men2.UBound
    Men2(i).BackColor = &H80FF&
    Men2(i).ForeColor = &HFFFFFF
Next i
If Index > 0 Then Men2(Index).ForeColor = &HFFFFFF
If Index > 0 Then Men2(Index).BackColor = &HFF0000
End Sub

Private Sub RegOpen()
CMD.ShowOpen

On Error GoTo Erry
Open CMD.FileName For Input As #1
Do Until EOF(1)
    Line Input #1, LData
    Strng = Strng & LData & vbLf
Loop
RLine.Text = Strng
Close #1

Filelbl.Caption = CMD.FileName

Erry:
Exit Sub
End Sub

Private Sub Decrypt_Click()
    Call GlobDecrypt
End Sub

Private Sub GlobDecrypt()
CMD.ShowOpen

On Error GoTo err
Open CMD.FileName For Input As #1
Do Until EOF(1)
    Line Input #1, LData
    Strng = Strng & LData & vbLf
Loop
RLine.Text = Strng
Close #1

Filelbl.Caption = CMD.FileName
Filelbl.ToolTipText = Filelbl.Caption

FSize = Len(Strng)
If FSize < 1024 Then Sze.Caption = "File Size: " & FSize & " Bytes"
If FSize > 1024 And FSize < 1024000 Then
    FSize = FSize / 1024
    FSize = CInt(FSize)
    Sze.Caption = "File Size: " & FSize & " Kilobytes"
End If
If FSize > 1024000 And FSize < 1024000000 Then
    FSize = FSize / 1024000
    FSize = CInt(FSize)
    Sze.Caption = "File Size: " & FSize & " Megabytes!"
End If
If FSize > 1024000000 Then
    FSize = FSize / 1024000000
    FSize = CInt(FSize)
    Sze.Caption = "File Size: " & FSize & " Gigabytes!?!?!"
End If
            'Decrypting!
            
RL2 = RLine.Text

LenRanNumT = Mid(RL2, Len(RL2) - 1, 1)
If LenRanNumT = "!" Then LenRanNum = 1
If LenRanNumT = "@" Then LenRanNum = 2
If LenRanNumT = "#" Then LenRanNum = 3
RanNumT = Mid(RL2, Len(RL2) - LenRanNum - 1, LenRanNum)
RanNumT = Replace(RanNumT, "!", "1")
RanNumT = Replace(RanNumT, "@", "2")
RanNumT = Replace(RanNumT, "#", "3")
RanNumT = Replace(RanNumT, "$", "4")
RanNumT = Replace(RanNumT, "%", "5")
RanNumT = Replace(RanNumT, "^", "6")
RanNumT = Replace(RanNumT, "&", "7")
RanNumT = Replace(RanNumT, "*", "8")
RanNumT = Replace(RanNumT, "(", "9")
RanNumT = Replace(RanNumT, ")", "0")
RanNum = RanNumT
RN.Text = ""
RN.Text = RanNum
RL2 = Mid(RL2, 1, Len(RL2) - LenRanNum - 2)

If Men2(4).Text = "Hide Ela. Time" Then
    Y = 0
End If
If Men2(4).Text = "Show Ela. Time" Then
    ProgressDec.Show
    ProgressDec.ti.Enabled = True
    Y = 1
End If
For i = 255 To 128 Step -1
        RL2 = Replace(RL2, Chr(i - 127 + RanNum), Chr(i))
        If Y = 1 Then DoEvents
Next i
For i = 0 To 127
        RL2 = Replace(RL2, Chr(i + 127), Chr(i))
        If Y = 1 Then DoEvents
Next i
If Y = 1 Then
    ProgressDec.Tme.Caption = " Elapsed: " & ProgressDec.T.Text
    ProgressDec.ti.Enabled = False
    ProgressDec.T.Text = "0"
End If

RLine.Text = ""
RLine.Text = RL2

err:
On Error GoTo Errr
Close #1
Exit Sub
Errr:
Exit Sub
End Sub

Private Sub Encrypt_Click()
    Call GlobEncrypt
End Sub

Private Sub GlobEncrypt()

RL = RLine.Text

For i = 0 To 127
        RL = Replace(RL, Chr(i), Chr(i + 127))
Next i
Ran1 = Int(Rnd * 126) + 1
For i = 128 To 255
        RL = Replace(RL, Chr(i), Chr(i - 127 + Ran1))
Next i

RN.Text = ""
RN.Text = Ran1

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

DLine.Text = RL & RQ

 CMD.ShowSave
 On Error GoTo Errt
 Open CMD.FileName For Output As #1
 Print #1, DLine.Text
 Close #1
 
 Filelbl.Caption = CMD.FileName
 
Errt:
Exit Sub
End Sub

Private Sub Form_Terminate()
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Lnk_Click()
Call ShellExecute(0, "open", "http://www.biggergaming.com", "", "", SW_SHOWNORMAL)
End Sub

Private Sub Menu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
For i = 1 To Men2.UBound
    Men2(i).Visible = False
Next i
For i = 1 To Men1.UBound
    Men1(i).Visible = False
Next i
Men1(0).ForeColor = &H0&
Men1(0).BackColor = &H404040
Men2(0).ForeColor = &H0&
Men2(0).BackColor = &H404040
End Sub

Private Sub QuickCrypt_Click()
QuickCrypt.SelStart = 0
CMD.ShowOpen

On Error GoTo Erryy
Open CMD.FileName For Input As #1
Do Until EOF(1)
    Line Input #1, LData
    Strng = Strng & LData & vbLf
Loop
Close #1
RL = Strng
For i = 0 To 127
        RL = Replace(RL, Chr(i), Chr(i + 127))
Next i
Ran1 = Int(Rnd * 126) + 1
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

DLine.Text = RL & RQ

 CMD.ShowSave
 On Error GoTo Errty
 Open CMD.FileName For Output As #1
 Print #1, DLine.Text
 Close #1
 
 MsgBox "QuickCrypt was successful in encrypting and saving the file you selected to " _
 & CMD.FileName & ".", vbInformation, "Successful!"
 Exit Sub
Errty:
Erryy:
 MsgBox "QuickCrypt was unsuccessful in encrypting and saving the file you selected. There was an error in either opening or saving the file, please check that you do not already have the file open.", vbCritical, "Unsuccessful!"
Exit Sub
End Sub

Private Sub RLine_Change()
    If Len(RLine.Text) < 1024 Then Sze.Caption = Len(RLine.Text) & " Bytes"
    If Len(RLine.Text) > 1024 And Len(RLine.Text) < 1024000 Then
        FSize = Len(RLine.Text) / 1024
        Sze.Caption = FSize & " Kilobytes"
    End If
    If Len(RLine.Text) > 1024000 And Len(RLine.Text) < 1024000000 Then
        FSize = Len(RLine.Text) / 1024000
        Sze.Caption = FSize & " Megabytes!"
    End If
    If Len(RLine.Text) > 1024000000 Then
        FSize = Len(RLine.Text) / 1024000000
        FSize = CInt(FSize)
        Sze.Caption = FSize & " Gigabytes!?!?!"
    End If
    Sze.ToolTipText = Sze.Caption
End Sub

Private Sub RLine_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
For i = 1 To Men2.UBound
    Men2(i).Visible = False
Next i
For i = 1 To Men1.UBound
    Men1(i).Visible = False
Next i
Men1(0).ForeColor = &H0&
Men1(0).BackColor = &H404040
Men2(0).ForeColor = &H0&
Men2(0).BackColor = &H404040
End Sub
