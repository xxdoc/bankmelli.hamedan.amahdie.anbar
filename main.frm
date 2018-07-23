VERSION 5.00
Begin VB.Form main 
   Caption         =   "Anbar"
   ClientHeight    =   10185
   ClientLeft      =   2025
   ClientTop       =   645
   ClientWidth     =   11925
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   10185
   ScaleWidth      =   11925
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   4080
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "‰”ŒÂ Â«Ì Ì«›  ‘œÂ"
      Height          =   2295
      Left            =   8400
      TabIndex        =   0
      Top             =   6840
      Visible         =   0   'False
      Width           =   3375
      Begin Project1.Button Button1 
         Height          =   375
         Left            =   1920
         TabIndex        =   2
         Top             =   1800
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Caption         =   "»” ‰ ﬂ«œ—"
      End
      Begin VB.FileListBox File1 
         BackColor       =   &H00FFFFFF&
         Height          =   1455
         Left            =   120
         Pattern         =   "*.mdb"
         TabIndex        =   1
         Top             =   240
         Width           =   3135
      End
      Begin VB.Image Image2 
         Height          =   255
         Left            =   0
         Picture         =   "main.frx":628A
         Stretch         =   -1  'True
         Top             =   0
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   255
         Left            =   0
         Picture         =   "main.frx":659E
         Stretch         =   -1  'True
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Frame Frame2 
      Height          =   135
      Left            =   -240
      TabIndex        =   5
      Top             =   9120
      Width           =   13335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   0
      TabIndex        =   3
      Top             =   9240
      Width           =   12855
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   8160
         TabIndex        =   16
         Top             =   0
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright By:NasserNiazy and MasoudAlavi      2008 All RightReserved.                                           GHYESHSOFT.Inc"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   50
      Left            =   0
      Top             =   3720
   End
   Begin Project1.Buttonl Button 
      Height          =   375
      Index           =   0
      Left            =   10320
      TabIndex        =   6
      Top             =   1080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "„‰ÊÌ «’·Ì"
   End
   Begin Project1.Buttonl Button 
      Height          =   375
      Index           =   9
      Left            =   10320
      TabIndex        =   7
      Top             =   4320
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "Œ—ÊÃ"
   End
   Begin Project1.Buttonl Button 
      Height          =   375
      Index           =   2
      Left            =   10320
      TabIndex        =   8
      Top             =   1800
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "„ÿ»Ê⁄« "
   End
   Begin Project1.Buttonl Button 
      Height          =   375
      Index           =   3
      Left            =   10320
      TabIndex        =   9
      Top             =   2160
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "‰„Ê‰Â Â«"
   End
   Begin Project1.Buttonl Button 
      Height          =   375
      Index           =   1
      Left            =   10320
      TabIndex        =   10
      Top             =   1440
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "„·“Ê„« "
   End
   Begin Project1.Buttonl Button 
      Height          =   375
      Index           =   7
      Left            =   10320
      TabIndex        =   13
      Top             =   3600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "„‘«ÂœÂ ‰”ŒÂ Å‘ Ì»«‰"
   End
   Begin Project1.Buttonl Button 
      Height          =   375
      Index           =   8
      Left            =   10320
      TabIndex        =   14
      Top             =   3960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "„«‘Ì‰ Õ”«»"
   End
   Begin Project1.Buttonl Button 
      Height          =   375
      Index           =   6
      Left            =   10320
      TabIndex        =   12
      Top             =   3240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "«ÌÃ«œ ‰”ŒÂ Å‘ Ì»«‰"
   End
   Begin Project1.Buttonl Button 
      Height          =   375
      Index           =   5
      Left            =   10320
      TabIndex        =   11
      Top             =   2880
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "œ—»«—Â „«"
   End
   Begin Project1.Buttonl Button 
      Height          =   375
      Index           =   4
      Left            =   10320
      TabIndex        =   17
      Top             =   2520
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "«›“Êœ‰ ‘⁄»Â"
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "»«‰ﬂ „·Ì «Ì—«‰"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1095
      Left            =   3600
      TabIndex        =   15
      Top             =   0
      Width           =   4335
   End
   Begin VB.Image Image4 
      Height          =   975
      Left            =   0
      Picture         =   "main.frx":68E2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11895
   End
   Begin VB.Image Image3 
      Height          =   8415
      Left            =   0
      Picture         =   "main.frx":6E3E
      Stretch         =   -1  'True
      Top             =   840
      Width           =   11895
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Use Option Explicit to avoid implicitly creating variables of type Variant         FixIT90210ae-R383-H1984
Dim J%
Dim i%, dr As Boolean
Private Sub Button_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
For i = 0 To 8
If i <> Index Then _
Me.Button(i).Refrash
Next
End Sub

Private Sub Button1_Click()
On Error GoTo 4
Frame6.Visible = False
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" button1_click() Of Main,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub

Private Sub Button2_Click()
On Error GoTo 4
Dim Fname$
Fname = InputBox("·ÿ›¬ „”Ì— —« »—«Ì –ŒÌ—Â »«‰ﬂ «ÿ·«⁄« Ì  «∆Ì‰ ﬂ‰Ìœ", "«‰»«—", App.Path + "\bakup")
If asd.FolderExists(Fname) = True Then
        asd.CopyFile App.Path & "\HAVALE.MDB", Fname & "\haveleh-bakuped " & Date$ & Str$(Hour(Time)) & "-" & CStr(Minute(Time)) & ".mdb"
End If
MsgBox "Å‘ Ì»«‰ êÌ—Ì ﬂ«„· ‘œ"
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" button2_click() Of Main,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)

End Sub

Private Sub Button3_Click()
On Error GoTo 4
Dim Fname$
Fname = InputBox("·ÿ›¬ „”Ì— —« »—«Ì Ã” ÃÊ »«‰ﬂ «ÿ·«⁄« Ì  «∆Ì‰ ﬂ‰Ìœ", "«‰»«—", App.Path + "\bakup")
If asd.FolderExists(Fname) = True Then
        Frame6.Visible = True
        File1.Path = Fname
End If
'If File1.ListCount = 0 Then Frame6.Visible = False
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" button3_click() Of Main,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)

End Sub

Private Sub Button4_Click()
    On Error GoTo 4
Unload Me
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" button4_click() Of Main,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub

Private Sub File1_Click()
On Error GoTo 4
Dim Pathdb$
Pathdb = File1.Path + "\" + File1.FileName
 If cnn.State = adStateOpen Then cnn.Close
 cnn.Mode = adModeRead
 cnn.Provider = "Microsoft.Jet.OLEDB.4.0"
 cnn.Open Pathdb
 If Err.Number Then MsgBox "error in connecting dbass"
FrmView.Show
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" File1_Click() Of Main,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub

Private Sub Form_Load()
On Error GoTo 4
Dim NB1$, NB2$
Me.BackColor = RGB(232, 237, 240): Frame2.BackColor = RGB(232, 237, 240)
Label1.Caption = " «„—Ê“: " & IRDate()
'------------------------------------------------------
NB1 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
NB2 = ";Persist Security Info=False"
If Len(App.Path) = 3 Then
Kalmdb = App.Path & "NEWBANK.MDB"
Havmdb = App.Path & "HAVALE.mdb"
Shomdb = App.Path & "SHOAB.MDB"
Else
Kalmdb = App.Path & "\NEWBANK.MDB"
Havmdb = App.Path & "\HAVALE.mdb"
Shomdb = App.Path & "\SHOAB.MDB"
End If
If asd.FileExists(Kalmdb) = False _
Or asd.FileExists(Havmdb) = False _
Or asd.FileExists(Shomdb) = False Then
MsgBox "›«Ì·Â«Ì «’·Ì »—‰«„Â ÅÌœ« ‰‘œ .·ÿ›¬ œÊ»«—Â ¬‰ —« ‰’» ﬂ‰Ìœ", vbCritical, "Œÿ«Ì »«—ê“«—Ì"
Unload Me
Exit Sub
End If
NB1 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
NB2 = ";Persist Security Info=False"
If Len(App.Path) = 3 Then
Kalmdb = NB1 & Kalmdb & NB2
Havmdb = NB1 & Havmdb & NB2
Shomdb = NB1 & Shomdb & NB2
Else
Kalmdb = NB1 & Kalmdb & NB2
Havmdb = NB1 & Havmdb & NB2
Shomdb = NB1 & Shomdb & NB2
End If
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" Form_Load() Of Main,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub



Private Sub Form_Resize()
On Error GoTo 4
Me.Enabled = False
4 Call Loger(" Form_Resize() Of Main,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Error GoTo 4
'// Nasser
Dim FF$
FF = GetSetting("anbar", "day", "frm", "None")
If FF = "" Or FF = "None" Then
        Call SaveSetting("anbar", "day", "frm", Date)
        asd.CopyFile App.Path + "\havale.mdb", App.Path & "\bakup\" & Date$ & ".mdb"
Loger "===================Log for " & Date$ & "================================"
Else
        If Val(Day(Date)) <> Val(Day(FF)) Then
        Call SaveSetting("anbar", "day", "frm", Date)
        asd.CopyFile App.Path + "\HAVALE.MDB", App.Path & "\bakup\" & Date$ & ".mdb"
        Loger "===================Log for " & Date$ & "================================"
        End If
End If
FF = "":
End
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" Form_UnLoad() Of Main,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description & " - FF:" & FF)
End Sub



Private Sub Frame6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo 4
Button1.Refrash
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" Frame6_Click() Of Main,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub

Private Sub Image2_Click()
On Error GoTo 4
Frame6.Visible = False: Image2.Visible = True
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" image2_Click() Of Main,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo 4
Image2.Visible = False
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" image2_MouseUP() Of Main,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub

Private Sub Img_Click(Index As Integer)
On Error GoTo 4
Select Case Index
 Case 1: Havaleh.Show
 Case 2: matmain.Show
 Case 3: MALMAIN.Show
 Case 4: frmSplash.Timer1.Enabled = False: frmSplash.Show: frmSplash.Enabled = True: frmSplash.Label1 = "About Form"
 Case 0: shoab.Show
End Select
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" Img_Click(" & Index & ") Of Main,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub


Private Sub Button_Click(Index As Integer)
Select Case (Index)
Case 0:
dr = Not dr
Timer3.Enabled = True
Case 1:
MALMAIN.Show
Case 2:
matmain.Show
Case 3:
Havaleh.Show
Case 9:
Unload Me
Case 5: frmSplash.Timer1.Enabled = False: frmSplash.Show: frmSplash.Enabled = True: frmSplash.Label1 = "About Form"
Case 6:
Dim Fname$
Fname = InputBox("·ÿ›¬ „”Ì— —« »—«Ì –ŒÌ—Â »«‰ﬂ «ÿ·«⁄« Ì  «∆Ì‰ ﬂ‰Ìœ", "«‰»«—", App.Path + "\bakup")
If asd.FolderExists(Fname) = True Then
        asd.CopyFile App.Path & "\HAVALE.MDB", Fname & "\haveleh-bakuped " & Date$ & Str$(Hour(Time)) & "-" & CStr(Minute(Time)) & ".mdb"
End If
MsgBox "Å‘ Ì»«‰ êÌ—Ì ﬂ«„· ‘œ"
Case 7:
Fname = InputBox("·ÿ›¬ „”Ì— —« »—«Ì Ã” ÃÊ »«‰ﬂ «ÿ·«⁄« Ì  «∆Ì‰ ﬂ‰Ìœ", "«‰»«—", App.Path + "\bakup")
If asd.FolderExists(Fname) = True Then
        Frame6.Visible = True
        File1.Path = Fname
End If
Case 8: Shell "calc.exe", vbNormalFocus
Case 4: shoab.Show 1
End Select


End Sub


Private Sub Timer2_Timer()
On Error GoTo 4
Me.Enabled = True
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" Timer2_timer() Of Main,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)

End Sub
Private Sub Timer3_Timer()
If dr = False Then
For i = 0 To 9
   If Button(i).Top <= 1080 Then
   Button(i).Top = 1080
   Else
   Button(i).Top = Button(i).Top - 50
   End If
Next
Else
For i = 0 To 9
   If Button(i).Top >= (i * 375) + 1080 Then
   Button(i).Top = (i * 375) + 1080
   Else
   Button(i).Top = Button(i).Top + 50
   End If
Next
End If
End Sub
