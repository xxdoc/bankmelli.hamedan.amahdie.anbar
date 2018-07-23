VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Havaleh 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«À«ÀÌÂ"
   ClientHeight    =   7605
   ClientLeft      =   2805
   ClientTop       =   3165
   ClientWidth     =   10350
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   10350
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   2640
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   495
      Left            =   2400
      Top             =   6000
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\project\1\meli.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\project\1\meli.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Kala"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Caption         =   "Frame7"
      Height          =   615
      Left            =   3480
      TabIndex        =   16
      Top             =   600
      Width           =   3495
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„ÿ»Ê⁄«                                      "
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   21.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Index           =   1
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   2415
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   615
      Left            =   6960
      TabIndex        =   15
      Top             =   600
      Width           =   3375
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "      «À«ÀÌÂ                   "
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   21.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Index           =   2
         Left            =   600
         TabIndex        =   20
         Top             =   -120
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      TabIndex        =   14
      Top             =   600
      Width           =   3495
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "    „·“Ê„«                                      "
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   21.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Index           =   0
         Left            =   240
         TabIndex        =   17
         Top             =   0
         Width           =   2415
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   615
      Left            =   -120
      TabIndex        =   10
      Top             =   0
      Width           =   10455
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "«‰»«— «À«ÀÌÂ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4320
         TabIndex        =   11
         Top             =   0
         Width           =   1935
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   135
      Left            =   -600
      TabIndex        =   1
      Top             =   6600
      Width           =   11055
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Height          =   15
         Left            =   5115
         TabIndex        =   2
         Top             =   -60
         Width           =   5580
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   -600
      TabIndex        =   0
      Top             =   6600
      Width           =   11655
      Begin Project1.Button Button1 
         Height          =   420
         Left            =   9480
         TabIndex        =   12
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   741
         Caption         =   "Œ—ÊÃ"
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "GhayeshSoft Anbar Maneger 1.0.2.05                    Copyright 2008 All Right Reserved"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   600
         TabIndex        =   9
         Top             =   120
         Width           =   4455
      End
   End
   Begin Project1.Buttonl Button 
      Height          =   375
      Index           =   0
      Left            =   8640
      TabIndex        =   21
      Top             =   1380
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "„‰ÊÌ «’·Ì"
   End
   Begin Project1.Buttonl Button 
      Height          =   375
      Index           =   5
      Left            =   8640
      TabIndex        =   22
      Top             =   2040
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "«›“Êœ‰ ‰„Ê‰Â"
   End
   Begin Project1.Buttonl Button 
      Height          =   375
      Index           =   2
      Left            =   8640
      TabIndex        =   23
      Top             =   2400
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "Ê—Êœ"
   End
   Begin Project1.Buttonl Button 
      Height          =   375
      Index           =   1
      Left            =   8640
      TabIndex        =   24
      Top             =   1680
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "Ã” ÃÊ"
   End
   Begin Project1.Buttonl Button 
      Height          =   375
      Index           =   3
      Left            =   8640
      TabIndex        =   25
      Top             =   2760
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "’«œ—"
   End
   Begin Project1.Buttonl Button 
      Height          =   375
      Index           =   4
      Left            =   8640
      TabIndex        =   26
      Top             =   3120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "ê“«—‘ „ÊÃÊœÌ"
   End
   Begin Project1.Buttonl Button 
      Height          =   375
      Index           =   6
      Left            =   8640
      TabIndex        =   27
      Top             =   3480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "„ÿ»Ê⁄« "
   End
   Begin Project1.Buttonl Button 
      Height          =   375
      Index           =   7
      Left            =   8640
      TabIndex        =   28
      Top             =   3840
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "„·“Ê„« "
   End
   Begin VB.Image Image2 
      Height          =   5760
      Left            =   -120
      Picture         =   "Form1.frx":628A
      Stretch         =   -1  'True
      Top             =   960
      Width           =   10680
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Label5"
      Height          =   495
      Left            =   3120
      TabIndex        =   18
      Top             =   600
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   6
      Left            =   4200
      Picture         =   "Form1.frx":90CD
      Stretch         =   -1  'True
      Top             =   4320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Caption         =   "»—«Ì œ” —”Ì »Â ›«Ì·Â«Ì „·“Ê„«  «” ›«œÂ „Ì ‘Êœ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   5280
      TabIndex        =   13
      Top             =   4320
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Caption         =   "«Ì‰ ê“Ì‰Â »—«Ì «›“Êœ‰  ‰„Ê‰Â ﬂ«·«Ì ÃœÌœ »Â ﬂ«·«Â«Ì „ÊÃÊœ «” ›«œÂ „Ì ‘Êœ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   5280
      TabIndex        =   8
      Top             =   2040
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Caption         =   "»—«Ì œ” —”Ì »Â ›«Ì·Â«Ì „ÿ»Ê⁄«  «” ›«œÂ „Ì ‘Êœ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   5280
      TabIndex        =   7
      Top             =   3840
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Caption         =   "«Ì‰ ê“Ì‰Â »—«Ì ê“«—‘ êÌ—Ì «“ „ÊÃÊœÌ «‰»«— Ê ê—œ‘ ﬂ«·« «” ›«œÂ „Ì ‘Êœ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   5280
      TabIndex        =   6
      Top             =   3360
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Caption         =   "«Ì‰ ê“Ì‰Â »—«Ì ’œÊ— ﬂ«·« «“ «‰»«— »Â ‘⁄»Â Â«Ì “Ì— „Ã„Ê⁄Â «” ›«œÂ „Ì ‘Êœ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   5280
      TabIndex        =   5
      Top             =   2880
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Caption         =   "«Ì‰ ê“Ì‰Â »—«Ì À»  ﬂ«·«Â«Ì Ê—ÊœÌ »Â «‰»«— «” ›«œÂ „Ì ‘Êœ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   5280
      TabIndex        =   4
      Top             =   2400
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Caption         =   "«Ì‰ ê“Ì‰Â »—«Ì Ê—Êœ Ì« ’«œ— ﬂ«·« «“ «‰»«— »Â ﬂ«— „Ì —Êœ Ê À»  ÕÊ«·Â Â« Â„ œ— «Ì‰ Õ«·  «‰Ã«„ „Ì ‘Êœ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   5280
      TabIndex        =   3
      Top             =   1560
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   855
      Index           =   5
      Left            =   4200
      Picture         =   "Form1.frx":B425
      Stretch         =   -1  'True
      Top             =   2040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   4
      Left            =   4200
      Picture         =   "Form1.frx":B976
      Stretch         =   -1  'True
      Top             =   3840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   3
      Left            =   4200
      Picture         =   "Form1.frx":BF4B
      Stretch         =   -1  'True
      Top             =   3240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   855
      Index           =   2
      Left            =   4200
      Picture         =   "Form1.frx":DAF6
      Stretch         =   -1  'True
      Top             =   2640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   855
      Index           =   1
      Left            =   4200
      Picture         =   "Form1.frx":E4F8
      Stretch         =   -1  'True
      Top             =   1680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   855
      Index           =   0
      Left            =   4200
      Picture         =   "Form1.frx":F129
      Stretch         =   -1  'True
      Top             =   1440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image ImgTemp3 
      Height          =   345
      Left            =   5040
      Picture         =   "Form1.frx":FCBF
      Top             =   2520
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Image ImgTemp2 
      Height          =   345
      Left            =   5160
      Picture         =   "Form1.frx":12BBA
      Top             =   2040
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Image imgTemp1 
      Height          =   345
      Left            =   4920
      Picture         =   "Form1.frx":159F9
      Top             =   3000
      Visible         =   0   'False
      Width           =   1725
   End
End
Attribute VB_Name = "Havaleh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i%, dr As Boolean

Private Sub Button_Click(Index As Integer)
Select Case Index

       Case 1: Form2.Changedbass "frmhavale", "frmkala": LodCombo data1: Form2.Show 1: frmKala.Show
       Case 2: frmHavaleh.txtFields(3).Enabled = False: frmHavaleh.txtFields(2).Enabled = True: frmHavaleh.Show: frmHavaleh.data1.Refresh
       Case 3: frmHavaleh.txtFields(2).Enabled = False: frmHavaleh.txtFields(3).Enabled = True: frmHavaleh.Show: frmHavaleh.data1.Refresh
       Case 6: Havaleh.Hide: matmain.Show
       Case 4: Form1.Adodc2.Refresh: Form1.Adodc2.Refresh: Form1.Show 1
       Case 5: frmKala.Show
       Case 7: Havaleh.Hide: MALMAIN.Show
End Select
dr = Not dr
Timer3.Enabled = True

End Sub

Private Sub Timer3_Timer()
If dr = False Then
For i = 0 To 7
   If Button(i).Top <= 1380 Then
   Button(i).Top = 1380
   Else
   Button(i).Top = Button(i).Top - 50
   End If
Next
'Timer1.Enabled = False
Else
For i = 0 To 7
   If Button(i).Top >= (i * 375) + 1380 Then
   Button(i).Top = (i * 375) + 1380
   Else
   Button(i).Top = Button(i).Top + 50
   End If
Next
'Timer1.Enabled = False
End If
End Sub
'FIXIT: Use Option Explicit to avoid implicitly creating variables of type Variant         FixIT90210ae-R383-H1984
Private Sub Button1_Click()
On Error GoTo 4
Unload Me
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" button1_click() Of havaleh,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub

Private Sub Button4_Click()
On Error GoTo 4
Frame6.Visible = False
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" button4_click() Of havaleh,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub

Private Sub Command1_Click()
On Error GoTo 4
frmtayid.Show
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" command1_click() Of havaleh,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub

Private Sub File1_Click()
On Error GoTo 4
        asd.CopyFile File1.List(File1.ListIndex), App.Path & "\"
        form11.Show
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" file1_click() Of havaleh,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub

Private Sub Form_Load()
On Error GoTo 4
data1.ConnectionString = Kalmdb
Me.Left = (Screen.Width \ 2) - (Me.Width \ 2)
Me.Top = (Screen.Height \ 2) - (Me.Height \ 2)
If asd.FolderExists(App.Path + "\Bakup") = False Then asd.CreateFolder App.Path + "\bakup"

Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" form_load() Of havaleh,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo 4
main.Show
main.WindowState = 0
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" form_unload() Of havaleh,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo 4
Button1.Refrash
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" frame2_MouseMove() Of havaleh,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub

Private Sub Button_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
For i = 0 To 7
If i <> Index Then _
Me.Button(i).Refrash
Next
End Sub

Private Sub Mmov(Index%)
On Error GoTo 4
For i = 0 To 6
Label(i).Visible = False
Image1(i).Visible = False
Next
Label(Index).Visible = True
Image1(Index).Visible = True
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" Mmov(" & Index & ") Of havaleh,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub

Private Sub Frame6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo 4
Button4.Refrash
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" Frame6_MouseMove() Of havaleh,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub

Private Sub Img_Click(Index As Integer)
Select Case Index
       Case 0: Form2.Changedbass "frmhavale", "frmkala": LodCombo data1: Form2.Show 1
       Case 1: frmHavaleh.txtFields(3).Enabled = False: frmHavaleh.txtFields(2).Enabled = True: frmHavaleh.Show: frmHavaleh.data1.Refresh
       Case 2: frmHavaleh.txtFields(2).Enabled = False: frmHavaleh.txtFields(3).Enabled = True: frmHavaleh.Show: frmHavaleh.data1.Refresh
       Case 4: Havaleh.Hide: matmain.Show
       Case 3: Form1.Adodc2.Refresh: Form1.Adodc2.Refresh: Form1.Show 1
       Case 5: frmKala.Show 1: frmKala.data1.Recordset.AddNew
       Case 6: Havaleh.Hide: MALMAIN.Show
End Select
End Sub

Private Sub Img_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Mmov Index
End Sub

Private Sub Label4_Click(Index As Integer)
Select Case (Index)
Case 0:
MALMAIN.Show
Havaleh.Hide
Case 1:
Havaleh.Hide
matmain.Show
End Select
End Sub

Private Sub Label5_Click()
Img_Click (4)
End Sub

Private Sub Label7_Click()
Img_Click (6)
End Sub


