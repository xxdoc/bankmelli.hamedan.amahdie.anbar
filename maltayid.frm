VERSION 5.00
Begin VB.Form maltayid 
   BackColor       =   &H80000011&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "„·“Ê„« "
   ClientHeight    =   1440
   ClientLeft      =   4905
   ClientTop       =   5805
   ClientWidth     =   5910
   Icon            =   "maltayid.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame5 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   615
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8175
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "«‰»«— „·“Ê„« "
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
         Left            =   1920
         TabIndex        =   4
         Top             =   0
         Width           =   3255
      End
   End
   Begin Project1.Button command3 
      Height          =   450
      Left            =   3960
      TabIndex        =   0
      Top             =   720
      Width           =   1600
      _ExtentX        =   2831
      _ExtentY        =   794
      Caption         =   "« „«„ ﬂ«—"
   End
   Begin Project1.Button command2 
      Height          =   450
      Left            =   2160
      TabIndex        =   1
      Top             =   720
      Width           =   1600
      _ExtentX        =   2831
      _ExtentY        =   794
      Caption         =   "ﬂ«·«Ì ÃœÌœ"
   End
   Begin Project1.Button command1 
      Height          =   450
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   1600
      _ExtentX        =   2831
      _ExtentY        =   794
      Caption         =   "ÕÊ«·Â ÃœÌœ"
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   855
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   5895
   End
End
Attribute VB_Name = "maltayid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Use Option Explicit to avoid implicitly creating variables of type Variant         FixIT90210ae-R383-H1984

Private Sub Command1_Click()
On Error GoTo 4
malhavaleh.Show
maltayid.Hide
malhavaleh.data1.Recordset.AddNew
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" command1_click() Of maltayid,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)

End Sub

Private Sub Command2_Click()

malhavaleh.Show
maltayid.Hide
With malhavaleh.data1
.Refresh
.Refresh
 .Recordset.MoveLast
 a = .Recordset(2).Value
 b = .Recordset(1).Value
 c = .Recordset(6).Value
 d = .Recordset(0).Value
 
 .Recordset.AddNew
 malhavaleh.Text1 = d
 malhavaleh.Combo(0).Text = a
 malhavaleh.Combo(1).Text = b
 malhavaleh.txtFields(5).Text = c

End With
End Sub

Private Sub Command3_Click()
On Error GoTo 4
Unload Me
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" command3_click() Of maltayid,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo 4
Command1.Refrash
Command2.Refrash
command3.Refrash
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" form_load() Of maltayid,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub
