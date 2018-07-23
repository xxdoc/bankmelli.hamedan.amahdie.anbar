VERSION 5.00
Begin VB.Form frmtayid 
   BackColor       =   &H00008000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "‰„Ê‰Â Â«"
   ClientHeight    =   1365
   ClientLeft      =   3930
   ClientTop       =   5430
   ClientWidth     =   7530
   Icon            =   "frmtayid.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   7530
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   615
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   7575
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "«‰»«— ‰„Ê‰Â"
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
         Left            =   3000
         TabIndex        =   4
         Top             =   0
         Width           =   1935
      End
   End
   Begin Project1.Button command3 
      Height          =   375
      Left            =   5160
      TabIndex        =   0
      Top             =   840
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Caption         =   "« „«„ ﬂ«—"
   End
   Begin Project1.Button command2 
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   840
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "ﬂ«·«Ì ÃœÌœ"
   End
   Begin Project1.Button command1 
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   840
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "ÕÊ«·Â ÃœÌœ"
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   855
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   7575
   End
End
Attribute VB_Name = "frmtayid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Use Option Explicit to avoid implicitly creating variables of type Variant         FixIT90210ae-R383-H1984

Private Sub Command1_Click()
For X = 2 To 5
frmHavaleh.txtFields(X).Locked = False
Next X
frmHavaleh.Text1.Locked = False
frmHavaleh.Combo(0).Locked = False
frmHavaleh.Combo(1).Locked = False
On Error GoTo 4
frmHavaleh.Show
frmtayid.Hide
frmHavaleh.data1.Recordset.AddNew
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" command1_click() Of frmtayid,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)

End Sub

Private Sub Command2_Click()
For X = 2 To 5
 frmHavaleh.txtFields(X).Locked = False
 Next X
frmHavaleh.Text1.Locked = False
frmHavaleh.Combo(0).Locked = False
frmHavaleh.Combo(1).Locked = False
frmHavaleh.Show
frmtayid.Hide
With frmHavaleh.data1
.Refresh
.Refresh

 .Recordset.MoveLast
 a = .Recordset(2).Value
 b = .Recordset(1).Value
 c = .Recordset(6).Value
 d = .Recordset(0).Value
 .Recordset.AddNew
 frmHavaleh.Text1 = d
 frmHavaleh.Combo(0).Text = a
 frmHavaleh.Combo(1).Text = b
 frmHavaleh.txtFields(5).Text = c
End With
End Sub

Private Sub Command3_Click()
On Error GoTo 4
mathavaleh.txtFields(2).Enabled = True
mathavaleh.txtFields(3).Enabled = True
Unload frmHavaleh
Unload Me

Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" command3_click() Of frmtayid,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.Refrash
Command2.Refrash
command3.Refrash
End Sub
