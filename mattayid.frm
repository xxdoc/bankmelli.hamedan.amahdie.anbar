VERSION 5.00
Begin VB.Form mattayid 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "„ÿ»Ê⁄« "
   ClientHeight    =   1500
   ClientLeft      =   5100
   ClientTop       =   5805
   ClientWidth     =   5940
   Icon            =   "mattayid.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame5 
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   10455
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "«‰»«— „ÿ»Ê⁄« "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   735
         Left            =   1800
         TabIndex        =   4
         Top             =   0
         Width           =   3255
      End
   End
   Begin Project1.Button command3 
      Height          =   450
      Left            =   3960
      TabIndex        =   2
      Top             =   840
      Width           =   1600
      _ExtentX        =   2831
      _ExtentY        =   794
      Caption         =   "« „«„ ﬂ«—"
   End
   Begin Project1.Button command2 
      Height          =   450
      Left            =   2160
      TabIndex        =   1
      Top             =   840
      Width           =   1600
      _ExtentX        =   2831
      _ExtentY        =   794
      Caption         =   "ﬂ«·«Ì ÃœÌœ"
   End
   Begin Project1.Button command1 
      Height          =   450
      Left            =   240
      TabIndex        =   0
      Top             =   840
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
Attribute VB_Name = "mattayid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'FIXIT: Use Option Explicit to avoid implicitly creating variables of type Variant         FixIT90210ae-R383-H1984
Private Sub Command1_Click()
On Error GoTo 4
mathavaleh.Show
mattayid.Hide
mathavaleh.data1.Recordset.AddNew
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger("Command1() Of Mattayid,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)

End Sub

Private Sub Command2_Click()
On Error GoTo 4
mathavaleh.Show
mattayid.Hide
With mathavaleh.data1
.Refresh
.Refresh
 .Recordset.MoveLast
 a = .Recordset(2).Value
 b = .Recordset(1).Value
 c = .Recordset(6).Value
 d = .Recordset(0).Value
 
 .Recordset.AddNew
 mathavaleh.Text1 = d
 mathavaleh.Combo(0).Text = a
 mathavaleh.Combo(1).Text = b
 mathavaleh.txtFields(5).Text = c
End With
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger("Command2_Click() Of Mattayid,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)

End Sub

Private Sub Command3_Click()
On Error GoTo 4
Unload Me
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger("Command3_Click() Of Mattayid,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo 4
Command1.Refrash: Command2.Refrash: command3.Refrash
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger("Form_MousMove() Of Mattayid,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)

End Sub
