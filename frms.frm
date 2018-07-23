VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frms 
   BackColor       =   &H00FFFFFF&
   Caption         =   "«À«ÀÌÂ"
   ClientHeight    =   8280
   ClientLeft      =   3015
   ClientTop       =   2595
   ClientWidth     =   10140
   LinkTopic       =   "Form3"
   ScaleHeight     =   8280
   ScaleWidth      =   10140
   Begin MSFlexGridLib.MSFlexGrid msf 
      Bindings        =   "frms.frx":0000
      Height          =   5895
      Left            =   0
      TabIndex        =   7
      Top             =   720
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   10398
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      BackColorBkg    =   -2147483636
      ScrollTrack     =   -1  'True
      RightToLeft     =   -1  'True
      AllowUserResizing=   3
      BorderStyle     =   0
      Appearance      =   0
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   7440
      Top             =   7680
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
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
      RecordSource    =   "havale"
      Caption         =   "‰«„ œ«œÂ"
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
   Begin VB.TextBox Text1 
      DataField       =   "tarikh"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   405
      Index           =   6
      Left            =   8760
      TabIndex        =   13
      Top             =   7080
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      DataField       =   "idkala"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   405
      Index           =   5
      Left            =   7320
      TabIndex        =   12
      Top             =   7080
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      DataField       =   "sadere"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   405
      Index           =   4
      Left            =   5880
      TabIndex        =   11
      Top             =   7080
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      DataField       =   "varede"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   405
      Index           =   3
      Left            =   4440
      TabIndex        =   10
      Top             =   7080
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      DataField       =   "nameshobe"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   405
      Index           =   2
      Left            =   3000
      TabIndex        =   9
      Top             =   7080
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      DataField       =   "idshobe"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   405
      Index           =   1
      Left            =   1560
      TabIndex        =   8
      Top             =   7080
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      DataField       =   "idhavale"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   405
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   7080
      Width           =   1335
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   855
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10095
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
         Left            =   3480
         TabIndex        =   3
         Top             =   0
         Width           =   2295
      End
   End
   Begin Project1.Button Img 
      CausesValidation=   0   'False
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   7680
      Width           =   1695
      _ExtentX        =   3201
      _ExtentY        =   873
      Caption         =   "Å—Ì‰ "
   End
   Begin Project1.Button Img 
      Height          =   375
      Index           =   3
      Left            =   1920
      TabIndex        =   1
      Top             =   7680
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Caption         =   "»—ê‘ "
   End
   Begin Project1.Button Img 
      CausesValidation=   0   'False
      Height          =   375
      Index           =   1
      Left            =   3600
      TabIndex        =   5
      Top             =   7680
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Caption         =   "ÊÌ—«Ì‘"
   End
   Begin Project1.Button Img 
      CausesValidation=   0   'False
      Height          =   375
      Index           =   2
      Left            =   5400
      TabIndex        =   6
      Top             =   7680
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Caption         =   "Õ–›"
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " «—ÌŒ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8760
      TabIndex        =   20
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ﬂœ ﬂ«·«"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   19
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "’«œ—Â"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   18
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ê«—œÂ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   17
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ ‘⁄»Â"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   16
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ﬂœ ‘⁄»Â"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   15
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "‘„«—Â ÕÊ«·Â"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   6600
      Width           =   1215
   End
End
Attribute VB_Name = "frms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim update As Boolean

Private Sub Form_Activate()
update = False
frmKala.Hide
End Sub

Private Sub Form_Load()
Me.BackColor = RGB(230, 255, 230)
frmKala.Hide
msf.AllowBigSelection = True
msf.Row = 0
msf.Col = 1
msf.Text = "‘„«—Â ÕÊ«·Â"
msf.Col = 2
msf.Text = "ﬂœ ‘⁄»Â"
msf.Col = 3
msf.Text = "‰«„ Ê«Õœ"
msf.Col = 4
msf.Text = "Ê«—œÂ"
msf.Col = 5
msf.Text = "’«œ—Â"
msf.Col = 6
msf.Text = "ﬂœ ﬂ«·«"
msf.Col = 7
msf.Text = " «—ÌŒ"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
For i = 0 To 3
Img(i).Refrash
Next
End Sub

Private Sub Ima_Click(Index As Integer)
 

frms.Hide
Form2.Show
For i = 1 To msf.Row - 1
msf.RemoveItem 2
Next i
End Sub

Private Sub Img_Click(Index As Integer)
Dim ii%, n%
On Error GoTo 4
Select Case (Index)
Case 0:
       Me.PrintForm
Case 1:
  If update = False Then
    For i = 0 To 6
       Text1(i).Enabled = True
    Next i
    Adodc1.Recordset.update
    With frmKala.data1
         .Recordset.update
          frmKala.txtFields(2).Text = Val(frmKala.txtFields(2).Text) - Val(Text1(3).Text)
          frmKala.txtFields(2).Text = Val(frmKala.txtFields(2).Text) + Val(Text1(4).Text)
          .Recordset.Save
          .Recordset.update
          Moj = .Recordset(2).Value
          Stre = "‰«„ ﬂ«·«:" & .Recordset(0).Value & " , „ÊÃÊœÌ:"
          Me.Caption = Stre & Moj
    End With
    update = True
    Exit Sub
  End If
 If update = True Then
       For i = 0 To 6
       Text1(i).Enabled = False
    Next i
       With frmKala.data1
         .Recordset.update
          frmKala.txtFields(2).Text = Val(frmKala.txtFields(2).Text) + Val(Text1(3).Text)
          frmKala.txtFields(2).Text = Val(frmKala.txtFields(2).Text) - Val(Text1(4).Text)
          .Recordset.Save
          Moj = .Recordset(2).Value
          Stre = "‰«„ ﬂ«·«:" & .Recordset(0).Value & " , „ÊÃÊœÌ:"
          Me.Caption = Stre & Moj
    End With
    update = False
    Adodc1.Recordset.Save
    Exit Sub
  End If




Case 2:
      With frmKala.data1
       Y = frmKala.txtFields(4).Text
      .Recordset.update
        frmKala.txtFields(2).Text = Val(frmKala.txtFields(2).Text) - Val(Text1(3).Text)
        frmKala.txtFields(2).Text = Val(frmKala.txtFields(2).Text) + Val(Text1(4).Text)
        .Recordset.Save
        .Refresh
         If frmKala.txtFields(2).Text <> Y Then
           MsgBox "hazf shod"
           End If
     
       Adodc1.Recordset.Delete
       Form2.Button2.Refrash
    End With
End Select
Exit Sub '-----{Call Loger In Erroring!}--------
4: MsgBox "œ” ê«Â Å—Ì‰ — ‘„« Ì« Ê’· ‰Ì”  Ì« ¬„«œÂ ‰Ì”  ! ·ÿ›¬ »——”Ì ﬂ‰Ìœ", vbExclamation, "Œÿ«œ— Å—Ì‰ "
Call Loger(" Command1_Click() Of ,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub

Private Sub msf_Click()
Adodc1.Recordset.MoveFirst
Y = msf.Text
msf.BackColorSel = Blue
For i = 0 To msf.Text - 1
Adodc1.Recordset.MoveNext
Next i
With frmKala.data1
 .Recordset.MoveFirst
 For X = 0 To .Recordset.RecordCount - 1
  If .Recordset(3).Value = Adodc1.Recordset(5).Value Then
   Exit Sub
  End If
  .Recordset.MoveNext
 Next X
End With
End Sub

