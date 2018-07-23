VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmKala 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Kala"
   ClientHeight    =   5070
   ClientLeft      =   4320
   ClientTop       =   3120
   ClientWidth     =   7350
   Icon            =   "frmKala.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   Begin Project1.Buttonl command1 
      Height          =   420
      Left            =   2640
      TabIndex        =   19
      Top             =   4200
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   741
      Caption         =   " «ÌÌœ"
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   375
      Left            =   0
      Top             =   4680
      Width           =   7335
      _ExtentX        =   12938
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
      RecordSource    =   "Kala"
      Caption         =   ""
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
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   615
      Left            =   -120
      TabIndex        =   17
      Top             =   0
      Width           =   7455
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
         Left            =   2640
         TabIndex        =   18
         Top             =   0
         Width           =   1935
      End
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "namekala"
      DataSource      =   "Data1"
      Height          =   350
      Index           =   0
      Left            =   2640
      MaxLength       =   50
      TabIndex        =   0
      Top             =   720
      Width           =   2775
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "vahedkala"
      DataSource      =   "Data1"
      Height          =   350
      Index           =   1
      Left            =   2640
      MaxLength       =   20
      TabIndex        =   3
      Top             =   2160
      Width           =   2775
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "mojodi"
      DataSource      =   "Data1"
      Height          =   350
      Index           =   2
      Left            =   3480
      TabIndex        =   2
      Top             =   1680
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "idkala"
      DataSource      =   "Data1"
      Height          =   350
      Index           =   3
      Left            =   3480
      TabIndex        =   1
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "seriyal"
      DataSource      =   "Data1"
      Height          =   350
      Index           =   4
      Left            =   2640
      MaxLength       =   25
      TabIndex        =   4
      Top             =   2640
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "tozohat"
      DataSource      =   "Data1"
      Height          =   350
      Left            =   2640
      TabIndex        =   5
      Top             =   3120
      Width           =   2775
   End
   Begin Project1.Button Button 
      Height          =   405
      Index           =   0
      Left            =   1320
      TabIndex        =   6
      Top             =   3645
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   714
      Caption         =   "«÷«›Â"
   End
   Begin Project1.Button Button 
      Height          =   405
      Index           =   1
      Left            =   2400
      TabIndex        =   7
      Top             =   3645
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   714
      Caption         =   "Õ–›"
   End
   Begin Project1.Button Button 
      Height          =   405
      Index           =   2
      Left            =   3480
      TabIndex        =   8
      Top             =   3645
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   714
      Caption         =   "»—Ê“ﬂ—œ‰"
   End
   Begin Project1.Button Button 
      Height          =   405
      Index           =   3
      Left            =   4560
      TabIndex        =   9
      Top             =   3645
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   714
      Caption         =   "ÊÌ—«Ì‘"
   End
   Begin Project1.Button Button 
      Height          =   405
      Index           =   4
      Left            =   5640
      TabIndex        =   10
      Top             =   3645
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   714
      Caption         =   "»” ‰"
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      BackColor       =   &H80000016&
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ ﬂ«·«"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   5520
      TabIndex        =   16
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ê«Õœ ﬂ«·«"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   5520
      TabIndex        =   15
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      BackColor       =   &H80000016&
      BackStyle       =   0  'Transparent
      Caption         =   "„ÊÃÊœÌ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   5640
      TabIndex        =   14
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ﬂœ ﬂ«·«"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   5520
      TabIndex        =   13
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      BackColor       =   &H80000016&
      BackStyle       =   0  'Transparent
      Caption         =   "”—Ì«·"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   4
      Left            =   5640
      TabIndex        =   12
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " Ê÷ÌÕ« "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5640
      TabIndex        =   11
      Top             =   3120
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   4725
      Left            =   -120
      Picture         =   "frmKala.frx":0ECA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7455
   End
End
Attribute VB_Name = "frmKala"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Adodc1_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

End Sub

'FIXIT: Use Option Explicit to avoid implicitly creating variables of type Variant         FixIT90210ae-R383-H1984
Private Sub Button_Click(Index As Integer)
On Error GoTo 4
Select Case Index
       Case 0:  data1.Recordset.AddNew
       Case 1:  data1.Recordset.Delete
       Case 2:  data1.Refresh
       Case 3:  data1.Recordset.update
       Case 4:  Unload Me
End Select
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" button_click(" & Index & ") Of FrmKala,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub


Private Sub cmdDelete_Click()
On Error GoTo 4
  'this may produce an error if you delete the last
  'record or the only record in the recordset
  data1.Recordset.Delete
  data1.Recordset.MoveNext
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" cmdDelet_click() Of FrmKala,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub



Private Sub Data1_Reposition()
On Error GoTo 4
  Screen.MousePointer = vbDefault
  On Error Resume Next
  'This will display the current record position
  'for dynasets and snapshots
  data1.Caption = "Record: " & (data1.Recordset.AbsolutePosition + 1)
  'for the table object you must set the index property when
  'the recordset gets created and use the following line
  'Data1.Caption = "Record: " & (Data1.Recordset.RecordCount * (Data1.Recordset.PercentPosition * 0.01)) + 1
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" data1_resposition() Of FrmKala,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub

Private Sub data1_Validate(Action As Integer, Save As Integer)
On Error GoTo 4
  'This is where you put validation code
  'This event gets called when the following actions occur
  Select Case Action
    Case vbDataActionMoveFirst
    Case vbDataActionMovePrevious
    Case vbDataActionMoveNext
    Case vbDataActionMoveLast
    Case vbDataActionAddNew
    Case vbDataActionUpdate
    Case vbDataActionDelete
    Case vbDataActionFind
    Case vbDataActionBookmark
    Case vbDataActionClose
  End Select
'  Screen.MousePointer = vbHourglass
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" data1_validate() Of FrmKala,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub



Private Sub Command1_Click()
data1.Recordset.Save
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo 4
For i = 0 To 4
Me.Button(i).Refrash
Next
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" form_MouseMove() Of FrmKala,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo 4
Form_MouseMove Button, Shift, X, Y
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" image1_KeyDown(" & Index & ") Of FrmKala,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub

Private Sub Image2_Click()

End Sub

Private Sub txtFields_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo 4
If Index = 3 Then
With Havaleh.data1
.Recordset.MoveFirst
For i = 0 To .Recordset.RecordCount - 1
If .Recordset(3) = txtFields(3).Text Then
txtFields(3).BackColor = vbYellow: Button(0).Visible = False
Me.Caption = "«Ì‰ ﬂœ ﬂ«·« ﬁ»·¬ ÊÃÊœ œ«—œ"
Exit Sub
End If
.Recordset.MoveNext
Next
Me.Caption = "«›“Êœ‰ ‰„Ê‰Â": Button(0).Visible = True
txtFields(3).BackColor = &H80000005
End With
End If
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" txtFields_KeyDown(" & Index & ") Of FrmKala,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)

End Sub



