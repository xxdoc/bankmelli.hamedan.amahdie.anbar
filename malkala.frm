VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form malkala 
   BackColor       =   &H80000011&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "„·“Ê„« "
   ClientHeight    =   5475
   ClientLeft      =   4125
   ClientTop       =   3315
   ClientWidth     =   7440
   Icon            =   "malkala.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   Begin Project1.Buttonl command1 
      Height          =   375
      Left            =   2280
      TabIndex        =   19
      Top             =   4770
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      Caption         =   " «ÌÌœ"
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   375
      Left            =   0
      Top             =   5160
      Width           =   7455
      _ExtentX        =   13150
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
      RecordSource    =   "malkala"
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
   Begin VB.Frame Frame5 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   615
      Left            =   0
      TabIndex        =   17
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
         Left            =   2640
         TabIndex        =   18
         Top             =   0
         Width           =   3255
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "tozohat"
      DataSource      =   "Data1"
      Height          =   340
      Left            =   1800
      TabIndex        =   5
      Top             =   3480
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "seriyal"
      DataSource      =   "Data1"
      Height          =   340
      Index           =   4
      Left            =   1800
      MaxLength       =   25
      TabIndex        =   4
      Top             =   3000
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "idkala"
      DataSource      =   "Data1"
      Height          =   340
      Index           =   3
      Left            =   3600
      TabIndex        =   1
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "mojodi"
      DataSource      =   "Data1"
      Height          =   340
      Index           =   2
      Left            =   3600
      TabIndex        =   3
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "vahedkala"
      DataSource      =   "Data1"
      Height          =   340
      Index           =   1
      Left            =   3600
      MaxLength       =   20
      TabIndex        =   2
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "namekala"
      DataSource      =   "Data1"
      Height          =   340
      Index           =   0
      Left            =   2400
      MaxLength       =   50
      TabIndex        =   0
      Top             =   1080
      Width           =   2775
   End
   Begin Project1.Button Button 
      Height          =   375
      Index           =   0
      Left            =   840
      TabIndex        =   6
      Top             =   4035
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      Caption         =   "«÷«›Â"
   End
   Begin Project1.Button Button 
      Height          =   375
      Index           =   1
      Left            =   1920
      TabIndex        =   7
      Top             =   4035
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      Caption         =   "Õ–›"
   End
   Begin Project1.Button Button 
      Height          =   375
      Index           =   2
      Left            =   3000
      TabIndex        =   8
      Top             =   4035
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      Caption         =   "»—Ê“ﬂ—œ‰"
   End
   Begin Project1.Button Button 
      Height          =   375
      Index           =   3
      Left            =   4080
      TabIndex        =   9
      Top             =   4035
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      Caption         =   "ÊÌ—«Ì‘"
   End
   Begin Project1.Button Button 
      Height          =   375
      Index           =   4
      Left            =   5280
      TabIndex        =   10
      Top             =   4035
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      Caption         =   "»” ‰"
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
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
      Height          =   315
      Left            =   4320
      TabIndex        =   16
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
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
      Height          =   315
      Index           =   4
      Left            =   4320
      TabIndex        =   15
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
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
      Height          =   315
      Index           =   3
      Left            =   4320
      TabIndex        =   14
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
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
      Height          =   315
      Index           =   2
      Left            =   4320
      TabIndex        =   13
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
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
      Height          =   315
      Index           =   1
      Left            =   4440
      TabIndex        =   12
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
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
      Height          =   315
      Index           =   0
      Left            =   4320
      TabIndex        =   11
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   4575
      Left            =   0
      Picture         =   "malkala.frx":0ECA
      Stretch         =   -1  'True
      Top             =   600
      Width           =   7455
   End
End
Attribute VB_Name = "malkala"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
4 Call Loger(" button_click(" & Index & ") Of Malkala,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub


Private Sub cmdDelete_Click()
On Error GoTo 4
  'this may produce an error if you delete the last
  'record or the only record in the recordset
  data1.Recordset.Delete
  data1.Recordset.MoveNext
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" cmdDelet_click() Of Malkala,Src:" & Err.Source & " ,Num:" _
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
4 Call Loger(" data1_resposition() Of Malkala,Src:" & Err.Source & " ,Num:" _
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
4 Call Loger(" data1_validate() Of Malkala,Src:" & Err.Source & " ,Num:" _
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
command1.Refrash
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" form_MouseMove() Of Malkala,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo 4
Form_MouseMove Button, Shift, X, Y
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" image1_MouseMove() Of Malkala,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub

Private Sub txtFields_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo 4
If Index = 3 Then
With MALMAIN.data1
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
4 Call Loger(" txtFields_KeyDown(" & Index & ") Of Malkala,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)

End Sub



