VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form malhavaleh 
   BackColor       =   &H80000011&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�������"
   ClientHeight    =   5130
   ClientLeft      =   3735
   ClientTop       =   3120
   ClientWidth     =   7320
   Icon            =   "malhavaleh.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   Begin Project1.Buttonl Command2 
      Height          =   375
      Left            =   2760
      TabIndex        =   23
      Top             =   4390
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "������"
   End
   Begin Project1.Buttonl Command1 
      Height          =   375
      Left            =   3960
      TabIndex        =   24
      Top             =   4390
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "�����"
   End
   Begin MSAdodcLib.Adodc data2 
      Height          =   495
      Left            =   6480
      Top             =   3720
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
      RecordSource    =   "shoab"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   375
      Left            =   0
      Top             =   4800
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
      RecordSource    =   "malhavale"
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
   Begin VB.ComboBox Combo 
      DataField       =   "nameshobe"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   1
      Left            =   3000
      TabIndex        =   3
      Top             =   1920
      Width           =   2295
   End
   Begin VB.ComboBox Combo 
      DataField       =   "idshobe"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   0
      Left            =   3000
      TabIndex        =   2
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   615
      Left            =   -840
      TabIndex        =   21
      Top             =   0
      Width           =   8175
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "����� �������"
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
         Left            =   3720
         TabIndex        =   22
         Top             =   0
         Width           =   3255
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "idhavale"
      DataSource      =   "Data1"
      Height          =   320
      Left            =   3000
      TabIndex        =   1
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox txtfields 
      Alignment       =   1  'Right Justify
      DataField       =   "tarikh"
      DataSource      =   "Data1"
      Height          =   320
      Index           =   5
      Left            =   3000
      TabIndex        =   6
      Top             =   3000
      Width           =   2295
   End
   Begin VB.TextBox txtfields 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      DataField       =   "idkala"
      DataSource      =   "Data1"
      Height          =   320
      Index           =   4
      Left            =   3000
      TabIndex        =   0
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox txtfields 
      Alignment       =   1  'Right Justify
      DataField       =   "sadere"
      DataSource      =   "Data1"
      Height          =   320
      Index           =   3
      Left            =   3000
      TabIndex        =   5
      Top             =   2640
      Width           =   2295
   End
   Begin VB.TextBox txtfields 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      DataField       =   "varede"
      DataSource      =   "Data1"
      Height          =   320
      Index           =   2
      Left            =   3000
      TabIndex        =   4
      Top             =   2280
      Width           =   2295
   End
   Begin Project1.Button Button 
      Height          =   375
      Index           =   5
      Left            =   1920
      TabIndex        =   7
      Top             =   840
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "��� ����"
   End
   Begin Project1.Button Button 
      Height          =   405
      Index           =   0
      Left            =   840
      TabIndex        =   8
      Top             =   3675
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   714
      Caption         =   "�����"
   End
   Begin Project1.Button Button 
      Height          =   405
      Index           =   1
      Left            =   1920
      TabIndex        =   9
      Top             =   3675
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   714
      Caption         =   "���"
   End
   Begin Project1.Button Button 
      Height          =   405
      Index           =   2
      Left            =   3000
      TabIndex        =   10
      Top             =   3675
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   714
      Caption         =   "��������"
   End
   Begin Project1.Button Button 
      Height          =   405
      Index           =   3
      Left            =   4080
      TabIndex        =   11
      Top             =   3675
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   714
      Caption         =   "������"
   End
   Begin Project1.Button Button 
      Height          =   405
      Index           =   4
      Left            =   5160
      TabIndex        =   12
      Top             =   3675
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   714
      Caption         =   "����"
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "����� �����"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   5040
      TabIndex        =   20
      Top             =   1200
      Width           =   1500
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "������ ��� ���� �� ���"
      Height          =   255
      Left            =   1080
      TabIndex        =   19
      Top             =   2640
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "�����"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Index           =   5
      Left            =   4920
      TabIndex        =   18
      Top             =   3000
      Width           =   1500
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      BackStyle       =   0  'Transparent
      Caption         =   "�� ����"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   4
      Left            =   4920
      TabIndex        =   17
      Top             =   840
      Width           =   1500
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "�����"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   3
      Left            =   4920
      TabIndex        =   16
      Top             =   2640
      Width           =   1500
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      BackStyle       =   0  'Transparent
      Caption         =   "�����"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   2
      Left            =   4920
      TabIndex        =   15
      Top             =   2280
      Width           =   1500
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "��� ����"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   1
      Left            =   4920
      TabIndex        =   14
      Top             =   1920
      Width           =   1500
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      BackStyle       =   0  'Transparent
      Caption         =   "�� ����"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   0
      Left            =   4920
      TabIndex        =   13
      Top             =   1575
      Width           =   1500
   End
   Begin VB.Image Image1 
      Height          =   4695
      Left            =   0
      Picture         =   "malhavaleh.frx":0ECA
      Stretch         =   -1  'True
      Top             =   120
      Width           =   7335
   End
End
Attribute VB_Name = "malhavaleh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Use Option Explicit to avoid implicitly creating variables of type Variant         FixIT90210ae-R383-H1984
Dim update As Boolean
Dim Moj%, Mojtemp%, Stre$


Private Sub Button_Click(Index As Integer)
On Error GoTo 4
Select Case Index
       Case 0:  data1.Recordset.AddNew
       txtFields(5).Text = Date$
       update = False
         For i = 0 To 4
       Button(i).Visible = False
       Next i
       command1.Visible = True
       command2.Visible = True:
       Case 1:
       With malkala.data1
       .Recordset.MoveFirst
       Y = txtFields(4).Text
       For X = 0 To .Recordset.RecordCount - 1
        If txtFields(4).Text = .Recordset(3).Value Then
         malkala.txtFields(2).Text = Val(malkala.txtFields(2).Text) - Val(txtFields(2).Text)
         malkala.txtFields(2).Text = Val(malkala.txtFields(2).Text) + Val(txtFields(3).Text)
         .Refresh
         If malkala.txtFields(2).Text <> Y Then
           data1.Recordset.AddNew
         End If
         Exit For
        End If
       Next
       End With
       data1.Recordset.Delete
       data1.Refresh
       Case 2:  data1.Refresh
       Case 3:
             With malkala.data1
      .Recordset.update
        malkala.txtFields(2).Text = Val(malkala.txtFields(2).Text) - Val(txtFields(2).Text)
        malkala.txtFields(2).Text = Val(malkala.txtFields(2).Text) + Val(txtFields(3).Text)
        .Recordset.Save
        .Recordset.Close
        .Recordset.Open
       For i = 0 To 4
       Button(i).Visible = False
       Next i
       command1.Visible = True
       command2.Visible = True:
       End With
       update = True
    data1.Recordset.update
                With malkala.data1
            Y = .Recordset.RecordCount
            .Refresh
            .Recordset.MoveFirst
            For i = 0 To Y - 1
            If .Recordset(3).Value = txtFields(4).Text Then
            Moj = .Recordset(2).Value: Button(5).Visible = False
            Stre = "��� ����:" & .Recordset(0).Value & " , ������:"
            Me.Caption = Stre & Moj
            Exit Sub
            End If
            .Recordset.MoveNext
            Next
            Stre = "��� ��� ���� ���� ���"
            Me.Caption = Stre: Button(5).Visible = True
End With
       Case 4:  Unload Me
       Case 5:  malkala.Show 1
                malkala.data1.Recordset.AddNew
                malkala.txtFields(3).Text = txtFields(4).Text
End Select
txtFields(5).Text = IRDate2()
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" button_click(" & Index & ") Of MalHavaleh,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub

'---for mojody ,for mojodytemp , for title of form
Private Sub cmdAdd_Click()
On Error GoTo 4
  data1.Recordset.AddNew
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" cmdadd() Of MalHavaleh,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub

Private Sub cmdDelete_Click()
On Error GoTo 4
  'this may produce an error if you delete the last
  'record or the only record in the recordset
  data1.Recordset.Delete
  data1.Recordset.MoveNext
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" cmddelet() Of MalHavaleh,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub

Private Sub cmdRefresh_Click()
On Error GoTo 4
  'this is really only needed for multi user apps
  data1.Refresh
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" cmdresrash() Of MalHavaleh,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub



Private Sub cmdClose_Click()
On Error GoTo 4
  Unload Me
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" cmdclose() Of MalHavaleh,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub




Private Sub Combo_Click(Index As Integer)
Select Case (Index)
Case 0:
With data2
 .Recordset.MoveLast
   Y = .Recordset.RecordCount - 1
  .Refresh
  .Recordset.MoveFirst
        For i = 0 To Y
            If .Recordset(1).Value = Combo(0).Text Then
                 Combo(1).Text = .Recordset(0).Value
                 Exit Sub
            End If
              .Recordset.MoveNext
            Next
End With
Case 1:
With data2
 .Recordset.MoveLast
   Y = .Recordset.RecordCount - 1
  .Refresh
  .Recordset.MoveFirst
        For i = 0 To Y
            If .Recordset(0).Value = Combo(1).Text Then
                 Combo(0).Text = .Recordset(1).Value
                 Exit Sub
            End If
              .Recordset.MoveNext
            Next
End With
End Select
End Sub

Private Sub Combo_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
  If Index = 1 Then
    If txtFields(2).Enabled Then txtFields(2).SetFocus
    If txtFields(3).Enabled Then txtFields(3).SetFocus
  End If
 If Index = 0 Then Combo(1).SetFocus
End If
End Sub

Private Sub Combo_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case (Index)
Case 0:
With data2
 .Recordset.MoveLast
   Y = .Recordset.RecordCount - 1
  .Refresh
  .Recordset.MoveFirst
        For i = 0 To Y
            If .Recordset(1).Value = Combo(0).Text Then
                 Combo(1).Text = .Recordset(0).Value
                 Exit Sub
            End If
              .Recordset.MoveNext
            Next
End With
Case 1:
With data2
 .Recordset.MoveLast
   Y = .Recordset.RecordCount - 1
  .Refresh
  .Recordset.MoveFirst
        For i = 0 To Y
            If .Recordset(0).Value = Combo(1).Text Then
                 Combo(0).Text = .Recordset(1).Value
                 Exit Sub
            End If
              .Recordset.MoveNext
            Next
End With
End Select
End Sub

Private Sub Command1_Click()
  With malkala.data1
 Y = malkala.txtFields(2).Text
   .Recordset.update
   malkala.txtFields(2).Text = Val(malkala.txtFields(2).Text) + Val(txtFields(2).Text)
   malkala.txtFields(2).Text = Val(malkala.txtFields(2).Text) - Val(txtFields(3).Text)
   .Recordset.Save
   data1.Recordset.Save
   .Recordset.Close
   .Refresh
    If malkala.txtFields(2).Text <> Y Then
       If txtFields(2).Text = "" Then txtFields(2).Text = 0
       If txtFields(3).Text = "" Then txtFields(3).Text = 0
       If update = True Then
       MsgBox "taghirat emal shod"
       command1.Visible = False
       command2.Visible = False
       For i = 0 To 4
       Button(i).Visible = True
       Next i
       End If
       If update = False Then malhavaleh.Hide: maltayid.Show
    End If
End With
data1.Recordset.Save
data1.Refresh
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
4 Call Loger(" data_resposition() Of MalHavaleh,Src:" & Err.Source & " ,Num:" _
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
4 Call Loger(" data1_validation() Of MalHavaleh,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub







Private Sub Command2_Click()
data1.Recordset.CancelUpdate
command1.Visible = False
command2.Visible = False
For i = 0 To 4
Button(i).Visible = True
Next i
End Sub

Private Sub Form_Load()
With malhavaleh.data2
 .Refresh
 .Recordset.MoveLast
 Y = .Recordset.RecordCount - 1
 .Recordset.MoveFirst
 For i = 0 To Y
  Combo(0).AddItem .Recordset(1).Value
  Combo(1).AddItem .Recordset(0).Value
  .Recordset.MoveNext
 Next i
End With

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo 4
For i = 0 To 5
Me.Button(i).Refrash
Next
command1.Refrash
command2.Refrash
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" form_MouseMove() Of MalHavaleh,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub


Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo 4
Form_MouseMove Button, Shift, X, Y
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" image1_MouseMove() Of MalHavaleh,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub



Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Combo(0).SetFocus
End If
End Sub

Private Sub txtFields_Change(Index As Integer)
On Error GoTo 4
Select Case Index
Case 0
With data2
 .Recordset.MoveLast
   Y = .Recordset.RecordCount - 1
  .Refresh
  .Recordset.MoveFirst
        For i = 0 To Y
            If .Recordset(1).Value = txtFields(0).Text Then
                 txtFields(1).Text = .Recordset(0).Value
                 Exit Sub
            End If
              .Recordset.MoveNext
            Next
End With
Case 4 '-------kod kala
            With malkala.data1
            Y = .Recordset.RecordCount
            .Refresh
            .Recordset.MoveFirst
            For i = 0 To Y - 1
            If .Recordset(3).Value = txtFields(4).Text Then
            Moj = .Recordset(2).Value: Button(5).Visible = False
            Stre = "��� ����:" & .Recordset(0).Value & " , ������:"
            Me.Caption = Stre & Moj
            Exit Sub
            End If
            .Recordset.MoveNext
            Next
            Stre = "��� ��� ���� ���� ���"
            Me.Caption = Stre: Button(5).Visible = True
End With
Case 2 '--------varedeh
Me.Caption = Stre & (Moj + Val(txtFields(2).Text))
Case 3 '--------sadereh
Me.Caption = Stre & (Moj - Val(txtFields(3).Text))
End Select
txtFields(3).BackColor = IIf(((Moj - Val(txtFields(3).Text)) < 0), vbYellow, vbWhite)
Label1.Visible = IIf(((Moj - Val(txtFields(3).Text)) < 0), True, False)
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" txtFields_change(" & Index & ") Of MalHavaleh,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)

End Sub

Private Sub txtFields_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
  If Index = 2 Then txtFields(5).SetFocus
  If Index = 3 Then txtFields(5).SetFocus
  If Index = 4 Then Text1.SetFocus
  If Index = 5 Then command1.SetFocus
 End If
End Sub
