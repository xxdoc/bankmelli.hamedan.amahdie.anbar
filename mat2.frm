VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form mat2 
   Caption         =   "�������"
   ClientHeight    =   7140
   ClientLeft      =   2970
   ClientTop       =   3525
   ClientWidth     =   10380
   LinkTopic       =   "Form3"
   ScaleHeight     =   7140
   ScaleWidth      =   10380
   Begin MSAdodcLib.Adodc data1 
      Height          =   375
      Left            =   480
      Top             =   6600
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      RecordSource    =   "mathavale"
      Caption         =   "Adodc3"
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
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   615
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   10455
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
         ForeColor       =   &H00000000&
         Height          =   735
         Left            =   3840
         TabIndex        =   29
         Top             =   0
         Width           =   3255
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "����� ���� ���"
      Height          =   5415
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   10335
      Begin VB.TextBox Text 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   6
         Left            =   4080
         TabIndex        =   11
         Top             =   3480
         Width           =   1215
      End
      Begin VB.TextBox Text 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   5
         Left            =   4080
         TabIndex        =   10
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox Text 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   4
         Left            =   4080
         TabIndex        =   9
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox Text 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   3
         Left            =   7320
         TabIndex        =   8
         Top             =   4080
         Width           =   1215
      End
      Begin VB.TextBox Text 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   2
         Left            =   7320
         TabIndex        =   7
         Top             =   3480
         Width           =   1215
      End
      Begin VB.TextBox Text 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   1
         Left            =   7320
         TabIndex        =   6
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox Text 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   0
         Left            =   7320
         TabIndex        =   5
         Top             =   2280
         Width           =   1215
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00FFFFFF&
         Height          =   1425
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   10215
      End
      Begin Project1.Button Button2 
         Height          =   375
         Left            =   1440
         TabIndex        =   12
         Top             =   4080
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         Caption         =   "�����"
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         Caption         =   "����� �����"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   5
         Left            =   5400
         TabIndex        =   27
         Top             =   3480
         Width           =   1500
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         Caption         =   "�����"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   4
         Left            =   5400
         TabIndex        =   26
         Top             =   2880
         Width           =   1500
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         Caption         =   "������"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   3
         Left            =   5400
         TabIndex        =   25
         Top             =   2280
         Width           =   1500
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         Caption         =   "�����"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   2
         Left            =   8640
         TabIndex        =   24
         Top             =   4080
         Width           =   1500
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         Caption         =   "�����"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   11
         Left            =   8640
         TabIndex        =   23
         Top             =   3480
         Width           =   1500
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         Caption         =   "������"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   0
         Left            =   8640
         TabIndex        =   22
         Top             =   2880
         Width           =   1500
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         Caption         =   "��� ����"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   1
         Left            =   8640
         TabIndex        =   21
         Top             =   2280
         Width           =   1500
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "����� �� ���"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8760
         TabIndex        =   20
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         Height          =   3375
         Left            =   120
         Top             =   1920
         Width           =   10215
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "����� �����"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   2040
         TabIndex        =   19
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "�����"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   2040
         TabIndex        =   18
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "������"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   2040
         TabIndex        =   17
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "�����"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   4800
         TabIndex        =   16
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "�����"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   4800
         TabIndex        =   15
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "������"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   4800
         TabIndex        =   14
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "��� ����"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   4800
         TabIndex        =   13
         Top             =   2160
         Width           =   1215
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   5880
      Width           =   10455
      Begin VB.Frame Frame2 
         Height          =   55
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   10455
      End
      Begin Project1.Button Ima 
         Height          =   375
         Index           =   0
         Left            =   9000
         TabIndex        =   2
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "��� ����"
      End
   End
End
Attribute VB_Name = "mat2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Use Option Explicit to avoid implicitly creating variables of type Variant         FixIT90210ae-R383-H1984
Dim Mojode%, Indesk



Private Sub Button2_Click()
Dim Sstr$
If Text(Indesk).Text = "" Then
MsgBox "���� ������ �� ���� ����� �� ��� �� ���� �� ��� ���� ", vbInformation, "�����"
Exit Sub
End If

        mat2.Hide: mats.Show
        data1.Refresh
        data1.Recordset.MoveLast
        Y = data1.Recordset.RecordCount
        data1.Recordset.MoveFirst
        For i = 0 To Y - 1
          If (Text(0).Text = "") Or (Text(0).Text = data1.Recordset(1).Value) Then
           If (Text(1).Text = "") Or (Text(1).Text = data1.Recordset(2).Value) Then
             If (Text(2).Text = "") Or (Text(2).Text = data1.Recordset(3).Value) Then
                If (Text(3).Text = "") Or (Text(3).Text = data1.Recordset(4).Value) Then
                    If (Text(4).Text = "") Or (Text(4).Text = data1.Recordset(5).Value) Then
                        If (Text(5).Text = "") Or (Text(5).Text = data1.Recordset(6).Value) Then
                           If (Text(6).Text = "") Or (Text(6).Text = data1.Recordset(0).Value) Then
          mats.msf.AddItem ""
          mats.msf.Row = mats.msf.Row + 1
          mats.msf.Col = 0
          mats.msf.Text = i
          For X = 1 To 7
           mats.msf.Col = X
           mats.msf.Text = data1.Recordset(X - 1).Value
          Next X
          End If
           End If
            End If
             End If
              End If
               End If
                 End If
           
          data1.Recordset.MoveNext
         Next i
End Sub

Private Sub Combo1_Change()
On Error GoTo 4
If Combo1.ListIndex = -1 Then Exit Sub
Adodc2.Recordset.MoveFirst
Adodc2.Recordset.Move Combo1.ListIndex
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" Combo1_Change() Of mat2,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub

Private Sub Combo1_Click()
On Error GoTo 4
If Combo1.ListIndex = -1 Then Exit Sub
Adodc2.Recordset.MoveFirst
Adodc2.Recordset.Move Combo1.ListIndex
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" Combo1_Click() Of mat2,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo 4
If Combo1.ListIndex = -1 Then Exit Sub
Adodc2.Recordset.MoveFirst
Adodc2.Recordset.Move Combo1.ListIndex
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" Combo1_Keydown() Of mat2,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub

Private Sub Command1_Click()
On Error GoTo 4
If Label1.Caption = "������ �����" Then
Adodc1.Recordset.update
Adodc2.Recordset.update 'Adodc2.Recordset.Fields(2), Val(Label10.Caption)
Frame7.Visible = False
ElseIf Label1.Caption = "������ �����" Then
Adodc1.Recordset.update
DataGrid1.Refresh
Frame7.Visible = False
End If
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" Command1_Click() Of mat2,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub

Private Sub Command2_Click()
On Error GoTo 4
On Error Resume Next
Adodc1.Recordset.CancelBatch adAffectAllChapters
Adodc1.Recordset.CancelUpdate
Frame7.Visible = False
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" Command2_Click() Of mat2,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub

Private Sub Form_Activate()
On Error GoTo 4
Adodc1.RecordSource = Hav1: Adodc2.RecordSource = Kal1
Adodc1.Refresh
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" Form_activate() Of mat2,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
data1.Refresh
End Sub

Private Sub Form_Load()
On Error GoTo 4
Adodc1.ConnectionString = Havmdb
Adodc2.ConnectionString = Kalmdb
For i = LBound(Kala1) To UBound(Kala1)
Combo1.AddItem Trim$(Kala1(i).Name)
Next
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" Form_Load() Of mat2,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo 4
For i = 0 To 5
Ima(i).Refrash
Next
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" Form_Load() Of mat2,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo 4
Button1.Refrash: Button2.Refrash
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" Frame3_MouseMove() Of mat2,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub

Private Sub Ima_Click(Index As Integer)

Dim ssrt$
Select Case Index
    Case 0: Unload Me 'close dbase
    Case 2:
    Case 3:
     Frame3.Visible = True
     For i = 0 To 6
     Text(i).Text = ""
     Next
    Case 5:
         frmHavaleh.Show 1
           
   
End Select
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo 4
Image2.Visible = False
Image1.Visible = True
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" Image2_MouseDown() Of mat2,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo 4
Image2.Visible = True
Image1.Visible = False
Frame7.Visible = False
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" Image2_MouseUP() Of mat2,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub

Private Sub Frame6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo 4
Adodc1.RecordSource = Hav1: Adodc2.RecordSource = Kal1
Adodc1.Refresh: Adodc2.Refresh 'mathavaleh-frmhavale-havaleh
For i = 0 To 5
Ima(i).Refrash
Next
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" Frame6_MouseMove() Of mat2,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub

Private Sub Image4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo 4
Image4.Visible = False
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" image4_MouseDown() Of mat2,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub

Private Sub Image4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo 4
Image4.Visible = True
Frame3.Visible = False: List1.Clear
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" Label9_MouseUp() Of mat2,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label19_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo 4
Command1.Refrash
Command2.Refrash
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" Label9_MouseMove() Of mat2,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub

Private Sub Image5_Click()

End Sub

Private Sub List1_Click()
On Error GoTo 4
Adodc1.Recordset.MoveFirst
'FIXIT: Replace 'Left' function with 'Left$' function                                      FixIT90210ae-R9757-R1B8ZE
Adodc1.Recordset.Move Val(Left$(List1.Text, 3))
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" List1_Click() Of mat2,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub

Private Sub Text_Change(Index As Integer)
On Error GoTo 4
For i = 0 To 6
Text(i).BackColor = vbWhite
Next
Text(Index).BackColor = RGB(220, 255, 220)
Indesk = Index
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" text_changed(" & Index & ") Of mat2,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub

Private Sub Text3_Change()
On Error GoTo 4
If Frame7.Visible = False Then Exit Sub
Mojode = Havaleh.data1.Recordset(2).Value
Label10.Caption = Mojode + Val(Text3.Text)
Label14.Caption = Label10.Caption
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" Text3_Change() Of mat2,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub

Private Sub Text4_Change()
On Error GoTo 4
If Frame7.Visible = False Then Exit Sub
Mojode = Havaleh.data1.Recordset(2).Value
Label10.Caption = Val(Text3.Text) - Mojode
Label14.Caption = Label10.Caption
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" Text4_Change() Of mat2,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub

Public Sub Changedbass(ByVal RecsurHav As String, ByRef RecsurKal As String)
On Error GoTo 4
 Adodc1.RecordSource = RecsurHav
 Adodc2.RecordSource = RecsurKal
 Hav1 = RecsurHav
 Kal1 = RecsurKal
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" ChangeDbasse() Of mat2,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub




