VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form mal1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "„·“Ê„« "
   ClientHeight    =   7575
   ClientLeft      =   2385
   ClientTop       =   2925
   ClientWidth     =   10185
   Icon            =   "mal1.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   10185
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame5 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   615
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   10575
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
         Left            =   3840
         TabIndex        =   17
         Top             =   0
         Width           =   3255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      Height          =   855
      Left            =   -1680
      TabIndex        =   6
      Top             =   555
      Width           =   12015
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         DataField       =   "ﬂœﬂ«·«"
         Height          =   320
         Left            =   9600
         TabIndex        =   8
         Top             =   360
         Width           =   735
      End
      Begin VB.ComboBox Combo1 
         DataField       =   "namekala"
         DataSource      =   "Adodc2"
         Height          =   315
         ItemData        =   "mal1.frx":0ECA
         Left            =   6000
         List            =   "mal1.frx":0ECC
         TabIndex        =   7
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "⁄œœ"
         DataField       =   "Ê«Õœﬂ«·«"
         Height          =   315
         Left            =   2640
         TabIndex        =   13
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         DataField       =   "„ÊÃÊœÌ"
         Height          =   315
         Left            =   3480
         TabIndex        =   12
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "„ÊÃÊœÌ"
         Height          =   315
         Left            =   4560
         TabIndex        =   11
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "‰Ê⁄ ﬂ«·«"
         Height          =   315
         Left            =   7920
         TabIndex        =   10
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ﬂœ ﬂ«·«"
         Height          =   315
         Left            =   9960
         TabIndex        =   9
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Height          =   135
      Left            =   0
      TabIndex        =   5
      Top             =   6600
      Width           =   10215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "mal1.frx":0ECE
      Height          =   5175
      Left            =   0
      TabIndex        =   15
      Top             =   1440
      Visible         =   0   'False
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   9128
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   16777215
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   15
      RightToLeft     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   6600
      Width           =   10455
      Begin VB.OptionButton Opt 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÕÊ«·Â Â«"
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Opt 
         BackColor       =   &H00FFFFFF&
         Caption         =   "‰„Ê‰Â Â«Ì „ÊÃÊœ"
         Height          =   255
         Index           =   0
         Left            =   4440
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   2055
      End
      Begin Project1.Button Command2 
         Height          =   375
         Left            =   8400
         TabIndex        =   3
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "«‰’—«›"
      End
      Begin Project1.Button Command1 
         CausesValidation=   0   'False
         Height          =   375
         Left            =   7080
         TabIndex        =   4
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "Å—Ì‰ "
      End
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   330
         Left            =   1680
         Top             =   480
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
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
         Caption         =   "Adodc2"
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
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   240
         Top             =   480
         Visible         =   0   'False
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   582
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
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "mal1.frx":0EE3
      Height          =   5415
      Left            =   -120
      TabIndex        =   14
      Top             =   1440
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   9551
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   15
      RightToLeft     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "mal1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Use Option Explicit to avoid implicitly creating variables of type Variant         FixIT90210ae-R383-H1984

Private Sub Combo1_Click()
malkala.data1.Refresh
With malkala.data1
.Recordset.MoveLast
Y = .Recordset.RecordCount
.Recordset.MoveFirst
For X = 0 To Y - 1
 If Combo1.Text = .Recordset(0).Value Then
    Label10.Caption = .Recordset(2).Value
    Label16.Caption = .Recordset(1).Value
    Text6.Text = .Recordset(3).Value
            .Refresh
      Exit For
 End If
.Recordset.MoveNext
Next
End With
End Sub

Private Sub Command1_Click()
On Error GoTo 4
Dim ii%, n%
DataGrid2.Columns(2).Width = 1000
DataGrid2.Columns(1).Width = 1500
DataGrid2.Columns(3).Width = 1000
DataGrid2.Columns(5).Width = 1000
DataGrid2.Columns(0).Width = 1000
'-------------------------
DataGrid1.Columns(3).Width = 1000
DataGrid1.Columns(5).Width = 1000
DataGrid1.Columns(0).Width = 1000
DataGrid1.RowHeight = 250
DataGrid2.RowHeight = 250
n = IIf(Opt(0).Value, DataGrid1.ApproxCount, DataGrid2.ApproxCount)
Me.PrintForm
ii = 18
While ii < DataGrid2.DataBindings.Count
    DataGrid2.Scroll 0, ii
    Me.PrintForm
    ii = ii + 18
Wend
Exit Sub '-----{Call Loger In Erroring!}--------
4: MsgBox "œ” ê«Â Å—Ì‰ — ‘„« Ì« Ê’· ‰Ì”  Ì« ¬„«œÂ ‰Ì”  ! ·ÿ›¬ »——”Ì ﬂ‰Ìœ", vbExclamation, "Œÿ«œ— Å—Ì‰ "
Call Loger(" command1_click() Of Mal1,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub

Private Sub Command2_Click()
On Error GoTo 4
Unload Me
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" command2_click() Of Mal1,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub



Private Sub Form_Activate()
'Adodc1.Refresh
Adodc2.Refresh
DataGrid2.Refresh
DataGrid1.Refresh
End Sub

Private Sub Form_Load()
On Error GoTo 4
Adodc1.Refresh
Adodc2.Refresh
Adodc1.ConnectionString = Havmdb
Adodc2.ConnectionString = Kalmdb
malkala.data1.Refresh
With malkala.data1
 .Recordset.MoveLast
 X = .Recordset.RecordCount
 .Recordset.MoveFirst
 For i = 0 To X - 1
  Combo1.AddItem .Recordset(0).Value
 .Recordset.MoveNext
 Next
End With
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" form_load() Of Mal1,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.Refrash: Command2.Refrash
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" frame3_MouseMove() Of Mal1,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub

Private Sub Opt_Click(Index As Integer)
On Error GoTo 4
Adodc2.Refresh
DataGrid2.Visible = Not DataGrid2.Visible
DataGrid1.Visible = Not DataGrid1.Visible
Frame1.Visible = Not Frame1.Visible
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" Opt(" & Index & ") Of ,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub

Private Sub Text6_Change()
On Error GoTo 4
With malkala.data1
.Recordset.MoveFirst
For X = 0 To .Recordset.RecordCount - 1
 If Text6.Text = .Recordset(3).Value Then
    Label10.Caption = .Recordset(2).Value
    Label16.Caption = .Recordset(1).Value
    Combo1.Text = .Recordset(0).Value
      .Refresh
      Exit For
 End If
 If Text6.Text <> .Recordset(3).Value Then
    Label10.Caption = ""
    Label16.Caption = ""
    Combo1.Text = ""
 End If
.Recordset.MoveNext
Next
End With
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" text6_change() Of Mal1,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub



