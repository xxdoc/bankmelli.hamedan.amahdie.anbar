VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ê“«—‘ „ÊÃÊœÌ"
   ClientHeight    =   7815
   ClientLeft      =   3000
   ClientTop       =   3165
   ClientWidth     =   9915
   Icon            =   "FrmPrint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   9915
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   615
      Left            =   0
      TabIndex        =   15
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
         TabIndex        =   16
         Top             =   0
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      Height          =   855
      Left            =   -2160
      TabIndex        =   6
      Top             =   600
      Width           =   12015
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         DataField       =   "ﬂœﬂ«·«"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0;(0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   0
         EndProperty
         Height          =   320
         Left            =   9360
         TabIndex        =   8
         Top             =   360
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
         DataField       =   "namekala"
         DataSource      =   "Adodc1"
         Height          =   315
         ItemData        =   "FrmPrint.frx":0ECA
         Left            =   6000
         List            =   "FrmPrint.frx":0ECC
         TabIndex        =   7
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         DataField       =   "„ÊÃÊœÌ"
         Height          =   315
         Left            =   3360
         TabIndex        =   14
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "⁄œœ"
         DataField       =   "Ê«Õœﬂ«·«"
         Height          =   315
         Left            =   2520
         TabIndex        =   12
         Top             =   360
         Width           =   855
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
         Left            =   10080
         TabIndex        =   9
         Top             =   360
         Width           =   1575
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "FrmPrint.frx":0ECE
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   9340
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
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   0
      TabIndex        =   1
      Top             =   6720
      Width           =   10455
      Begin Project1.Button Command2 
         Height          =   375
         Left            =   8400
         TabIndex        =   5
         Top             =   240
         Width           =   1215
         _extentx        =   2143
         _extenty        =   661
         caption         =   "«‰’—«›"
      End
      Begin Project1.Button Command1 
         CausesValidation=   0   'False
         Height          =   375
         Left            =   7080
         TabIndex        =   4
         Top             =   240
         Width           =   1215
         _extentx        =   2143
         _extenty        =   661
         caption         =   "Å—Ì‰ "
      End
      Begin VB.OptionButton Opt 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÕÊ«·Â Â«"
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Opt 
         BackColor       =   &H00FFFFFF&
         Caption         =   "‰„Ê‰Â Â«Ì „ÊÃÊœ"
         Height          =   255
         Index           =   0
         Left            =   4200
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   1935
      End
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   330
         Left            =   2280
         Top             =   480
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
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
         RecordSource    =   "havale"
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
         Width           =   1920
         _ExtentX        =   3387
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
         RecordSource    =   "kala"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "FrmPrint.frx":0EE3
      Height          =   6015
      Left            =   120
      TabIndex        =   13
      Top             =   720
      Visible         =   0   'False
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   10610
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
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Combo1_Click()
Havaleh.data1.Refresh
 With Havaleh.data1
 .Recordset.MoveLast
 Y = .Recordset.RecordCount
.Recordset.MoveFirst
For X = 0 To Y - 1
 If Combo1.Text = .Recordset(0).Value Then
    Label1.Caption = .Recordset(2).Value
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
While i < DataGrid2.DataBindings.Count
    DataGrid2.Scroll 0, ii
    Me.PrintForm
    ii = ii + 18
Wend
Exit Sub '-----{Call Loger In Erroring!}--------
4: MsgBox "œ” ê«Â Å—Ì‰ — ‘„« Ì« Ê’· ‰Ì”  Ì« ¬„«œÂ ‰Ì”  ! ·ÿ›¬ »——”Ì ﬂ‰Ìœ", vbExclamation, "Œÿ«œ— Å—Ì‰ "
Call Loger(" Command1_Click() Of ,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub

Private Sub Command2_Click()
On Error GoTo 4
Unload Me
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" Comand2_Click() Of ,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub





Private Sub Form_Activate()
Adodc2.Refresh
End Sub

Private Sub Form_Load()
Adodc2.Refresh
With frmKala.data1
 .Recordset.MoveLast
 X = .Recordset.RecordCount
 .Recordset.MoveFirst
 For i = 0 To X - 1
 Combo1.AddItem .Recordset(0).Value
 .Recordset.MoveNext
 Next
End With
End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Command1.Refrash: Command2.Refrash

End Sub

Private Sub Opt_Click(Index As Integer)
Adodc2.Refresh
DataGrid2.Visible = Not DataGrid2.Visible
DataGrid1.Visible = Not DataGrid1.Visible
Frame1.Visible = Not Frame1.Visible

End Sub

Private Sub Text6_Change()
frmKala.data1.Refresh
With frmKala.data1
.Recordset.MoveLast
Y = .Recordset.RecordCount
.Recordset.MoveFirst
For X = 0 To Y - 1
 If Text6.Text = .Recordset(3).Value Then
    Label1.Caption = .Recordset(2).Value
    Label16.Caption = .Recordset(1).Value
    Combo1.Text = .Recordset(0).Value
      .Refresh
      Exit For
 End If
 If Text6.Text <> .Recordset(3).Value Then
    Label1.Caption = ""
    Label16.Caption = ""
    Combo1.Text = ""
 End If
.Recordset.MoveNext
Next
End With
Exit Sub '-----{Call Loger In Erroring!}--------

End Sub
