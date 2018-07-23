VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form shoab 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "shoab"
   ClientHeight    =   7020
   ClientLeft      =   4710
   ClientTop       =   3120
   ClientWidth     =   6720
   Icon            =   "shoab.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc data1 
      Height          =   375
      Left            =   0
      Top             =   5760
      Width           =   6735
      _ExtentX        =   11880
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
      RecordSource    =   "shoab"
      Caption         =   "shoab"
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   135
      Left            =   -360
      TabIndex        =   11
      Top             =   2160
      Width           =   7095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "shoab.frx":0ECA
      Height          =   3495
      Left            =   0
      TabIndex        =   9
      Top             =   2280
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   6165
      _Version        =   393216
      BackColor       =   -2147483639
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   855
      Left            =   -120
      Top             =   3000
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1508
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
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "nameshobe"
      DataSource      =   "Data1"
      Height          =   370
      Index           =   0
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   1
      Top             =   480
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "idshobe"
      DataSource      =   "Data1"
      Height          =   370
      Index           =   1
      Left            =   1320
      MaxLength       =   20
      TabIndex        =   0
      Top             =   960
      Width           =   3375
   End
   Begin Project1.Button Button 
      Height          =   375
      Index           =   0
      Left            =   840
      TabIndex        =   2
      Top             =   1560
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "«÷«›Â"
   End
   Begin Project1.Button Button 
      Height          =   375
      Index           =   1
      Left            =   1920
      TabIndex        =   3
      Top             =   1560
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "Õ–›"
   End
   Begin Project1.Button Button 
      Height          =   375
      Index           =   2
      Left            =   3000
      TabIndex        =   4
      Top             =   1560
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "»—Ê“ﬂ—œ‰"
   End
   Begin Project1.Button Button 
      Height          =   375
      Index           =   3
      Left            =   4080
      TabIndex        =   5
      Top             =   1560
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "ÊÌ—«Ì‘"
   End
   Begin Project1.Button Button 
      Height          =   375
      Index           =   4
      Left            =   5160
      TabIndex        =   6
      Top             =   1560
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "»” ‰"
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   855
      Left            =   -480
      TabIndex        =   12
      Top             =   6120
      Width           =   7215
      Begin Project1.Button Command2 
         Height          =   375
         Left            =   5880
         TabIndex        =   13
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "«‰’—«›"
      End
      Begin Project1.Button Command1 
         CausesValidation=   0   'False
         Height          =   375
         Left            =   4560
         TabIndex        =   14
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "Å—Ì‰ "
      End
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "ÊÌ—«Ì‘ Ê «›“Êœ‰ ‘⁄»Â"
      Height          =   255
      Left            =   4680
      TabIndex        =   10
      Top             =   120
      Width           =   1575
   End
   Begin VB.Line Line1 
      X1              =   2760
      X2              =   3960
      Y1              =   2520
      Y2              =   3000
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ﬂœ ‘⁄»Â"
      Height          =   375
      Index           =   1
      Left            =   3960
      TabIndex        =   7
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ ‘⁄»Â"
      Height          =   255
      Index           =   0
      Left            =   3840
      TabIndex        =   8
      Top             =   600
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   1815
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   6255
   End
End
Attribute VB_Name = "shoab"
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
               ' data1.Recordset.Bookmark = data1.Recordset.LastModified
       Case 4:  Unload Me
End Select
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger("Button_Click() Of Shobe,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)

End Sub

Private Sub data1_Validate(Action As Integer, Save As Integer)

End Sub


Private Sub Command1_Click()
Me.PrintForm
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On erro GoTo 4
For i = 1 To 4
Me.Button(i).Refrash
Next
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger("Form_MouseMove() Of Shobe,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.Refrash
Command2.Refrash
End Sub
