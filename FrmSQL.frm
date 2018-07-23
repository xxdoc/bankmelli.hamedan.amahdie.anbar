VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmSQL 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Advance Search"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10620
   Icon            =   "FrmSQL.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   10620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   7440
      TabIndex        =   15
      Top             =   840
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "FrmSQL.frx":0ECA
      Height          =   3495
      Left            =   1560
      TabIndex        =   14
      Top             =   1320
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   6165
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   6480
      TabIndex        =   12
      Top             =   840
      Width           =   855
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "FrmSQL.frx":0EDF
      Left            =   3480
      List            =   "FrmSQL.frx":0EF2
      TabIndex        =   11
      Text            =   "Combo2"
      Top             =   720
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid Fgrid 
      Height          =   3735
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   6588
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   120
      Top             =   0
      Width           =   3855
      _ExtentX        =   6800
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Project\1\HAVALE.MDB;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Project\1\HAVALE.MDB;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "havaleh"
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
   Begin Project1.Button Button2 
      Height          =   375
      Left            =   5040
      TabIndex        =   8
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "SELECT"
   End
   Begin Project1.Button Button1 
      Height          =   375
      Left            =   5280
      TabIndex        =   7
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Caption         =   "Conect"
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "FrmSQL.frx":0F21
      Left            =   960
      List            =   "FrmSQL.frx":0F23
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   720
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   450
      ItemData        =   "FrmSQL.frx":0F25
      Left            =   3720
      List            =   "FrmSQL.frx":0F2F
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Value"
      Height          =   255
      Left            =   6480
      TabIndex        =   13
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "FildName"
      Height          =   375
      Left            =   2640
      TabIndex        =   10
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   5520
      Width           =   7575
   End
   Begin VB.Label Label3 
      Caption         =   "Table"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "DBase"
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "FrmSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Use Option Explicit to avoid implicitly creating variables of type Variant         FixIT90210ae-R383-H1984
Public cnn As New Connection
'Dichiaro public anche il recordset per gestire la Fgrid
Public rcd As New Recordset
'Variabile che memorizza la tabella correntemente selezionata nella combo
Public CurrentTable As String
'Variabile che memorizza il percorso del database
Public PathDatabase As String


Private Sub Button1_Click()

  '  On Error Resume Next

    'sistemo la mask e resetto FileName per gestire l'Annulla
    
    'memorizzo il pathname prima nella var globale
    PathDatabase = App.Path + "\" + List1.Text + ".mdb"
    
    'se premo annulla deve uscire
    If List1.Text = "" Then MsgBox "Plase Select Dbasse": Exit Sub
    
    'se la connessione Ë aperta la chiudo
    If cnn.State = adStateOpen Then cnn.Close
    
    'pizzo la connection
    cnn.Mode = adModeRead
    cnn.Provider = "Microsoft.Jet.OLEDB.4.0"
    cnn.Open PathDatabase
    
    'in caso di errore me ne esco
    If Err.Number <> 0 Then
        MsgBox "Errori rilevati durante la connessione al database specificato. Impossibile continuare" & vbCrLf & _
        Err.Description, vbCritical, "Errore interno"
        Exit Sub
    End If
    
    'svuoto la Fgrid
    Fgrid.Clear
    Fgrid.Cols = 2
    Fgrid.Rows = 2
    'dichiaro lo stato attuale della connessione
    Label4.Caption = "Status : Connected!"
    'visualizzo il pathname nella textbox (cosÏ parte l'evento change)
   Text1.Text = PathDatabase
 
End Sub


Private Sub LeggiTabelle(ByVal pathname As String)
Dim cat As New Catalog
Dim tbl As Table

    'bindo il catalog alla connection
    Set cat.ActiveConnection = cnn
    
    'mi giro tutte le tabelle
    For Each tbl In cat.Tables
'FIXIT: Replace 'LCase' function with 'LCase$' function                                    FixIT90210ae-R9757-R1B8ZE
'FIXIT: Replace 'Left' function with 'Left$' function                                      FixIT90210ae-R9757-R1B8ZE
        If LCase$(Left$(tbl.Name, 4)) <> "msys" Then
            Combo1.AddItem tbl.Name
        End If
    Next tbl

End Sub

Private Sub Button2_Click()
Combo1_Change
End Sub
'"‰«„ ﬂ«·«-Ê«Õœﬂ«·«-„ÊÃÊœÌ-ﬂœﬂ«·«-”—Ì«·"
Private Sub Combo1_Change()
Dim bng As Adodc

''    On Error Resume Next
   Dim StrSql$
'    rcd.Close
    If Err.Number <> 0 Then
        Err.Clear
    End If
    StrSql = "SELECT " & Combo1.Text & ".[" & Combo2.Text & "]" & " FORM " & Combo1.Text
    If Text2.Text <> "" Then StrSql = StrSql & vbCrLf & _
    "WHERE (([" & Combo1.Text & "]![" & Combo2.Text & "]=" & Text2.Text & "));"
    'memorizzo la tabella correntemente selezionata
    CurrentTable = Combo1.List(Combo1.ListIndex)
    'apro la tabella selezionata
    Debug.Print StrSql
    StrSql = "SELECT " & Combo1.Text & ".[‘„«—Â ÕÊ«·Â], havaleh.[‰«„ ‘⁄»Â], havaleh.[ﬂœ ‘⁄»Â], havaleh.[ﬂœ ﬂ«·«], havaleh.’«œ—Â, havaleh.Ê«—œÂ, havaleh. «—ÌŒ  From Havaleh WHERE (([ﬂœ ﬂ«·«]='001'));"
'"SELECT havaleh.[‘„«—Â ÕÊ«·Â] , havaleh.[‰«„ ‘⁄»Â], havaleh.[ﬂœ ‘⁄»Â], havaleh.[ﬂœ ﬂ«·«] FROM havaleh "
    rcd.Open StrSql, cnn, adOpenStatic, adCmdTable
'Adodc1.Recordset.Close
'Adodc1.Recordset.Open StrSql, cnn, adOpenStatic, adCmdTable
'Adodc1.RecordSource = StrSql
'DataGrid1.Refresh
'    For i = 1 To rcd.RecordCount
'Adodc1.Recordset.Fields(i) = rcd.Fields(i)
'    Next
 Call RiempiGrid(rcd)
    'quindi chiudo per evitare conflitti
    rcd.Close
'SELECT havaleh.[‘„«—Â ÕÊ«·Â], havaleh.[‰«„ ‘⁄»Â], havaleh.[ﬂœ ‘⁄»Â], havaleh.[ﬂœ ﬂ«·«], Kala.[‰«„ ﬂ«·«], Kala.„ÊÃÊœÌ, Kala.ﬂœﬂ«·«
'FROM Kala INNER JOIN havaleh ON Kala.ﬂœﬂ«·« = havaleh.[ﬂœ ﬂ«·«]
'WHERE (([havaleh]![ﬂœ ﬂ«·«]="76"));

End Sub

Private Sub RiempiGrid(RecSet As Recordset)
Dim campo As Field
Dim temp As String
    'On Error Resume Next
    'il numero di righe Ë pari al numero di record
    Fgrid.Rows = RecSet.RecordCount + 1
    'il numero di colonne Ë pari al numero di campi per record
    Fgrid.Cols = RecSet.Fields.Count
    ct = 0
    'setto i nomi delle colonne
    For Each campo In RecSet.Fields
        Fgrid.TextMatrix(0, ct) = campo.Name
        ct = ct + 1
    Next campo
    'scrivo i contenuti del recordset nella Fgrid
    For r = 1 To Fgrid.Rows - 1
        c = 0
        For Each campo In RecSet.Fields
            Fgrid.TextMatrix(r, c) = Fild2Str(campo.Value)
            If Err.Number <> 0 Then
                Fgrid.TextMatrix(r, c) = ""
                Err.Clear
            End If
            c = c + 1
        Next campo
        RecSet.MoveNext
    Next r
    If Err.Number <> 0 Then
        MsgBox "Errori rilevati aprendo la tabella selezionata. Impossibile continuare", vbCritical, "Errore Interno"
        MsgBox Err.Description
        Exit Sub
    End If
   Label4.Caption = " Elementi trovati : " & rcd.RecordCount
End Sub

Private Sub DataGrid1_Click()

End Sub

Private Sub Text1_Change()
   Combo1.Clear
    
    If PathDatabase <> "" Then
        Call LeggiTabelle(PathDatabase)
    Else
        'se non ho un pathname per il database lo segnalo
        Combo1.AddItem "- nessuna tabella selezionata -"
    End If
End Sub
