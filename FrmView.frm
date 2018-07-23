VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmView 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form3"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9765
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   9765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid FGrid 
      Height          =   6375
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   11245
      _Version        =   393216
      RightToLeft     =   -1  'True
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   6480
      Width           =   9855
      Begin VB.Frame Frame2 
         Height          =   135
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   9855
      End
      Begin VB.OptionButton Option 
         BackColor       =   &H00FFFFFF&
         Caption         =   "„·“Ê„« "
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Option 
         BackColor       =   &H00FFFFFF&
         Caption         =   "„ÿ»Ê⁄« "
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option 
         BackColor       =   &H00FFFFFF&
         Caption         =   "‰„Ê‰Â Â«"
         Height          =   255
         Index           =   0
         Left            =   2280
         TabIndex        =   2
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
      Begin Project1.Button Button1 
         Height          =   375
         Left            =   8280
         TabIndex        =   1
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Caption         =   "»” ‰ ﬂ«œ—"
      End
      Begin Project1.Button Button2 
         Height          =   375
         Left            =   6840
         TabIndex        =   5
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Caption         =   "Å—Ì‰ "
      End
   End
End
Attribute VB_Name = "FrmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Use Option Explicit to avoid implicitly creating variables of type Variant         FixIT90210ae-R383-H1984
Private Sub Button1_Click()
On Error GoTo 4
Unload Me: main.Show
main.WindowState = 0
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" button1_click() Of frmview,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub

Private Sub Button2_Click()
On Error GoTo 4
Me.PrintForm
Exit Sub '-----{Call Loger In Erroring!}--------
MsgBox "Å—Ì‰ —Ì »—«Ì Å—Ì‰  Ì«›  ‰‘œ .·ÿ›¬ Å—Ì‰ — —« ‰’» ﬂ—œÂ Ì« « ’«· ¬‰ —« »——”Ì ﬂ‰Ìœ", vbInformation, "Œÿ« œ— ç«Åê—"
4 Call Loger(" button2_click() Of frmview,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub

Private Sub Form_Load()
On Error GoTo 5
   rcd.Open "havaleh", cnn, adOpenStatic, adCmdTable
Call RiempiGrid(rcd)
Exit Sub
5
MsgBox "»«‰ﬂ «ÿ·«⁄« Ì «‰ Œ«»Ì ‘„« »« ”«Œ «— «” «‰œ«—œ »—‰«„Â ”«“ê«— ‰Ì”  .·ÿ›¬ ›«Ì· œÌê—Ì —« «‰ Œ«» ﬂ‰Ìœ", vbInformation, "Anbar"
Unload Me
End Sub

Private Sub RiempiGrid(RecSet As Recordset)
On Error GoTo 4
Dim campo As Field
Dim temp As String
    On Error Resume Next
    FGrid.Rows = RecSet.RecordCount + 1
    'il numero di colonne Ë pari al numero di campi per record
    FGrid.Cols = RecSet.Fields.Count
    
    ct = 0
    'setto i nomi delle colonne
    For Each campo In RecSet.Fields
        FGrid.TextMatrix(0, ct) = campo.Name
        ct = ct + 1
    Next campo
    
    
    'scrivo i contenuti del recordset nella Fgrid
    For r = 1 To FGrid.Rows - 1
        c = 0
        For Each campo In RecSet.Fields
            
            FGrid.TextMatrix(r, c) = Fild2Str(campo.Value)
            
            If Err.Number <> 0 Then
                FGrid.TextMatrix(r, c) = ""
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
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" Riempigrid() Of frmview,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub

Private Sub Option_Click(Index As Integer)
Dim Fh$
On Error GoTo 4
rcd.Close
 Select Case Index
Case 0: Fh = "malhavaleh"
Case 1: Fh = "mathavale"
Case 2: Fh = "havaleh"
 End Select
rcd.Open Fh, cnn, adOpenStatic, adCmdTable
Call RiempiGrid(rcd)
4 Call Loger(" option_Click(" & Index & ") Of frmview,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub
