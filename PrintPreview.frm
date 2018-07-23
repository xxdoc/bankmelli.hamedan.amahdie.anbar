VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmPrint 
   Caption         =   "Print"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9075
   LinkTopic       =   "Form3"
   ScaleHeight     =   6780
   ScaleWidth      =   9075
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   960
      Top             =   6480
      Width           =   2895
      _ExtentX        =   5106
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
   Begin VB.PictureBox Picture1 
      Height          =   5775
      Left            =   0
      ScaleHeight     =   5715
      ScaleWidth      =   8715
      TabIndex        =   0
      Top             =   0
      Width           =   8775
   End
End
Attribute VB_Name = "FrmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_oPreview As Preview

Public Sub Cload()
    With m_oPreview
        .Cls
        With .Pages
            .ScaleMode = vbInches
            .Width = 8.5
            .Height = 11
            .Add
            With .ActivePage
                .DrawPicture 3.25, 0.5, 2, 1.2, Picture1.Picture, True
                .SetFont "Tahoma", 24, True
                .DrawText "RamoSoft Print Preview Dll", 1, 1.5, 6, 2, vbBlue, , vbCenter
                .SetFont "Tahoma", 14, , True
                .DrawText "This demo show drawing capabilities of the RamoSoft Print Preview dll", _
                0.8, 3, 7, 2, vbBlack, vbCyan, vbCenter
                .SetFont "OCR A Extended", 72, True, , , , -45
                .DrawText "Cool!", 2, 3, 5, 4, vbRed
                '.DrawBox 0.5, 0.5, 1, 2.5, vbBlue, vbRed, 2
            End With
        End With
        .Show
    End With
End Sub

Private Sub DataRepeater1_Click()

End Sub

Private Sub Form_Load()
    Set m_oPreview = New Preview
    m_oPreview.Container = Picture1.hWnd
End Sub


Private Sub Form_Resize()
Picture1.Width = Me.Width
Picture1.Height = Me.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m_oPreview = Nothing
End Sub


