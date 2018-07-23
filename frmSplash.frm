VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5175
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   10320
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":628A
   ScaleHeight     =   5175
   ScaleWidth      =   10320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   240
      Top             =   2040
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Suported By:WWW.VBOOK.COO.IR"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7200
      TabIndex        =   5
      Top             =   4440
      Width           =   3135
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Licenc By:Mr Zareiy Fo Meli Bank"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7200
      TabIndex        =   4
      Top             =   4680
      Width           =   2895
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright By:NasserNiazy and MasoudAlavi      2008 All RightReserved.                                           GHYESHSOFT.Inc"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   4440
      Width           =   4095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0.2.05"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8760
      TabIndex        =   2
      Top             =   0
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "GhayeshSoft Anbar DBasse 2008"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Loading ..."
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   4920
      Width           =   10935
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim D6%

Private Sub Form_Click()
On Error GoTo 4
Unload Me
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" Form_Click() Of FrmSplash,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub

Private Sub Form_Load()
On Error GoTo 4
Me.Left = (Screen.Width \ 2) - (Me.Width \ 2)
Me.Top = (Screen.Height \ 2) - (Me.Height \ 2)
main.Show
main.Visible = False
main.WindowState = 1
Me.Show
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" Form_Load() Of FrmSplash,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo 4
main.Visible = True
main.WindowState = 0
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" Form_Unload() Of FrmSplash,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub

Private Sub Timer1_Timer()
On Error GoTo 4
Me.Enabled = False
D6 = D6 + 10
Randomize Timer
Label1.Caption = "Loading ... " & App.Path & "DataFill" & Int(Rnd() * 1000) & ".DLL"
If D6 = 270 Then main.Visible = True
If D6 >= 300 Then
Unload Me
End If
Exit Sub '-----{Call Loger In Erroring!}--------
4 Call Loger(" Timer1_Timer() Of FrmSplash,Src:" & Err.Source & " ,Num:" _
                        & Err.Number & " Bug:" & Err.Description)
End Sub
