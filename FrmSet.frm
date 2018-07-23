VERSION 5.00
Object = "{8AB3C700-28FF-476A-852B-5E1558F44746}#1.0#0"; "command.ocx"
Begin VB.Form FrmSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form3"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5910
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin OsenXPCntrl.Command Command2 
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   4320
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Cancel"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "FrmSet.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.Command Command1 
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   4320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "OK"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "FrmSet.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame2 
      Caption         =   " „"
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   5655
      Begin VB.OptionButton Option3 
         Caption         =   "¬»Ì"
         Height          =   495
         Left            =   2760
         TabIndex        =   6
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         Caption         =   "”»“"
         Height          =   375
         Left            =   1680
         TabIndex        =   5
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "«’·Ì"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "‰Ê⁄ ‰„«Ì‘ œﬂ„Â Â«"
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "FrmSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
