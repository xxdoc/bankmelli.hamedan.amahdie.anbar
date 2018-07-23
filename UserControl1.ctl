VERSION 5.00
Begin VB.UserControl Button 
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1875
   ScaleHeight     =   495
   ScaleWidth      =   1875
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Button1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Image Image33 
      Height          =   495
      Left            =   1440
      Picture         =   "UserControl1.ctx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   345
   End
   Begin VB.Image Image32 
      Height          =   495
      Left            =   0
      Picture         =   "UserControl1.ctx":034A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
   Begin VB.Image Image12 
      Height          =   495
      Left            =   0
      Picture         =   "UserControl1.ctx":067D
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image11 
      Height          =   495
      Left            =   1440
      Picture         =   "UserControl1.ctx":0F03
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image22 
      Height          =   495
      Left            =   0
      Picture         =   "UserControl1.ctx":1264
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image21 
      Height          =   495
      Left            =   1440
      Picture         =   "UserControl1.ctx":15CB
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   0
      Picture         =   "UserControl1.ctx":1933
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   0
      Picture         =   "UserControl1.ctx":1BFE
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Image Image3 
      Height          =   495
      Left            =   0
      Picture         =   "UserControl1.ctx":1EC6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "Button"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'FIXIT: Use Option Explicit to avoid implicitly creating variables of type Variant         FixIT90210ae-R383-H1984
Const DefultCap = "Button1"
Public Event Click()
Public Event MouseMove(Button%, Shift%, X!, Y!)
Public Event MouseDown(Button%, Shift%, X!, Y!)
Public Event MouseUp(Button%, Shift%, X!, Y!)
'-------------------------
Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo 4
Image1.Visible = False: Image11.Visible = False: Image12.Visible = False
Image2.Visible = True: Image22.Visible = True: Image21.Visible = True
RaiseEvent MouseDown(Button, Shift, X, Y)
4
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo 4
Image2.Visible = False: Image22.Visible = False: Image21.Visible = False
Image3.Visible = True: Image33.Visible = True: Image32.Visible = True
RaiseEvent MouseUp(Button, Shift, X, Y)
4
End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo 4
Image2.Visible = False: Image22.Visible = False: Image21.Visible = False
Image3.Visible = True: Image32.Visible = True: Image33.Visible = True
RaiseEvent MouseUp(Button, Shift, X, Y)
4
End Sub
Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo 4
Image3.Visible = False: Image32.Visible = False: Image33.Visible = False
Image1.Visible = True: Image12.Visible = True: Image11.Visible = True
RaiseEvent MouseMove(Button, Shift, X, Y)
4
End Sub

Private Sub Image12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo 4
Call Image1_MouseDown(Button, Shift, X, Y)
4
End Sub
Private Sub Image12_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo 4
Call Image1_MouseUp(Button, Shift, X, Y)
4
End Sub
'--------------------------
Private Sub Image11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo 4
Call Image1_MouseDown(Button, Shift, X, Y)
4
End Sub
Private Sub Image11_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo 4
Call Image1_MouseUp(Button, Shift, X, Y)
4
End Sub
'--------------------------
Private Sub Image21_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo 4
Call Image2_MouseUp(Button, Shift, X, Y)
4
End Sub
'--------------------------
Private Sub Image22_Mouseup(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo 4
Call Image2_MouseUp(Button, Shift, X, Y)
4
End Sub
'-------------------------------------------
Private Sub Image32_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo 4
Call Image3_MouseMove(Button, Shift, X, Y)
4
End Sub
'---------------------------
Private Sub Image33_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo 4
Call Image3_MouseMove(Button, Shift, X, Y)
4
End Sub

Private Sub Label1_Click()
On Error GoTo 4
RaiseEvent Click
4
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo 4
Image1.Visible = False: Image12.Visible = False: Image11.Visible = False
Image2.Visible = True: Image22.Visible = True: Image21.Visible = True
RaiseEvent MouseDown(Button, Shift, X, Y)
4
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo 4
Image3.Visible = False: Image33.Visible = False: Image32.Visible = False
Image1.Visible = True: Image12.Visible = True: Image12.Visible = True
RaiseEvent MouseMove(Button, Shift, X, Y)
4
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo 4
Image1.Visible = False: Image11.Visible = False: Image12.Visible = False
Image3.Visible = True: Image33.Visible = True: Image32.Visible = True
RaiseEvent MouseUp(Button, Shift, X, Y)
4
End Sub

Private Sub UserControl_Initialize()
UserControl.Refresh

End Sub

Private Sub UserControl_Resize()
On Error GoTo 4
Dim W%, H%
H = UserControl.Height: W = UserControl.Width
Image1.Width = W: Image1.Height = H
Image2.Width = W: Image2.Height = H
Image3.Width = W: Image3.Height = H
Label1.Width = W: Label1.Top = (H \ 2) - (Label1.Height \ 2)
'------------------------
Image12.Height = H: Image11.Height = H
Image22.Height = H: Image21.Height = H
Image32.Height = H: Image33.Height = H
'------------------------
Image11.Left = W - Image11.Width
Image21.Left = W - Image21.Width
Image33.Left = W - Image33.Width
4
End Sub
Public Sub Refrash()
On Error GoTo 4
Image2.Visible = False: Image21.Visible = False: Image22.Visible = False
Image1.Visible = False: Image11.Visible = False: Image12.Visible = False
Image3.Visible = True: Image32.Visible = True: Image33.Visible = True: CapValue = Label1.Caption
4
End Sub
Public Property Get Caption() As String
On Error GoTo 4
 Caption = Label1.Caption
4
End Property
'------------------------
Public Property Let Caption(Value$)
On Error GoTo 4
Label1.Caption = Value: PropertyChanged "Caption"
4
End Property

Private Sub UserControl_Terminate()
UserControl.Refresh
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error GoTo 4
Call PropBag.WriteProperty("Caption", Label1.Caption, DefultCap)
4
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error GoTo 4
Label1.Caption = PropBag.ReadProperty("Caption", DefultCap)
4
End Sub
