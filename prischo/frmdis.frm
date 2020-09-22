VERSION 5.00
Begin VB.Form frmdis 
   BorderStyle     =   0  'None
   ClientHeight    =   2835
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   2835
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Project1.ctrl_SkinableForm ctrl_SkinableForm 
      Height          =   735
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1296
      Caption         =   "Discount form"
      CaptionTop      =   0
   End
   Begin Project1.ctrl_SkinableButton btnOk 
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   2280
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Ok"
   End
   Begin VB.TextBox txtamo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   1800
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Discount Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1245
   End
End
Attribute VB_Name = "frmdis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnOk_Click()
Dim p, s, d As Integer
p = Val(txtamo.Text)
s = Val(frmpay.lblpay.Caption)
d = p + s
frmpay.lblpay.Caption = d
Unload Me

End Sub

Private Sub Combo1_Click()
Dim rs As Recordset
Call link
Set rs = db.OpenRecordset("select * from discount where discount='" & Combo1.Text & "'")
With rs
txtamo = rs("amount")
End With
End Sub

Private Sub Form_Load()
Me.Top = Val(Screen.Height - Me.Height) / 2
Me.Left = Val(Screen.Width - Me.Width) / 2
inito
StayOnTop Me, True
Dim rs As Recordset
Call link
Set rs = db.OpenRecordset("select * from discount")
With rs
While Not .EOF
Combo1.AddItem rs("discount")
.MoveNext
Wend
End With
End Sub

Sub inito()
With Me
.ctrl_SkinableForm.SkinPath = App.Path & "\Skins\Deco"
        .ctrl_SkinableForm.BackColor = &HCECECE
        .ctrl_SkinableForm.CaptionTop = 300
        .ctrl_SkinableForm.CaptionColor = &H0&
        Call Me.ctrl_SkinableForm.LoadSkin(Me)
        
        .btnOk.SkinPath = App.Path & "\Skins\ALPI"
        .btnOk.ForeColor = &H0&
        .btnOk.LoadSkin
        .btnOk.Refresh

End With
End Sub

Private Sub txtamo_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
                Case Asc("0") To Asc("9")
                Case Asc(".")
                    KeyAscii = IIf(Index = 5 Or Index = 1, 0, KeyAscii)
                Case KeyAscii = vbKeyBack
                Case Else
                    KeyAscii = 0
        End Select
End Sub
