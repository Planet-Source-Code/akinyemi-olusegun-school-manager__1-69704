VERSION 5.00
Begin VB.Form frmdeduction 
   BorderStyle     =   0  'None
   ClientHeight    =   3690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4620
   LinkTopic       =   "Form2"
   ScaleHeight     =   3690
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtamo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   2040
      TabIndex        =   7
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox txtded 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   2040
      TabIndex        =   5
      Top             =   1680
      Width           =   1695
   End
   Begin VB.ComboBox cbomon 
      Height          =   315
      ItemData        =   "frmdeduction.frx":0000
      Left            =   2040
      List            =   "frmdeduction.frx":0028
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1080
      Width           =   1695
   End
   Begin Project1.ctrl_SkinableForm ctrl_SkinableForm 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      Caption         =   "Deduction"
      CaptionTop      =   0
   End
   Begin Project1.ctrl_SkinableButton btnquit 
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   3000
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "Close"
   End
   Begin Project1.ctrl_SkinableButton btnprint 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   3000
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "&Save"
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Deduct"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3840
      TabIndex        =   9
      Top             =   2520
      Width           =   630
   End
   Begin VB.Label Label1 
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
      Left            =   360
      TabIndex        =   8
      Top             =   2160
      Width           =   645
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Deduction Type"
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
      Left            =   360
      TabIndex        =   6
      Top             =   1680
      Width           =   1365
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Month"
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
      Left            =   360
      TabIndex        =   4
      Top             =   1200
      Width           =   540
   End
End
Attribute VB_Name = "frmdeduction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnprint_Click()
StayOnTop Me, 0
Dim rs As Recordset
Call link
Set rs = db.OpenRecordset("select * from deduction")
    With rs
        .AddNew
            rs("type") = txtded.Text
            rs("month") = cbomon.Text
            rs("Amount") = txtamo.Text
            
                   
        .Update
    
    End With
    StayOnTop Me, True
End Sub


Private Sub btnprint_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
btnprint.Refresh
End Sub

Private Sub btnquit_Click()
Unload Me


End Sub

Private Sub btnquit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
btnquit.Refresh
End Sub

Private Sub Form_Load()
Me.Top = Val(Screen.Height - Me.Height) / 2
Me.Left = Val(Screen.Width - Me.Width) / 2
inito
StayOnTop Me, True
End Sub
Sub inito()
With Me
.ctrl_SkinableForm.SkinPath = App.Path & "\Skins\Deco"
        .ctrl_SkinableForm.BackColor = &HCECECE
        .ctrl_SkinableForm.CaptionTop = 300
        .ctrl_SkinableForm.CaptionColor = &H0&
        Call Me.ctrl_SkinableForm.LoadSkin(Me)

.btnquit.SkinPath = App.Path & "\Skins\Deco"
        .btnquit.ForeColor = &H0&
        .btnquit.LoadSkin
        .btnquit.Refresh
        
        .btnprint.SkinPath = App.Path & "\Skins\Deco"
        .btnprint.ForeColor = &H0&
        .btnprint.LoadSkin
    .btnprint.Refresh
End With
End Sub

Private Sub Label2_Click()
frmded.Show
Unload Me

End Sub
