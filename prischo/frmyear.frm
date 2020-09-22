VERSION 5.00
Begin VB.Form frmyear 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   2055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4215
   LinkTopic       =   "Form2"
   ScaleHeight     =   2055
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox t2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   2520
      TabIndex        =   5
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox t1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
   Begin Project1.ctrl_SkinableForm ctrl_SkinableForm 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1720
      Caption         =   "School Year"
      CaptionTop      =   0
   End
   Begin Project1.ctrl_SkinableButton btnupdate 
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   1440
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "Update"
   End
   Begin Project1.ctrl_SkinableButton btnquit 
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   1440
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "Quit"
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   2400
      X2              =   2505
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "School Year"
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
      Left            =   480
      TabIndex        =   4
      Top             =   840
      Width           =   1050
   End
End
Attribute VB_Name = "frmyear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
        
        .btnupdate.SkinPath = App.Path & "\Skins\Deco"
        .btnupdate.ForeColor = &H0&
        .btnupdate.LoadSkin
        .btnupdate.Refresh
        
End With
End Sub

Private Sub btnquit_Click()
Unload Me

End Sub

Private Sub btnquit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
btnquit.Refresh
End Sub

Private Sub btnupdate_Click()
StayOnTop Me, 0
Dim a, B As Integer
a = Val(Trim(t1.Text))
B = Val(Trim(t2.Text))
c = Val(B - a)
'MsgBox c
'Exit Sub
If c = 1 Then
    


        Dim rs As Recordset
        Call link
        Set rs = db.OpenRecordset("select * from year")
        With rs
        
            .AddNew
            rs("year") = t1 + "/" + t2
            .Update
        d = MsgBox("Year Updated")
        
        End With

Else

    j = MsgBox("The Year is Ilogical try again")
Exit Sub

End If
StayOnTop Me, True
End Sub

Private Sub btnupdate_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
btnupdate.Refresh
End Sub

Private Sub Form_Load()
StayOnTop Me, True
Me.Top = Val(Screen.Height - Me.Height) / 2
Me.Left = Val(Screen.Width - Me.Width) / 2
inito
Filler
End Sub
Sub Filler()
Dim rs As Recordset
Call link
Set rs = db.OpenRecordset("select * from year")
With rs
.MoveLast
t1 = Mid(rs("year"), 1, 4)
t2 = Right(rs("year"), 4)

End With

End Sub

Private Sub t1_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
                Case Asc("0") To Asc("9")
                Case Asc(".")
                    KeyAscii = IIf(Index = 5 Or Index = 1, 0, KeyAscii)
                Case KeyAscii = vbKeyBack
                Case Else
                    KeyAscii = 0
        End Select
End Sub



Private Sub t2_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
                Case Asc("0") To Asc("9")
                Case Asc(".")
                    KeyAscii = IIf(Index = 5 Or Index = 1, 0, KeyAscii)
                Case KeyAscii = vbKeyBack
                Case Else
                    KeyAscii = 0
        End Select
End Sub
