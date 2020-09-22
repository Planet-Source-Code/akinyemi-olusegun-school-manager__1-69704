VERSION 5.00
Begin VB.Form frmlog 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   4155
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5835
   LinkTopic       =   "Form2"
   ScaleHeight     =   4155
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   480
      Picture         =   "frmlog.frx":0000
      ScaleHeight     =   495
      ScaleWidth      =   480
      TabIndex        =   9
      Top             =   1200
      Width           =   480
   End
   Begin VB.ComboBox cboid 
      Height          =   315
      ItemData        =   "frmlog.frx":0CA2
      Left            =   1920
      List            =   "frmlog.frx":0CA4
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox txtnam 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   1920
      Width           =   3135
   End
   Begin VB.TextBox txtpass 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   2520
      Width           =   1935
   End
   Begin Project1.ctrl_SkinableForm ctrl_SkinableForm 
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   855
      _extentx        =   1508
      _extenty        =   1296
      caption         =   "Log in"
      forecolor       =   16777215
      captiontop      =   0
      captioncolor    =   16777215
   End
   Begin Project1.ctrl_SkinableButton btnsave 
      Height          =   495
      Left            =   1080
      TabIndex        =   7
      Top             =   3240
      Width           =   1455
      _extentx        =   2566
      _extenty        =   873
      caption         =   "Login"
   End
   Begin Project1.ctrl_SkinableButton btnquit 
      Height          =   375
      Left            =   3600
      TabIndex        =   8
      Top             =   3240
      Width           =   1455
      _extentx        =   2566
      _extenty        =   661
      caption         =   "Quit"
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Staff ID."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1080
      TabIndex        =   6
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   960
      TabIndex        =   4
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   960
      TabIndex        =   3
      Top             =   2520
      Width           =   825
   End
End
Attribute VB_Name = "frmlog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnquit_Click()
Unload Me
End Sub

Private Sub btnquit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
btnquit.Refresh
End Sub

Private Sub btnsave_Click()
StayOnTop Me, 0
Dim rs As Recordset
Call link
Set rs = db.OpenRecordset("select * from user where id='" & cboid.Text & "' and password='" & txtpass.Text & "'")
With rs
    If .RecordCount < 1 Then
        h = MsgBox("Invalid password or account doesnt exist " + vbCrLf + "Contact the administrator", vbCritical, PROJ)
        
    Exit Sub
    Else
    f = rs("name")
   Unload frmmain
         'Unload Me
         'Unload frmlog
         frmlog.Hide
         
      frmmain.status.Panels(1).Text = "User:  " + txtnam.Text
               frmmain.mnuuser.Caption = f

        frmmain.Show
    Unload Me
    
End If

End With
End Sub

Private Sub btnsave_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
btnsave.Refresh
End Sub

Private Sub cboid_Click()
Dim rs As Recordset
Call link
Set rs = db.OpenRecordset("select * from user where id='" & cboid.Text & "'")
    With rs
        txtnam.Text = rs("name")
            
    End With
End Sub

Private Sub Form_Load()

Me.Top = Val(Screen.Height - Me.Height) / 2
Me.Left = Val(Screen.Width - Me.Width) / 2
On Local Error Resume Next
inito
StayOnTop Me, True


Dim rs As Recordset
Call link
Set rs = db.OpenRecordset("select all id from user")
With rs
    While Not .EOF
        cboid.AddItem rs("id")
    .MoveNext
    Wend

End With
End Sub
Sub inito()
With Me
 .ctrl_SkinableForm.SkinPath = App.Path & "\Skins\Blue"
        .ctrl_SkinableForm.BackColor = &HBD6E06
        .ctrl_SkinableForm.CaptionTop = 250
        .ctrl_SkinableForm.CaptionColor = &H0&
        Call Me.ctrl_SkinableForm.LoadSkin(Me)
        
        .btnquit.SkinPath = App.Path & "\Skins\Blue"
       .btnquit.ForeColor = &H0&
       .btnquit.LoadSkin
        .btnquit.Refresh
        
        .btnsave.SkinPath = App.Path & "\Skins\Blue"
       .btnsave.ForeColor = &H0&
       .btnsave.LoadSkin
        .btnsave.Refresh
End With
End Sub

Private Sub txtpass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
StayOnTop Me, 0
Dim rs As Recordset
Call link
Set rs = db.OpenRecordset("select * from user where id='" & cboid.Text & "' and password='" & txtpass.Text & "'")
With rs
    If .RecordCount < 1 Then
        h = MsgBox("Invalid password or account doesnt exist " + vbCrLf + "Contact the administrator", vbCritical, PROJ)
        
    Exit Sub
    Else
    f = rs("name")
   Unload frmmain
         'Unload Me
         'Unload frmlog
         frmlog.Hide
         
      frmmain.status.Panels(1).Text = "User:  " + txtnam.Text
               frmmain.mnuuser.Caption = f

        frmmain.Show
    Unload Me
    
End If

End With

End If
End Sub
