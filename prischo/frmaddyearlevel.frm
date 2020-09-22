VERSION 5.00
Begin VB.Form frmaddyearlevel 
   BorderStyle     =   0  'None
   Caption         =   "Add Fee for Class"
   ClientHeight    =   2295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3885
   LinkTopic       =   "Form2"
   ScaleHeight     =   2295
   ScaleWidth      =   3885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cbocla 
      Height          =   315
      ItemData        =   "frmaddyearlevel.frx":0000
      Left            =   1440
      List            =   "frmaddyearlevel.frx":0022
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   840
      Width           =   1815
   End
   Begin Project1.ctrl_SkinableButton btnupdate 
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   1800
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "Update"
   End
   Begin VB.TextBox txtamo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   1320
      Width           =   1815
   End
   Begin Project1.ctrl_SkinableForm ctrl_SkinableForm 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      Caption         =   "Class Fee"
      BackColor       =   255
      ForeColor       =   16777215
      CaptionTop      =   0
      CaptionColor    =   16777215
   End
   Begin Project1.ctrl_SkinableButton btnquit 
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   1800
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "Quit"
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Class"
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
      Left            =   720
      TabIndex        =   5
      Top             =   840
      Width           =   465
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
      Left            =   720
      TabIndex        =   2
      Top             =   1320
      Width           =   645
   End
End
Attribute VB_Name = "frmaddyearlevel"
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

Private Sub btnupdate_Click()
StayOnTop Me, 0
If txtamo = "" Or cbocla.Text = "" Then
    f = MsgBox("Suppy the neccessary information before clicking on update")
    Exit Sub
End If

Dim rs As Recordset
Call link
Set rs = db.OpenRecordset("select * from classlevel where class='" & cbocla.Text & "'")
With rs
    If .RecordCount > 0 Then
        .Edit
        rs("class") = cbocla.Text
        rs("amount") = txtamo.Text
        .Update
        l = MsgBox("Record updated", vbOKCancel, PROJ)
       
    Else
        f = MsgBox("Record does not exist in the file" + vbcrl + "Do you want to create ? yes/no", vbYesNo, PROJ)
        If f = vbYes Then
            Call add
        Else
        End If

    End If
End With
 rs.Close
 StayOnTop Me, True
End Sub
Sub add()
Dim rs As Recordset
Call link

Set rs = db.OpenRecordset("select * from classlevel where class='" & cbocla.Text & "'")
With rs
    If .RecordCount < 1 Then
        .AddNew
        rs("class") = cbocla.Text
        rs("amount") = txtamo.Text
        .Update
        l = MsgBox("New Record added", vbOKCancel, PROJ)
    End If
        rs.Close
End With
End Sub
Sub inito()
With frmaddyearlevel
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

Private Sub btnupdate_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
btnupdate.Refresh
End Sub

Private Sub Form_Load()
StayOnTop Me, True
Me.Top = Val(Screen.Height - Me.Height) / 2
Me.Left = Val(Screen.Width - Me.Width) / 2
inito
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
