VERSION 5.00
Begin VB.Form frmscale 
   BorderStyle     =   0  'None
   ClientHeight    =   5505
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6765
   LinkTopic       =   "Form2"
   ScaleHeight     =   5505
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Project1.ctrl_SkinableButton btnsave1 
      Height          =   375
      Left            =   360
      TabIndex        =   14
      Top             =   4920
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Caption         =   "Save"
   End
   Begin VB.ComboBox cbodesi 
      Height          =   315
      ItemData        =   "frmscale.frx":0000
      Left            =   3120
      List            =   "frmscale.frx":0019
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   4320
      Width           =   1815
   End
   Begin VB.TextBox txtamo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   4800
      TabIndex        =   10
      Top             =   3720
      Width           =   1695
   End
   Begin VB.TextBox txtall 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   1680
      TabIndex        =   8
      Top             =   3720
      Width           =   1695
   End
   Begin Project1.ctrl_SkinableForm ctrl_SkinableForm1 
      Height          =   735
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      CaptionTop      =   0
   End
   Begin Project1.ctrl_SkinableButton btnsave 
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   2280
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      Caption         =   "Save"
   End
   Begin VB.TextBox txtsal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   4800
      TabIndex        =   2
      Top             =   1800
      Width           =   1695
   End
   Begin VB.ComboBox cbodes 
      Height          =   315
      ItemData        =   "frmscale.frx":0066
      Left            =   1320
      List            =   "frmscale.frx":007F
      TabIndex        =   0
      Top             =   1800
      Width           =   1815
   End
   Begin Project1.ctrl_SkinableButton btnedit 
      Height          =   375
      Left            =   2520
      TabIndex        =   15
      Top             =   4920
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Caption         =   "Edit"
   End
   Begin Project1.ctrl_SkinableButton btnquit 
      Height          =   375
      Left            =   4680
      TabIndex        =   16
      Top             =   4920
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Caption         =   "Close"
   End
   Begin Project1.ctrl_SkinableButton btnedit1 
      Height          =   375
      Left            =   3960
      TabIndex        =   17
      Top             =   2280
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      Caption         =   "Edit"
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Categories of staff entitle to it"
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
      Left            =   240
      TabIndex        =   13
      Top             =   4320
      Width           =   2550
   End
   Begin VB.Label Label5 
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
      Left            =   3840
      TabIndex        =   11
      Top             =   3720
      Width           =   645
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Allowance Type"
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
      Left            =   240
      TabIndex        =   9
      Top             =   3720
      Width           =   1365
   End
   Begin VB.Line Line5 
      BorderColor     =   &H000000FF&
      BorderWidth     =   5
      X1              =   4200
      X2              =   6600
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Allowances"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   2520
      TabIndex        =   7
      Top             =   3240
      Width           =   1365
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000FF&
      BorderWidth     =   5
      X1              =   0
      X2              =   2400
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Basic Salary"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   2520
      TabIndex        =   6
      Top             =   1080
      Width           =   1500
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   5
      X1              =   0
      X2              =   2400
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Basic Salary"
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
      Left            =   3480
      TabIndex        =   3
      Top             =   1800
      Width           =   1065
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Designation"
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
      Left            =   240
      TabIndex        =   1
      Top             =   1800
      Width           =   1020
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   5
      X1              =   0
      X2              =   6600
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   5
      X1              =   4200
      X2              =   6600
      Y1              =   1200
      Y2              =   1200
   End
End
Attribute VB_Name = "frmscale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnedit_Click()
Dim rs As Recordset
StayOnTop Me, 0
Call link
Set rs = db.OpenRecordset("select * from scale2 where type='" & txtall.Text & "' and Category='" & cbodesi.Text & "'")
    With rs
        If .RecordCount > 0 Then
            .Edit
                rs("type") = txtall.Text
                rs("amount") = txtamo.Text
                rs("Category") = cbodesi.Text
             .Update
        d = MsgBox("Record Edited")
        Else
            d = MsgBox("Record does not exist, click on save to save it", vbInformation + vbCritical, PROJ)
            Exit Sub
        End If
        
    End With
txtall = ""
txtamo = ""
cbodesi = ""
StayOnTop Me, True
End Sub

Private Sub btnedit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
btnedit.Refresh
End Sub

Private Sub btnedit1_Click()
Dim rs As Recordset
StayOnTop Me, 0
Call link
Set rs = db.OpenRecordset("select * from scale1 where Designation='" & cbodes.Text & "'")
    With rs
    
        If .RecordCount > 0 Then
            .Edit
            
                rs("Designation") = cbodes.Text
                    rs("salary") = txtsal.Text
            .Update
    
    s = MsgBox("Information Edited", vbInformation, PROJ)
    
    Else
    f = MsgBox("Record does not exist, click on save to save it", vbOKCancel, PROJ)
    Exit Sub
    End If
    End With
    cbodes.Text = ""
    txtsal = ""
    StayOnTop Me, True
End Sub

Private Sub btnedit1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
btnedit1.Refresh
End Sub

Private Sub btnquit_Click()
Unload Me
End Sub

Private Sub btnquit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
btnquit.Refresh
End Sub

Private Sub btnsave_Click()

Dim rs As Recordset
StayOnTop Me, 0
Call link
Set rs = db.OpenRecordset("select * from scale1 where Designation='" & cbodes.Text & "'")
    With rs
    
        If .RecordCount < 1 Then
            .AddNew
            
                rs("Designation") = cbodes.Text
                    rs("salary") = txtsal.Text
            .Update
    
    s = MsgBox("Information saved", vbInformation, PROJ)
    
    Else
    f = MsgBox("Record exist already, click on edit to edit it", vbOKCancel, PROJ)
    Exit Sub
    End If
    End With
    cbodes.Text = ""
    txtsal = ""
    StayOnTop Me, True
End Sub

Private Sub btnsave_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
btnsave.Refresh
End Sub

Private Sub btnsave1_Click()
Dim rs As Recordset
StayOnTop Me, 0
Call link
Set rs = db.OpenRecordset("select * from scale2 where type='" & txtall.Text & "' and Category='" & cbodesi.Text & "'")
    With rs
        If .RecordCount < 1 Then
            .AddNew
                rs("type") = txtall.Text
                rs("amount") = txtamo.Text
                rs("Category") = cbodesi.Text
             .Update
        d = MsgBox("Record saved")
        Else
            d = MsgBox("Record exist already, click on edit to edit it", vbInformation + vbCritical, PROJ)
            Exit Sub
        End If
        
    End With
txtall = ""
txtamo = ""
cbodesi = ""
StayOnTop Me, True
End Sub

Private Sub btnsave1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
btnsave1.Refresh
End Sub

Private Sub Form_Load()
Me.Top = Val(Screen.Height - Me.Height) / 2
Me.Left = Val(Screen.Width - Me.Width) / 2
inito
StayOnTop Me, True

End Sub
Sub inito()
With Me
.ctrl_SkinableForm1.SkinPath = App.Path & "\Skins\Deco"
        .ctrl_SkinableForm1.BackColor = &HCECECE
        .ctrl_SkinableForm1.CaptionTop = 300
        .ctrl_SkinableForm1.CaptionColor = &H0&
        Call Me.ctrl_SkinableForm1.LoadSkin(Me)

.btnquit.SkinPath = App.Path & "\Skins\Deco"
        .btnquit.ForeColor = &H0&
        .btnquit.LoadSkin
        .btnquit.Refresh
        
        .btnedit.SkinPath = App.Path & "\Skins\Deco"
       .btnedit.ForeColor = &H0&
       .btnedit.LoadSkin
        .btnedit.Refresh
        
        .btnedit1.SkinPath = App.Path & "\Skins\Deco"
       .btnedit1.ForeColor = &H0&
       .btnedit1.LoadSkin
        .btnedit1.Refresh
        
        .btnsave.SkinPath = App.Path & "\Skins\Deco"
        .btnsave.ForeColor = &H0&
      .btnsave.LoadSkin
       .btnsave.Refresh


 .btnsave1.SkinPath = App.Path & "\Skins\Deco"
        .btnsave1.ForeColor = &H0&
      .btnsave1.LoadSkin
       .btnsave1.Refresh

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

Private Sub txtsal_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
                Case Asc("0") To Asc("9")
                Case Asc(".")
                    KeyAscii = IIf(Index = 5 Or Index = 1, 0, KeyAscii)
                Case KeyAscii = vbKeyBack
                Case Else
                    KeyAscii = 0
        End Select
End Sub
