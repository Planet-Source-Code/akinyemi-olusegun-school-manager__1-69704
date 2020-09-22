VERSION 5.00
Begin VB.Form frmuser 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3990
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7095
   LinkTopic       =   "Form2"
   ScaleHeight     =   3990
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5400
      Picture         =   "frmuser.frx":0000
      ScaleHeight     =   495
      ScaleWidth      =   480
      TabIndex        =   14
      Top             =   2520
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   600
      Picture         =   "frmuser.frx":0CA2
      ScaleHeight     =   495
      ScaleWidth      =   480
      TabIndex        =   13
      Top             =   1320
      Width           =   480
   End
   Begin VB.TextBox txtdes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2400
      TabIndex        =   11
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox txtid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2400
      TabIndex        =   10
      Top             =   1320
      Width           =   1935
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
      Left            =   2400
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2760
      Width           =   1935
   End
   Begin VB.TextBox txtnam 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   2400
      TabIndex        =   1
      Top             =   1800
      Width           =   3135
   End
   Begin Project1.ctrl_SkinableForm ctrl_SkinableForm 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      Caption         =   "Create User"
      CaptionTop      =   0
   End
   Begin Project1.ctrl_SkinableButton btnsave 
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   3360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Caption         =   "Save"
   End
   Begin Project1.ctrl_SkinableButton btnedit 
      Height          =   495
      Left            =   1920
      TabIndex        =   7
      Top             =   3360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Caption         =   "Edit"
   End
   Begin Project1.ctrl_SkinableButton btndel 
      Height          =   495
      Left            =   3720
      TabIndex        =   8
      Top             =   3360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Caption         =   "Delete"
   End
   Begin Project1.ctrl_SkinableButton btnquit 
      Height          =   375
      Left            =   5280
      TabIndex        =   12
      Top             =   3360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "Quit"
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
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1320
      TabIndex        =   9
      Top             =   1320
      Width           =   735
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
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1320
      TabIndex        =   5
      Top             =   2760
      Width           =   825
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
      Left            =   1320
      TabIndex        =   3
      Top             =   2280
      Width           =   1020
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
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1320
      TabIndex        =   2
      Top             =   1800
      Width           =   495
   End
End
Attribute VB_Name = "frmuser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btndel_Click()
StayOnTop Me, 0
Dim rs As Recordset
Call link
Set rs = db.OpenRecordset("select * from user where id='" & txtid.Text & "'")
With rs
If .RecordCount > 0 Then
.Delete
f = MsgBox("Record deleted", vbInformation, PROJ)

Else
g = MsgBox("Staff ID not recognized", vbCritical, PROJ)
Exit Sub
End If
End With
StayOnTop Me, True
End Sub

Private Sub btndel_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
btndel.Refresh
End Sub

Private Sub btnedit_Click()

StayOnTop Me, False

If Not Tval() = "" Then
    d = MsgBox(Tval(), vbCritical, PROJ)
    Exit Sub
End If

Dim rs As Recordset
Call link
Set rs = db.OpenRecordset("select * from user where id='" & txtid.Text & "'")

With rs
If .RecordCount > 0 Then
.Edit
 rs("id") = txtid.Text
    rs("name") = txtnam.Text
    rs("designation") = txtdes.Text
    rs("password") = txtpass.Text
    
.Update
u = MsgBox("Record Modified", vbInformation, PROJ)

Else
    f = MsgBox("Record does not exist, click on add to add it", vbCritical, PROJ)
Exit Sub
End If
End With
StayOnTop Me, True
End Sub

Private Sub btnedit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
btnedit.Refresh
End Sub

Private Sub btnquit_Click()
Unload Me
End Sub

Private Sub btnquit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
btnquit.Refresh
End Sub

Private Sub btnsave_Click()
StayOnTop Me, False
If Tval() <> "" Then
    d = MsgBox(Tval(), vbCritical, PROJ)
    Exit Sub
End If

Dim rs As Recordset
Call link
Set rs = db.OpenRecordset("select * from user where id='" & txtid.Text & "'")
With rs
If .RecordCount < 1 Then
.AddNew
    rs("id") = txtid.Text
    rs("name") = txtnam.Text
    rs("designation") = txtdes.Text
    rs("password") = txtpass.Text
    
.Update
g = MsgBox("Record Saved", vbInformation, PROJ)

Else
    f = MsgBox("Record exist already, click on edit to edit is", vbCritical, PROJ)
    Exit Sub
End If
    

End With

StayOnTop Me, True
End Sub

Function Tval() As String
If txtid = "" Then
    Tval = "Enter staff Id"
    Exit Function
Else
Tval = ""
End If
    
    
If txtpass = "" Then
    Tval = "Enter staff password"
    Exit Function
Else
Tval = ""
End If
    
End Function
Function CheckStaff(ID As String) As Boolean
Dim rs As Recordset
Call link
Set rs = db.OpenRecordset("select * from staff where ID='" & ID & "'")
With rs
    If .RecordCount > 0 Then
        CheckStaff = True
    Else
        CheckStaff = False
    End If



End With

End Function

Private Sub btnsave_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
btnsave.Refresh
End Sub

Private Sub Form_Load()
Me.Top = Val(Screen.Height - Me.Height) / 2
Me.Left = Val(Screen.Width - Me.Width) / 2
On Local Error Resume Next
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
        
        
        .btnsave.SkinPath = App.Path & "\Skins\Deco"
        .btnsave.ForeColor = &H0&
        .btnsave.LoadSkin
        .btnsave.Refresh
         
         .btnedit.SkinPath = App.Path & "\Skins\Deco"
        .btnedit.ForeColor = &H0&
        .btnedit.LoadSkin
        .btnedit.Refresh
        
         .btndel.SkinPath = App.Path & "\Skins\Deco"
        .btndel.ForeColor = &H0&
        .btndel.LoadSkin
        .btndel.Refresh
        
        
        .btnquit.SkinPath = App.Path & "\Skins\Deco"
     .btnquit.ForeColor = &H0&
        .btnquit.LoadSkin
       .btnquit.Refresh
End With
End Sub

Private Sub txtid_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
StayOnTop Me, 0
    If CheckStaff(txtid.Text) = False Then
    p = MsgBox("You need to be a staff to use this application " + vbCrLf + "Close this application now", vbCritical, PROJ)
    Unload Me
    Me.Show
    Exit Sub
Else
Dim rs As Recordset
Call link
Set rs = db.OpenRecordset("select * from staff where id='" & txtid.Text & "'")
With rs
    txtnam.Text = rs("name")
    txtdes.Text = rs("Designation")
End With

End If
StayOnTop Me, True
End If
End Sub

Private Sub txtid_LostFocus()
StayOnTop Me, 0
If CheckStaff(txtid.Text) = False Then
    p = MsgBox("You need to be a staff to use this application " + vbCrLf + "Close this application now", vbCritical, PROJ)
    Unload Me
    Me.Show
    Exit Sub
Else
Dim rs As Recordset
Call link
Set rs = db.OpenRecordset("select * from staff where id='" & txtid.Text & "'")
With rs
    txtnam.Text = rs("name")
        txtdes.Text = rs("Designation")

End With

End If
StayOnTop Me, True
End Sub
