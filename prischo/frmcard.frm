VERSION 5.00
Begin VB.Form frmcard 
   BorderStyle     =   0  'None
   ClientHeight    =   6720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7770
   LinkTopic       =   "Form2"
   ScaleHeight     =   6720
   ScaleWidth      =   7770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstca 
      Height          =   2595
      ItemData        =   "frmcard.frx":0000
      Left            =   2520
      List            =   "frmcard.frx":0002
      TabIndex        =   22
      Top             =   3120
      Width           =   735
   End
   Begin VB.ListBox lstscore 
      Height          =   2595
      ItemData        =   "frmcard.frx":0004
      Left            =   1800
      List            =   "frmcard.frx":0006
      TabIndex        =   21
      Top             =   3120
      Width           =   735
   End
   Begin VB.ListBox lstCourse 
      Height          =   2595
      ItemData        =   "frmcard.frx":0008
      Left            =   240
      List            =   "frmcard.frx":000A
      TabIndex        =   20
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox txtca 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   6000
      TabIndex        =   19
      ToolTipText     =   "Press enter to add to list"
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox txtexa 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   4440
      TabIndex        =   17
      Top             =   2280
      Width           =   855
   End
   Begin VB.ComboBox cbosub 
      Height          =   315
      ItemData        =   "frmcard.frx":000C
      Left            =   1080
      List            =   "frmcard.frx":0019
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   2280
      Width           =   2055
   End
   Begin VB.ComboBox cboterm 
      Height          =   315
      ItemData        =   "frmcard.frx":0033
      Left            =   3000
      List            =   "frmcard.frx":0040
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   1680
      Width           =   1095
   End
   Begin VB.ComboBox cbocla 
      Height          =   315
      ItemData        =   "frmcard.frx":005A
      Left            =   840
      List            =   "frmcard.frx":007C
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1080
      Width           =   1095
   End
   Begin VB.ComboBox cboid 
      Height          =   315
      Left            =   3000
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txtnam 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1080
      Width           =   2415
   End
   Begin VB.TextBox txtmod 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1680
      Width           =   1095
   End
   Begin Project1.ctrl_SkinableButton btnsave 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   6120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      Caption         =   "Save"
   End
   Begin Project1.ctrl_SkinableForm ctrl_SkinableForm 
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1508
      Caption         =   "Student Performance"
      CaptionTop      =   0
   End
   Begin Project1.ctrl_SkinableButton btnquit 
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   6120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      Caption         =   "Quit"
   End
   Begin Project1.ctrl_SkinableButton btnrem 
      Height          =   375
      Left            =   3360
      TabIndex        =   26
      Top             =   3120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      Caption         =   "Remove Subject"
   End
   Begin Project1.ctrl_SkinableButton ctrl_SkinableButton1 
      Height          =   375
      Left            =   3360
      TabIndex        =   27
      Top             =   3600
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      Caption         =   "Remove Score"
   End
   Begin Project1.ctrl_SkinableButton ctrl_SkinableButton2 
      Height          =   375
      Left            =   3360
      TabIndex        =   28
      Top             =   4080
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      Caption         =   "Remove CA"
   End
   Begin VB.Label Label12 
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   "CA"
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
      Height          =   255
      Left            =   2640
      TabIndex        =   25
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Label11 
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   "Score"
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
      Height          =   255
      Left            =   1920
      TabIndex        =   24
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   "Subject"
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
      Height          =   255
      Left            =   360
      TabIndex        =   23
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C.A"
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
      Left            =   5640
      TabIndex        =   18
      Top             =   2280
      Width           =   315
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Examination:"
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
      Left            =   3240
      TabIndex        =   16
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subject:"
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
      TabIndex        =   15
      Top             =   2280
      Width           =   720
   End
   Begin VB.Label Label5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   4560
      TabIndex        =   13
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Term:"
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
      Left            =   2400
      TabIndex        =   11
      Top             =   1680
      Width           =   495
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
      Left            =   240
      TabIndex        =   10
      Top             =   1080
      Width           =   465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID:"
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
      Left            =   2520
      TabIndex        =   9
      Top             =   1080
      Width           =   270
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
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
      Left            =   4320
      TabIndex        =   8
      Top             =   1080
      Width           =   555
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mode:"
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
      TabIndex        =   7
      Top             =   1680
      Width           =   540
   End
End
Attribute VB_Name = "frmcard"
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

Private Sub btnrem_Click()
StayOnTop Me, 0
Dim a As Integer
a = MsgBox("The selected Course will be removed from the list", vbOKCancel + vbExclamation)
If a = vbCancel Then
 Exit Sub
Else
 If lstCourse.ListIndex <> -1 Then
 lstCourse.RemoveItem (lstCourse.ListIndex)
Else
 MsgBox "You did not select a course"
End If
End If
StayOnTop Me, True
End Sub

Private Sub btnrem_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
btnrem.Refresh
End Sub

Function TestNull() As String
If cboid.Text = "" Then
        TestNull = "Student ID is missing"
        Exit Function
Else
        TestNull = ""
End If

If cboterm.Text = "" Then
        TestNull = "Choose a term"
        Exit Function
Else
        TestNull = ""
End If

If cbocla.Text = "" Then
        TestNull = "Specify the class before continuing"
        
        Exit Function
Else
        TestNull = ""
End If



End Function
Private Sub btnsave_Click()
On Error Resume Next
StayOnTop Me, False

If TestNull() <> "" Then
        m = MsgBox(TestNull(), vbOKCancel, PROJ)
               Exit Sub
End If

Dim rs As Recordset
Call link
Set rs = db.OpenRecordset("select * from GP where class='" & cbocla.Text & "' and ID='" & cboid.Text & "' and term='" & cboterm.Text & "'")
With rs
If .RecordCount = 0 Then

.AddNew
rs("ID") = cboid.Text
rs("subject1") = (lstCourse.List(0))
rs("subject2") = (lstCourse.List(1))
rs("subject3") = (lstCourse.List(2))
rs("subject4") = (lstCourse.List(3))
rs("subject5") = (lstCourse.List(4))
rs("subject6") = (lstCourse.List(5))
rs("subject7") = (lstCourse.List(6))
rs("NoofCourses") = lstCourse.ListCount

rs("Examination1") = Val(lstscore.List(0))
rs("Examination2") = Val(lstscore.List(1))
rs("Examination3") = Val(lstscore.List(2))
rs("Examination4") = Val(lstscore.List(3))
rs("Examination5") = Val(lstscore.List(4))
rs("Examination6") = Val(lstscore.List(5))
rs("Examination7") = Val(lstscore.List(6))

rs("CA1") = Val(lstca.List(0))
rs("CA2") = Val(lstca.List(1))
rs("CA3") = Val(lstca.List(2))
rs("CA4") = Val(lstca.List(3))
rs("CA5") = Val(lstca.List(4))
rs("CA6") = Val(lstca.List(5))
rs("CA7") = Val(lstca.List(6))


rs("term") = cboterm.Text
rs("class") = cbocla.Text
.Update
fol = MsgBox("Record saved", vbInformation, PROJ)
Else
Call EDitA
End If
End With

StayOnTop Me, True



End Sub
Sub EDitA()
Dim rs As Recordset
Call link
Set rs = db.OpenRecordset("select * from GP where class='" & cbocla.Text & "' and ID='" & cboid.Text & "' and term='" & cboterm.Text & "'")
With rs
If .RecordCount > 0 Then
.Edit
rs("ID") = cboid.Text
rs("subject1") = lstCourse.List(0)
rs("subject2") = lstCourse.List(1)
rs("subject3") = lstCourse.List(2)
rs("subject4") = lstCourse.List(3)
rs("subject5") = lstCourse.List(4)
rs("subject5") = lstCourse.List(5)
rs("subject7") = lstCourse.List(6)

rs("NoofCourses") = lstCourse.ListCount

rs("Examination1") = Val(lstscore.List(0))
rs("Examination2") = Val(lstscore.List(1))
rs("Examination3") = Val(lstscore.List(2))
rs("Examination4") = Val(lstscore.List(3))
rs("Examination5") = Val(lstscore.List(4))
rs("Examination6") = Val(lstscore.List(5))
rs("Examination7") = Val(lstscore.List(6))

rs("CA1") = Val(lstca.List(0))
rs("CA2") = Val(lstca.List(1))
rs("CA3") = Val(lstca.List(2))
rs("CA4") = Val(lstca.List(3))
rs("CA5") = Val(lstca.List(4))
rs("CA6") = Val(lstca.List(5))
rs("CA7") = Val(lstca.List(6))

rs("term") = cboterm.Text
rs("class") = cbocla.Text
.Update
End If
End With
End Sub


Private Sub btnsave_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
btnsave.Refresh
End Sub

Private Sub cbosub_Click()
lstCourse.AddItem (cbosub.Text)
End Sub

Private Sub ctrl_SkinableButton1_Click()
StayOnTop Me, 0
Dim a As Integer
a = MsgBox("The selected Score will be removed from the list", vbOKCancel + vbExclamation)
If a = vbCancel Then
 Exit Sub
Else
 If lstscore.ListIndex <> -1 Then
 lstscore.RemoveItem (lstscore.ListIndex)
Else
 MsgBox "You did not select a Score"
End If
End If
StayOnTop Me, True
End Sub

Private Sub ctrl_SkinableButton1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
ctrl_SkinableButton1.Refresh
End Sub

Private Sub ctrl_SkinableButton2_Click()
StayOnTop Me, 0
Dim a As Integer
a = MsgBox("The selected Course unit will be removed from the list", vbOKCancel + vbExclamation)
If a = vbCancel Then
 Exit Sub
Else
 If lstca.ListIndex <> -1 Then
 lstca.RemoveItem (lstca.ListIndex)
Else
 MsgBox "You did not select a course Unit"
End If
End If
StayOnTop Me, True
End Sub

Private Sub ctrl_SkinableButton2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
ctrl_SkinableButton2.Refresh
End Sub

Private Sub Form_Load()
Me.Top = Val(Screen.Height - Me.Height) / 2
Me.Left = Val(Screen.Width - Me.Width) / 2
inito


StayOnTop Me, True
End Sub
Private Sub cbocla_Click()
Dim rs As Recordset
Call link
Set rs = db.OpenRecordset("select * from studinfo where class='" & cbocla.Text & "'")
With rs
While Not .EOF

cboid.AddItem rs("ID")

.MoveNext
Wend


End With
cbosub.Clear
GetCourselist

End Sub
Sub GetCourselist()
Dim rs As Recordset
Call link
Set rs = db.OpenRecordset("select * from courses where class='" & cbocla.Text & "'")
With rs
While Not .EOF
        cbosub.AddItem rs("subject")
    .MoveNext
Wend

End With
End Sub

Private Sub cboid_Click()

y = Year(Now)
Label5.Caption = y

Dim rs, rs1, rs2 As Recordset
Call link
Set rs2 = db.OpenRecordset("select * from year")
With rs2
d = rs2("year")
End With

Set rs = db.OpenRecordset("select * from studinfo where ID='" & cboid.Text & "'")
With rs
txtnam.Text = rs("name")
txtmod = rs("mode")

End With




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
        
        .ctrl_SkinableButton1.SkinPath = App.Path & "\Skins\Deco"
        .ctrl_SkinableButton1.ForeColor = &H0&
        .ctrl_SkinableButton1.LoadSkin
        .ctrl_SkinableButton1.Refresh
        
        .ctrl_SkinableButton2.SkinPath = App.Path & "\Skins\Deco"
        .ctrl_SkinableButton2.ForeColor = &H0&
        .ctrl_SkinableButton2.LoadSkin
        .ctrl_SkinableButton2.Refresh
        
       .btnrem.SkinPath = App.Path & "\Skins\Deco"
        .btnrem.ForeColor = &H0&
      .btnrem.LoadSkin
        .btnrem.Refresh
        
        .btnsave.SkinPath = App.Path & "\Skins\Deco"
        .btnsave.ForeColor = &H0&
      .btnsave.LoadSkin
       .btnsave.Refresh

End With
End Sub

Private Sub txtca_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
                Case Asc("0") To Asc("9")
                Case Asc(".")
                    KeyAscii = IIf(Index = 5 Or Index = 1, 0, KeyAscii)
                Case KeyAscii = vbKeyBack
                
        
        
Case 13

lstca.AddItem txtca.Text
Case Else
                    KeyAscii = 0
End Select
End Sub

Private Sub txtexa_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
                Case Asc("0") To Asc("9")
                Case Asc(".")
                    KeyAscii = IIf(Index = 5 Or Index = 1, 0, KeyAscii)
                Case KeyAscii = vbKeyBack
                Case Else
                    KeyAscii = 0
        End Select
        
End Sub

Private Sub txtexa_LostFocus()
lstscore.AddItem txtexa.Text
End Sub

Private Sub txtxa_Change()

End Sub

Private Sub txtxa_KeyPress(KeyAscii As Integer)


End Sub

