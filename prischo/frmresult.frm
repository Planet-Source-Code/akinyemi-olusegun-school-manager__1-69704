VERSION 5.00
Begin VB.Form frmresult 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   2670
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   2670
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboid 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1680
      Width           =   1095
   End
   Begin VB.ComboBox cboterm 
      Height          =   315
      ItemData        =   "frmresult.frx":0000
      Left            =   1080
      List            =   "frmresult.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.ComboBox cbocla 
      Height          =   315
      ItemData        =   "frmresult.frx":0027
      Left            =   3120
      List            =   "frmresult.frx":0049
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1080
      Width           =   1095
   End
   Begin Project1.ctrl_SkinableForm ctrl_SkinableForm 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   975
      _extentx        =   1720
      _extenty        =   1508
      caption         =   "Student Result"
      captiontop      =   0
   End
   Begin Project1.ctrl_SkinableButton btnquit 
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   2040
      Width           =   1575
      _extentx        =   2778
      _extenty        =   661
      caption         =   "Close"
   End
   Begin Project1.ctrl_SkinableButton btnprint 
      Height          =   375
      Left            =   2760
      TabIndex        =   8
      Top             =   1560
      Width           =   1575
      _extentx        =   2778
      _extenty        =   661
      caption         =   "Print"
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Class:"
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
      TabIndex        =   6
      Top             =   1080
      Width           =   525
   End
   Begin VB.Label Label1 
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
      Left            =   360
      TabIndex        =   5
      Top             =   1080
      Width           =   495
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
      Left            =   360
      TabIndex        =   4
      Top             =   1680
      Width           =   270
   End
End
Attribute VB_Name = "frmresult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnprint_Click()
On Error Resume Next
StayOnTop Me, False
Dim rs As Recordset
Dim rs1 As Recordset
Call link
Set rs1 = db.OpenRecordset("select * from studinfo where class='" & cbocla.Text & "' and ID='" & cboid.Text & "'")
With rs1
cla = rs1("class")
nam = rs1("name")

End With


Set rs = db.OpenRecordset("select * from GP where class='" & cbocla.Text & "' and ID='" & cboid.Text & "' and term='" & cboterm.Text & "'")
With rs
    If .RecordCount < 1 Then
        sd = MsgBox("Result for this student/term has not been prepared yet", vbCritical, "Task aborted")
        Exit Sub
    Else
         dRept.rsCommand1.Filter = "ID ='" & (cboid.Text) & "' and class='" & (cbocla.Text) & "' and term='" & (cboterm.Text) & "'"
         reportcard.Sections(1).Controls(2).Caption = nam
          reportcard.Sections(1).Controls(4).Caption = cla
          reportcard.Sections(1).Controls(6).Caption = rs("term")

        reportcard.Show
    End If
End With
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

Private Sub cbocla_Click()
cboid.Clear
Dim rs As Recordset
Call link
    Set rs = db.OpenRecordset("select * from studinfo where class='" & cbocla.Text & "'")
    With rs
        While Not .EOF
                cboid.AddItem rs("id")
                .MoveNext
        Wend
    
    End With
    
End Sub

Private Sub Form_Load()
Me.Top = Val(Screen.Height - Me.Height) / 2
Me.Left = Val(Screen.Width - Me.Width) / 2
inito
'StayOnTop Me, True
End Sub
Sub inito()
With Me
.ctrl_SkinableForm.SkinPath = App.Path & "\Skins\Deco"
        .ctrl_SkinableForm.BackColor = &HCECECE
        .ctrl_SkinableForm.CaptionTop = 300
        .ctrl_SkinableForm.CaptionColor = &H0&
        Call Me.ctrl_SkinableForm.LoadSkin(Me)
        
        .btnprint.SkinPath = App.Path & "\Skins\Deco"
        .btnprint.ForeColor = &H0&
        .btnprint.LoadSkin
        .btnprint.Refresh
        
        
        .btnquit.SkinPath = App.Path & "\Skins\Deco"
        .btnquit.ForeColor = &H0&
        .btnquit.LoadSkin
        .btnquit.Refresh
End With

End Sub
