VERSION 5.00
Begin VB.Form frmgenerate 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   825
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6555
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   825
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   5400
      TabIndex        =   6
      Top             =   240
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2880
      Top             =   120
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generat......"
      Height          =   495
      Left            =   4200
      TabIndex        =   5
      Top             =   240
      Width           =   1095
   End
   Begin VB.ComboBox cbocla 
      Height          =   315
      ItemData        =   "frmgenerate.frx":0000
      Left            =   960
      List            =   "frmgenerate.frx":0022
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.ComboBox cboterm 
      Height          =   315
      ItemData        =   "frmgenerate.frx":0094
      Left            =   2760
      List            =   "frmgenerate.frx":00A1
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
   Begin Project1.ProgressGuage ga 
      Height          =   495
      Left            =   120
      Top             =   240
      Width           =   6255
      _extentx        =   11033
      _extenty        =   873
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Generate result here"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   2130
   End
   Begin VB.Label Label1 
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
      Index           =   1
      Left            =   360
      TabIndex        =   3
      Top             =   360
      Width           =   465
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
      Left            =   2160
      TabIndex        =   2
      Top             =   360
      Width           =   495
   End
End
Attribute VB_Name = "frmgenerate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
StayOnTop Me, 0
ga.Visible = True
Command2.Visible = False
DoSQL
Me.Label1(1).Visible = False
Me.Label4.Visible = False
Me.cbocla.Visible = False
Me.cboterm.Visible = False
Me.Command1.Visible = False
Timer1.Enabled = True
'Unload Me
Label1(0).Caption = "Generating Result wait......"
StayOnTop Me, True
End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()
Me.Top = Val(Screen.Height - Me.Height) / 2
Me.Left = Val(Screen.Width - Me.Width) / 2
ga.Visible = False
StayOnTop Me, True
End Sub
Sub DoSQL()
'On Error GoTo er
Dim rs As Recordset
Call link

'Set rs = db.OpenRecordset("INSERT INTO GP _
' select Sum(Examination1 + CA1) AS To1, Sum(Examination2 + CA2) AS To2, Sum(Examination3 + CA3) AS To3, Sum(Examination4 + CA4) AS To4, Sum(Examination5 + CA5) AS To5, Sum(Examination6 + CA6) AS To6, Sum(Examination7 + CA7) AS To7 from GP where term='" & cboterm.Text & "'
'and class='" & cbocla.Text & "'")

Set rs = db.OpenRecordset("select * from GP where term='" & cboterm.Text & "' and class='" & cbocla.Text & "'")

With rs
While Not .EOF
.Edit
rs("To1") = Val(rs("Examination1") + rs("CA1"))
rs("To2") = Val(rs("Examination2") + rs("CA2"))
rs("To3") = Val(rs("Examination3") + rs("CA3"))
rs("To4") = Val(rs("Examination4") + rs("CA4"))
rs("To5") = Val(rs("Examination5") + rs("CA5"))
rs("To6") = Val(rs("Examination6") + rs("CA6"))
rs("To7") = Val(rs("Examination7") + rs("CA7"))
p = (Val(rs("To1")) + Val(rs("To2")) + Val(rs("To3")) + Val(rs("To4")) + Val(rs("To5")) + Val(rs("To6")) + Val(rs("To7"))) / Val(rs("NoofCourses"))
rs("percentage") = p

If rs("term") = "Third" And Val(p) > 50 Then
rs("status") = "Promoted"
End If

.Update
.MoveNext
Wend
End With
End Sub
Private Sub Timer1_Timer()
If ga.Value = 100 Then
Unload Me
d = MsgBox("Result Generation completed click, click on print result for more", vbOKOnly + vbInformation, PROJ)
Else
ga.Value = ga.Value + (Rnd * 5) * 3
End If
End Sub
