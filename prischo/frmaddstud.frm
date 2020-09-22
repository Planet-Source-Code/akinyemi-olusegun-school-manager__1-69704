VERSION 5.00
Begin VB.Form frmaddstud 
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   FillColor       =   &H80000007&
   FillStyle       =   3  'Vertical Line
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtdiet 
      Enabled         =   0   'False
      Height          =   855
      Left            =   8880
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   40
      Top             =   5400
      Width           =   2415
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00CECECE&
      Caption         =   "Special Diet ?"
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
      Height          =   495
      Left            =   6720
      TabIndex        =   39
      Top             =   5400
      Width           =   1935
   End
   Begin VB.TextBox txtsick 
      Enabled         =   0   'False
      Height          =   855
      Left            =   8880
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   38
      Top             =   4440
      Width           =   2415
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00CECECE&
      Caption         =   "Special Medication or Sickness"
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
      Height          =   855
      Left            =   6720
      TabIndex        =   37
      Top             =   4320
      Width           =   1935
   End
   Begin VB.ComboBox cbodisability 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmaddstud.frx":0000
      Left            =   8880
      List            =   "frmaddstud.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   36
      Top             =   3840
      Width           =   2415
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00CECECE&
      Caption         =   "Any Disability ?"
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
      Height          =   375
      Left            =   6720
      TabIndex        =   35
      Top             =   3840
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   195
      Left            =   9000
      TabIndex        =   34
      Top             =   720
      Visible         =   0   'False
      Width           =   855
   End
   Begin Project1.ctrl_SkinableButton btnload 
      Height          =   375
      Left            =   8040
      TabIndex        =   31
      Top             =   2040
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "Load from File "
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   6600
      ScaleHeight     =   1425
      ScaleWidth      =   1305
      TabIndex        =   30
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox txtnam 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   3240
      TabIndex        =   22
      Top             =   1440
      Width           =   2775
   End
   Begin VB.TextBox txtid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   960
      TabIndex        =   21
      Top             =   1440
      Width           =   1335
   End
   Begin VB.ComboBox cbosex 
      Height          =   315
      ItemData        =   "frmaddstud.frx":0024
      Left            =   1080
      List            =   "frmaddstud.frx":002E
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtage 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   3240
      TabIndex        =   19
      Top             =   1920
      Width           =   1455
   End
   Begin VB.ComboBox cbocla 
      Height          =   315
      ItemData        =   "frmaddstud.frx":0040
      Left            =   1200
      List            =   "frmaddstud.frx":0062
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   2400
      Width           =   1095
   End
   Begin VB.ComboBox cbomod 
      Height          =   315
      ItemData        =   "frmaddstud.frx":00D4
      Left            =   3240
      List            =   "frmaddstud.frx":00DE
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txtemail 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   2040
      TabIndex        =   15
      Top             =   5760
      Width           =   2175
   End
   Begin VB.TextBox txtpar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   1320
      TabIndex        =   10
      Top             =   3720
      Width           =   3135
   End
   Begin VB.TextBox txtadd 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H00C00000&
      Height          =   885
      Left            =   1320
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   9
      Top             =   4200
      Width           =   3135
   End
   Begin VB.TextBox txttel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   2040
      TabIndex        =   8
      Top             =   5280
      Width           =   2175
   End
   Begin Project1.ctrl_SkinableButton btnsave 
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   7080
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      Caption         =   "Save"
   End
   Begin Project1.ctrl_SkinableForm ctrl_SkinableForm 
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1575
      _ExtentX        =   1508
      _ExtentY        =   1296
      Caption         =   "Candidate Registration Form"
      CaptionTop      =   0
   End
   Begin Project1.ctrl_SkinableButton btnedit 
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   7080
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      Caption         =   "Edit"
   End
   Begin Project1.ctrl_SkinableButton btndel 
      Height          =   495
      Left            =   4920
      TabIndex        =   3
      Top             =   7080
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      Caption         =   "Delete"
   End
   Begin Project1.ctrl_SkinableButton ctrl_SkinableButton1 
      Height          =   495
      Left            =   12000
      TabIndex        =   4
      Top             =   4800
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      Caption         =   "Edit"
   End
   Begin Project1.ctrl_SkinableButton btncle 
      Height          =   495
      Left            =   7080
      TabIndex        =   5
      Top             =   7080
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      Caption         =   "Clear"
   End
   Begin Project1.ctrl_SkinableButton btnclearer 
      Height          =   375
      Left            =   9840
      TabIndex        =   32
      Top             =   2040
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "Clear "
   End
   Begin Project1.ctrl_SkinableButton btnquit 
      Height          =   495
      Left            =   9240
      TabIndex        =   42
      Top             =   7080
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      Caption         =   "Quit"
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Medical Info"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   8160
      TabIndex        =   41
      Top             =   3360
      Width           =   1275
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Height          =   3135
      Left            =   6480
      Top             =   3240
      Width           =   5055
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Browse for student picture"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   8280
      TabIndex        =   33
      Top             =   1440
      Width           =   2685
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Height          =   1935
      Left            =   6480
      Top             =   1080
      Width           =   5055
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Personal Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1920
      TabIndex        =   29
      Top             =   1200
      Width           =   1755
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
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   2640
      TabIndex        =   28
      Top             =   1440
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
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   600
      TabIndex        =   27
      Top             =   1440
      Width           =   270
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sex"
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
      Height          =   195
      Left            =   600
      TabIndex        =   26
      Top             =   1920
      Width           =   330
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Age"
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
      Height          =   195
      Left            =   2640
      TabIndex        =   25
      Top             =   1920
      Width           =   345
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
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   600
      TabIndex        =   24
      Top             =   2400
      Width           =   465
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mode"
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
      Height          =   195
      Left            =   2640
      TabIndex        =   23
      Top             =   2400
      Width           =   480
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Height          =   1935
      Left            =   480
      Top             =   1080
      Width           =   5655
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email Address"
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
      Height          =   195
      Left            =   600
      TabIndex        =   16
      Top             =   5760
      Width           =   1200
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Parent/Guardian"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   2400
      TabIndex        =   14
      Top             =   3360
      Width           =   1425
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Name"
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
      Height          =   195
      Left            =   600
      TabIndex        =   13
      Top             =   3720
      Width           =   555
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      Height          =   195
      Left            =   600
      TabIndex        =   12
      Top             =   4200
      Width           =   690
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Telephone No"
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
      Height          =   195
      Left            =   600
      TabIndex        =   11
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Height          =   3135
      Left            =   480
      Top             =   3240
      Width           =   5655
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To Delete,supply the Student ID and click delete"
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
      Left            =   1920
      TabIndex        =   7
      Top             =   7680
      Width           =   4185
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Press F1 for Help"
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
      Left            =   1920
      TabIndex        =   6
      Top             =   8160
      Width           =   1485
   End
End
Attribute VB_Name = "frmaddstud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btncle_Click()
Unload Me
Me.Show
End Sub

Private Sub btncle_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
btncle.Refresh
End Sub

Private Sub btnclearer_Click()
On Error Resume Next
pic.Picture = LoadPicture("")

End Sub

Private Sub btnclearer_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
btnclearer.Refresh
End Sub

Private Sub btndel_Click()
StayOnTop Me, 0
Dim rs As Recordset
Call link
Set rs = db.OpenRecordset("select * from studinfo where ID='" & txtid.Text & "'")
With rs
    If .RecordCount > 0 Then
        .Delete
    Else
        MS = MsgBox("Cant delete unexisting record")

    End If
End With
StayOnTop Me, True
End Sub

Private Sub btndel_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
btndel.Refresh
End Sub

Private Sub btnedit_Click()
StayOnTop Me, 0
If Tval() <> "" Then
            m = MsgBox(Tval(), vbOKCancel + vbCritical, PROJ)
            Exit Sub
    End If
Dim rs As Recordset
Call link
Set rs = db.OpenRecordset("select * from studinfo where ID='" & txtid.Text & "'")
With rs
    If .RecordCount > 0 Then
        .Edit
            rs("ID") = txtid.Text
            rs("Name") = txtnam.Text
            rs("Age") = txtage.Text
            rs("Mode") = cbomod.Text
            rs("Sex") = cbosex.Text
            rs("Class") = cbocla.Text
            rs("Parent") = txtpar.Text
            rs("Address") = txtadd.Text
            rs("Telephone") = txttel.Text
            rs("Enrolled") = "No"
            
            rs("email") = txtemail.Text
            rs("pics") = Text1.Text
            
            If Check1.Value = 1 Then
                rs("disability") = cbodisability.Text
            End If
            
            If Check2.Value = 1 Then
                rs("sickness") = Trim(txtsick.Text)
            End If
            
            If Check3.Value = 1 Then
                rs("diet") = Trim(txtdiet.Text)
            End If
            
        .Update
            f = MsgBox("Student Record Edited", vbInformation, PROJ)
            'Unload Me
            'frmaddstud.Show
    Else
        MS = MsgBox("Student Record does not Exist" + vbCrLf + "Click on Edit to edit student info", vbCritical, PROJ)
        Exit Sub
        End If
End With
StayOnTop Me, True
End Sub

Private Sub btnedit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
btnedit.Refresh
End Sub

Private Sub btnload_Click()
On Error Resume Next
Me.WindowState = 1

frmplaceholder.Show

StayOnTop Me, True
End Sub

Private Sub btnload_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
btnload.Refresh
End Sub

Private Sub btnquit_Click()
Unload Me
End Sub

Private Sub btnquit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
btnquit.Refresh
End Sub

Private Sub btnsave_Click()
StayOnTop Me, 0
 If Tval() <> "" Then
            m = MsgBox(Tval(), vbOKCancel + vbCritical, PROJ)
            Exit Sub
    End If

Dim rs As Recordset
Call link
Set rs = db.OpenRecordset("select * from studinfo where ID='" & txtid.Text & "'")
With rs
    If .RecordCount < 1 Then
        .AddNew
            rs("ID") = txtid.Text
            rs("Name") = txtnam.Text
            rs("Age") = txtage.Text
            rs("Mode") = cbomod.Text
            rs("Sex") = cbosex.Text
            rs("Class") = cbocla.Text
            rs("Parent") = txtpar.Text
            rs("Address") = txtadd.Text
            rs("Telephone") = txttel.Text
            rs("Enrolled") = "No"
            rs("email") = txtemail.Text
            rs("pics") = Text1.Text
            
            If Check1.Value = 1 Then
                rs("disability") = cbodisability.Text
            End If
            
            If Check2.Value = 1 Then
                rs("sickness") = Trim(txtsick.Text)
            End If
            
            If Check3.Value = 1 Then
                rs("diet") = Trim(txtdiet.Text)
            End If
            
        .Update
            f = MsgBox("Student Record Saved", vbInformation, PROJ)
            'Unload Me
            'frmaddstud.Show
    Else
        MS = MsgBox("Student Record Exist Already" + vbCrLf + "Click on Edit to edit student info", vbCritical, PROJ)
        Exit Sub
        End If
End With
StayOnTop Me, True
End Sub

Private Sub btnsave_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
btnsave.Refresh
End Sub

Private Sub Command1_Click()

   
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command3_Click()
End Sub
Function Tval() As String
If txtid = "" Then
Tval = "Supply the candidate ID to continue"
Exit Function
Else
Tval = ""

End If

If txtnam = "" Then
Tval = "Candidate name field is empty"
Exit Function
Else
Tval = ""
End If

If txtage = "" Then
Tval = "Age field is empty"
Exit Function
Else
Tval = ""
End If

If cbosex.Text = "" Then
Tval = "Choose a sex for your candidate"
Exit Function
Else
Tval = ""
End If

If cbomod.Text = "" Then
Tval = "Enter mode of study"
Exit Function
Else
Tval = ""
End If

If txtpar = "" Then
Tval = "Parent and guardian name is important"
Exit Function
Else
Tval = ""
End If


If txtadd = "" Then
Tval = "Suppy the address now"
Exit Function
Else
Tval = ""
End If

If txttel = "" Then
Tval = "Telephone number needed"
Exit Function
Else
Tval = ""
End If

If cbocla.Text = "" Then
Tval = "Choose a class"
Exit Function
Else
Tval = ""
End If

End Function
Private Sub Command4_Click()
Unload Me

End Sub

Private Sub Command5_Click()

End Sub

Private Sub ctrl_ListObject_Click(Index As Integer)
Select Case Index
Case 0
Me.WindowState = 1
Case 1
Me.WindowState = 2
Case 2
Me.Hide
Case 3
Unload Me

End Select
End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
    cbodisability.Enabled = True
Else
    cbodisability.Enabled = False
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
    txtsick.Enabled = True
Else
    txtsick.Enabled = False
End If

End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
    txtdiet.Enabled = True
Else
txtdiet.Enabled = False
End If
End Sub

Private Sub Form_Load()
inito
StayOnTop Me, True
Dim rs As Recordset
Call link
Set rs = db.OpenRecordset("select * from studinfo")
With rs
While Not .EOF
If Len(str(rs.RecordCount)) = 1 Then
p = "0000" + Trim(str(rs.RecordCount + 1))
End If

If Len(str(rs.RecordCount)) = 2 Then
p = "000" + Trim(str(rs.RecordCount + 1))
End If

If Len(str(rs.RecordCount)) = 3 Then
p = "00" + Trim(str(rs.RecordCount + 1))
End If

If Len(str(rs.RecordCount)) = 4 Then
p = "0" + Trim(str(rs.RecordCount + 1))
End If


.MoveNext
txtid.Text = p

Wend
End With
End Sub

Sub inito()
With frmaddstud
.ctrl_SkinableForm.SkinPath = App.Path & "\Skins\Deco"
        .ctrl_SkinableForm.BackColor = &HCECECE
        .ctrl_SkinableForm.CaptionTop = 300
        .ctrl_SkinableForm.CaptionColor = &H0&
        Call Me.ctrl_SkinableForm.LoadSkin(Me)
    
         
         .btnsave.SkinPath = App.Path & "\Skins\Deco"
        .btnsave.ForeColor = &H0&
        .btnsave.LoadSkin
        .btnsave.Refresh
        
        .btnload.SkinPath = App.Path & "\Skins\TreasureChest"
       .btnload.ForeColor = &HFFFFFF
        .btnload.LoadSkin
        .btnload.Refresh
       
       
        .btnclearer.SkinPath = App.Path & "\Skins\TreasureChest"
       .btnclearer.ForeColor = &HFFFFFF
        .btnclearer.LoadSkin
        .btnclearer.Refresh
        
         .btnedit.SkinPath = App.Path & "\Skins\Deco"
        .btnedit.ForeColor = &H0&
        .btnedit.LoadSkin
        .btnedit.Refresh
        
        .btnquit.SkinPath = App.Path & "\Skins\Deco"
        .btnquit.ForeColor = &H0&
        .btnquit.LoadSkin
        .btnquit.Refresh
        
         .btndel.SkinPath = App.Path & "\Skins\Deco"
        .btndel.ForeColor = &H0&
        .btndel.LoadSkin
        .btndel.Refresh
        
        .btncle.SkinPath = App.Path & "\Skins\Deco"
        .btncle.ForeColor = &H0&
        .btncle.LoadSkin
        .btncle.Refresh
        
       
End With

End Sub

Private Sub txtage_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
                Case Asc("0") To Asc("9")
                Case Asc(".")
                    KeyAscii = IIf(Index = 5 Or Index = 1, 0, KeyAscii)
                Case KeyAscii = vbKeyBack
                Case Else
                    KeyAscii = 0
        End Select
End Sub

Private Sub txttel_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
                Case Asc("0") To Asc("9")
                Case Asc(".")
                    KeyAscii = IIf(Index = 5 Or Index = 1, 0, KeyAscii)
                Case KeyAscii = vbKeyBack
                Case Else
                    KeyAscii = 0
        End Select
End Sub
