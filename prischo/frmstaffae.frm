VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmstaffae 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7620
   LinkTopic       =   "Form2"
   ScaleHeight     =   9000
   ScaleWidth      =   7620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   1080
      ScaleHeight     =   1425
      ScaleWidth      =   1305
      TabIndex        =   28
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   195
      Left            =   1560
      TabIndex        =   27
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSMask.MaskEdBox date1 
      Height          =   255
      Left            =   2400
      TabIndex        =   23
      Top             =   5040
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtnam 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   3720
      TabIndex        =   8
      Top             =   3600
      Width           =   3135
   End
   Begin VB.TextBox txtid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   1440
      TabIndex        =   7
      Top             =   3600
      Width           =   1335
   End
   Begin VB.ComboBox cbosex 
      Height          =   315
      ItemData        =   "frmstaffae.frx":0000
      Left            =   1560
      List            =   "frmstaffae.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox txtage 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   3720
      TabIndex        =   5
      Top             =   4080
      Width           =   1455
   End
   Begin VB.ComboBox cboqua 
      Height          =   315
      ItemData        =   "frmstaffae.frx":001C
      Left            =   2280
      List            =   "frmstaffae.frx":002F
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   4560
      Width           =   1095
   End
   Begin VB.ComboBox cbodes 
      Height          =   315
      ItemData        =   "frmstaffae.frx":004D
      Left            =   4800
      List            =   "frmstaffae.frx":0063
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   4560
      Width           =   1935
   End
   Begin VB.TextBox txtadd 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H00C00000&
      Height          =   885
      Left            =   2040
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   5520
      Width           =   4815
   End
   Begin VB.TextBox txttel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   5040
      TabIndex        =   1
      Top             =   5040
      Width           =   1695
   End
   Begin Project1.ctrl_SkinableButton btnsave 
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   6720
      Width           =   1935
      _extentx        =   3413
      _extenty        =   873
      caption         =   "Save"
   End
   Begin Project1.ctrl_SkinableForm ctrl_SkinableForm 
      Height          =   735
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   735
      _extentx        =   1296
      _extenty        =   1296
      caption         =   "Staff Registration Form"
      captiontop      =   0
   End
   Begin Project1.ctrl_SkinableButton btnedit 
      Height          =   495
      Left            =   4080
      TabIndex        =   10
      Top             =   6720
      Width           =   1935
      _extentx        =   3413
      _extenty        =   873
      caption         =   "Edit"
   End
   Begin Project1.ctrl_SkinableButton btndel 
      Height          =   495
      Left            =   1560
      TabIndex        =   11
      Top             =   7200
      Width           =   1935
      _extentx        =   3413
      _extenty        =   873
      caption         =   "Delete"
   End
   Begin Project1.ctrl_SkinableButton ctrl_SkinableButton1 
      Height          =   495
      Left            =   12000
      TabIndex        =   12
      Top             =   4800
      Width           =   1935
      _extentx        =   3413
      _extenty        =   873
      caption         =   "Edit"
   End
   Begin Project1.ctrl_SkinableButton btncle 
      Height          =   495
      Left            =   4080
      TabIndex        =   13
      Top             =   7200
      Width           =   1935
      _extentx        =   3413
      _extenty        =   873
      caption         =   "Clear"
   End
   Begin Project1.ctrl_SkinableButton btnload 
      Height          =   375
      Left            =   2520
      TabIndex        =   29
      Top             =   2040
      Width           =   1575
      _extentx        =   2778
      _extenty        =   661
      caption         =   "Load from File "
   End
   Begin Project1.ctrl_SkinableButton btnclearer 
      Height          =   375
      Left            =   4800
      TabIndex        =   30
      Top             =   2040
      Width           =   1575
      _extentx        =   2778
      _extenty        =   661
      caption         =   "Clear "
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
      Left            =   2760
      TabIndex        =   31
      Top             =   1440
      Width           =   2685
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Height          =   1935
      Left            =   960
      Top             =   1080
      Width           =   6135
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Staff PIS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   3360
      TabIndex        =   26
      Top             =   3240
      Width           =   900
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Height          =   3495
      Left            =   960
      Top             =   3120
      Width           =   6135
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To Delete,supply the Staff ID and click delete"
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
      TabIndex        =   25
      Top             =   7800
      Width           =   3930
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date Employed"
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
      Left            =   1080
      TabIndex        =   24
      Top             =   5040
      Width           =   1290
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
      Left            =   3000
      TabIndex        =   22
      Top             =   3600
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
      Left            =   1080
      TabIndex        =   21
      Top             =   3600
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
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1080
      TabIndex        =   20
      Top             =   4080
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
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3000
      TabIndex        =   19
      Top             =   4080
      Width           =   345
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Qualification"
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
      Left            =   1080
      TabIndex        =   18
      Top             =   4560
      Width           =   1080
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
      Left            =   3720
      TabIndex        =   17
      Top             =   4560
      Width           =   1020
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
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1080
      TabIndex        =   16
      Top             =   5520
      Width           =   690
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone Number"
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
      Left            =   3720
      TabIndex        =   15
      Top             =   5040
      Width           =   1260
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
      Left            =   720
      TabIndex        =   14
      Top             =   8280
      Width           =   1485
   End
End
Attribute VB_Name = "frmstaffae"
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
pic.Picture = LoadPicture("")
End Sub

Private Sub btnclearer_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
btnclearer.Refresh
End Sub

Private Sub btndel_Click()
StayOnTop Me, 0
Dim rs As Recordset
Call link
Set rs = db.OpenRecordset("select * from staff where ID='" & txtid.Text & "'")
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
Set rs = db.OpenRecordset("select * from staff where ID='" & txtid.Text & "'")
With rs
    If .RecordCount > 0 Then
        .Edit
            rs("ID") = txtid.Text
            rs("Name") = txtnam.Text
            rs("Age") = txtage.Text
            rs("Qualification") = cboqua.Text
            rs("Sex") = cbosex.Text
            rs("Designation") = cbodes.Text
            rs("Date_Employed") = date1.Text
            rs("Address") = txtadd.Text
            rs("Telephone") = txttel.Text
            rs("pics") = Text1.Text
        .Update
            f = MsgBox("Staff Record Edited", vbInformation, PROJ)
            'Unload Me
            'frmaddstud.Show
    Else
        MS = MsgBox("Staff Record does not Exist" + vbCrLf + "Click on Edit to edit student info", vbCritical, PROJ)
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

frmstaffpics.Show

StayOnTop Me, True
End Sub
Private Sub btnload_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
btnload.Refresh
End Sub

Private Sub btnsave_Click()
StayOnTop Me, 0
 If Tval() <> "" Then
            m = MsgBox(Tval(), vbOKCancel + vbCritical, PROJ)
            Exit Sub
    End If

Dim rs As Recordset
Call link
Set rs = db.OpenRecordset("select * from staff where ID='" & txtid.Text & "'")
With rs
    If .RecordCount < 1 Then
        .AddNew
            rs("ID") = txtid.Text
            rs("Name") = txtnam.Text
            rs("Age") = txtage.Text
            rs("Qualification") = cboqua.Text
            rs("Sex") = cbosex.Text
            rs("Designation") = cbodes.Text
            rs("Date_Employed") = date1.Text
            rs("Address") = txtadd.Text
            rs("Telephone") = txttel.Text
            rs("pics") = Text1.Text
        .Update
            f = MsgBox("Staff Record Saved", vbInformation, PROJ)
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
Tval = "Supply the Staff ID to continue"
Exit Function
Else
Tval = ""

End If

If txtnam = "" Then
Tval = "Staff name field is empty"
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
Tval = "Choose a sex for the staff"
Exit Function
Else
Tval = ""
End If

If cbodes.Text = "" Then
Tval = "Choose a designation"
Exit Function
Else
Tval = ""
End If

If Not IsDate(date1.Text) Then
Tval = "The date is not correct"
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

If cboqua.Text = "" Then
Tval = "Choose a Qualification"
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

Private Sub Form_Load()
inito
Me.Top = Val(Screen.Height - Me.Height) / 2
Me.Left = Val(Screen.Width - Me.Width) / 2
StayOnTop Me, True
Dim rs As Recordset
Call link
Set rs = db.OpenRecordset("select * from staff")
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
With Me
.ctrl_SkinableForm.SkinPath = App.Path & "\Skins\Deco"
        .ctrl_SkinableForm.BackColor = &HCECECE
        .ctrl_SkinableForm.CaptionTop = 300
        .ctrl_SkinableForm.CaptionColor = &H0&
        Call Me.ctrl_SkinableForm.LoadSkin(Me)

          .btnload.SkinPath = App.Path & "\Skins\TreasureChest"
       .btnload.ForeColor = &HFFFFFF
        .btnload.LoadSkin
        .btnload.Refresh
       
       
        .btnclearer.SkinPath = App.Path & "\Skins\TreasureChest"
       .btnclearer.ForeColor = &HFFFFFF
        .btnclearer.LoadSkin
        .btnclearer.Refresh
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
