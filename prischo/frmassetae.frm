VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmassetae 
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form2"
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtdep 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   4680
      TabIndex        =   28
      Top             =   5160
      Width           =   975
   End
   Begin VB.TextBox Txtcos 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   4080
      TabIndex        =   26
      Top             =   4560
      Width           =   1335
   End
   Begin VB.TextBox Txtapre 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   7080
      TabIndex        =   24
      Top             =   4560
      Width           =   1695
   End
   Begin VB.TextBox Txtreg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   3840
      Width           =   2175
   End
   Begin VB.TextBox txtprod 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   5400
      TabIndex        =   6
      Top             =   3360
      Width           =   2175
   End
   Begin VB.TextBox txtrem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H00C00000&
      Height          =   885
      Left            =   6960
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Top             =   5160
      Width           =   1935
   End
   Begin VB.ComboBox cboloc 
      Height          =   315
      ItemData        =   "frmassetae.frx":0000
      Left            =   4560
      List            =   "frmassetae.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2760
      Width           =   1095
   End
   Begin VB.ComboBox cbotype 
      Height          =   315
      ItemData        =   "frmassetae.frx":0004
      Left            =   6600
      List            =   "frmassetae.frx":0014
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox txtid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   3840
      TabIndex        =   2
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox txtdec 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   4560
      TabIndex        =   1
      Top             =   2280
      Width           =   3495
   End
   Begin MSMask.MaskEdBox date1 
      Height          =   255
      Left            =   7320
      TabIndex        =   0
      Top             =   2760
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin Project1.ctrl_SkinableButton btnsave 
      Height          =   495
      Left            =   1680
      TabIndex        =   7
      Top             =   6480
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      Caption         =   "Save"
   End
   Begin Project1.ctrl_ListObject ctrl_ListObject 
      Height          =   2895
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   5106
   End
   Begin Project1.ctrl_SkinableForm ctrl_SkinableForm 
      Height          =   1095
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   1575
      _ExtentX        =   1508
      _ExtentY        =   1296
      Caption         =   "Assets Form"
      CaptionTop      =   0
   End
   Begin Project1.ctrl_SkinableButton btnedit 
      Height          =   495
      Left            =   3840
      TabIndex        =   10
      Top             =   6480
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      Caption         =   "Edit"
   End
   Begin Project1.ctrl_SkinableButton btndel 
      Height          =   495
      Left            =   6000
      TabIndex        =   11
      Top             =   6480
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      Caption         =   "Delete"
   End
   Begin Project1.ctrl_SkinableButton ctrl_SkinableButton1 
      Height          =   495
      Left            =   12000
      TabIndex        =   12
      Top             =   4800
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      Caption         =   "Edit"
   End
   Begin Project1.ctrl_SkinableButton btncle 
      Height          =   495
      Left            =   8160
      TabIndex        =   13
      Top             =   6480
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      Caption         =   "Clear"
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To Delete,supply the Asset ID and click delete"
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
      TabIndex        =   30
      Top             =   7800
      Width           =   3990
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Depreciation"
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
      TabIndex        =   29
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cost"
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
      TabIndex        =   27
      Top             =   4560
      Width           =   390
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Appreciation"
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
      Left            =   5880
      TabIndex        =   25
      Top             =   4560
      Width           =   1080
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Registration Number"
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
      TabIndex        =   23
      Top             =   3840
      Width           =   1740
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
      TabIndex        =   21
      Top             =   8160
      Width           =   1485
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product Number"
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
      TabIndex        =   20
      Top             =   3360
      Width           =   1380
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remark"
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
      Left            =   5880
      TabIndex        =   19
      Top             =   5160
      Width           =   660
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Location"
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
      TabIndex        =   18
      Top             =   2760
      Width           =   750
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
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
      Left            =   5880
      TabIndex        =   17
      Top             =   1560
      Width           =   435
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
      Left            =   3480
      TabIndex        =   16
      Top             =   1560
      Width           =   270
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
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
      TabIndex        =   15
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date Purchased"
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
      Left            =   5760
      TabIndex        =   14
      Top             =   2760
      Width           =   1380
   End
End
Attribute VB_Name = "frmassetae"
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

Private Sub btndel_Click()
StayOnTop Me, 0
Dim rs As Recordset
Call link
Set rs = db.OpenRecordset("select * from asset where ID='" & txtid.Text & "'")
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
Set rs = db.OpenRecordset("select * from asset where ID='" & txtid.Text & "'")
With rs
    If .RecordCount > 0 Then
        .Edit
             rs("ID") = txtid.Text
             rs("type") = cbotype.Text
            'rs("Name") = txtnam.Text
            rs("Description") = txtdec.Text
            rs("Location") = cboloc.Text
            rs("Cost") = Txtcos.Text
            rs("Appreciation") = Txtapre
            rs("Depreciation") = txtdep.Text
            rs("Remark") = txtrem.Text
            rs("Date_Acquired") = date1.Text
            rs("Product_No") = txtprod.Text
            If Txtreg.Locked = False Then
            rs("Registration") = Txtreg.Text
            Else
            rs("Registration") = "N/A"
            End If
        .Update
            f = MsgBox("Asset Record Edited", vbInformation, PROJ)
            'Unload Me
            'frmaddstud.Show
    Else
        MS = MsgBox("Asset Record does not Exist" + vbCrLf + "Click on save button to save asset info", vbCritical, PROJ)
        Exit Sub
        End If
End With
StayOnTop Me, True
End Sub

Private Sub btnedit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
btnedit.Refresh
End Sub

Private Sub btnsave_Click()
StayOnTop Me, 0
 If Tval() <> "" Then
            m = MsgBox(Tval(), vbOKCancel + vbCritical, PROJ)
            Exit Sub
    End If

Dim rs As Recordset
Call link
Set rs = db.OpenRecordset("select * from asset where ID='" & txtid.Text & "'")
With rs
    If .RecordCount < 1 Then
        .AddNew
            rs("ID") = txtid.Text
           ' rs("Name") = txtnam.Text
           rs("type") = cbotype.Text
            rs("Description") = txtdec.Text
            rs("Location") = cboloc.Text
            rs("Cost") = Txtcos.Text
            rs("Appreciation") = Txtapre
            rs("Depreciation") = txtdep.Text
            rs("Remark") = txtrem.Text
            rs("Date_Acquired") = date1.Text
            rs("Product_No") = txtprod.Text
            If Txtreg.Locked = False Then
            rs("Registration") = Txtreg.Text
            Else
            rs("Registration") = "N/A"
            End If
        .Update
            f = MsgBox("Asset Record Saved", vbInformation, PROJ)
            'Unload Me
            'frmaddstud.Show
    Else
        MS = MsgBox("Asset Record Exist Already" + vbCrLf + "Click on Edit to edit Asset info", vbCritical, PROJ)
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
Tval = "Supply the Asset ID to continue"
Exit Function
Else
Tval = ""

End If

If txtdec = "" Then
Tval = "Asset description field is empty"
Exit Function
Else
Tval = ""
End If

If Me.Txtapre = "" Then
Tval = "Appreciation field is empty"
Exit Function
Else
Tval = ""
End If

If cbotype.Text = "" Then
Tval = "Choose a Type for the Asset"
Exit Function
Else
Tval = ""
End If

If Me.Txtcos = "" Then
Tval = "Enter a cost"
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


If txtrem = "" Then
Tval = "Enter a remark specifying the condition of the asset"
Exit Function
Else
Tval = ""
End If

If Me.txtdep = "" Then
Tval = "Enter depreciation"
Exit Function
Else
Tval = ""
End If

If Me.txtprod.Text = "" Then
Tval = "Enter a production Number"
Exit Function
Else
Tval = ""
End If

If Me.cboloc.Text = "" Then
Tval = "Choose a location"
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



Private Sub cbotype_Click()
If cbotype.Text = "Vehicle" Then
Txtreg.Locked = 0
Else
Txtreg = ""

Txtreg.Locked = 1
End If
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
StayOnTop Me, True
Dim rs As Recordset
Call link
Set rs = db.OpenRecordset("select * from asset")
With rs
While Not .EOF
If Len(str(rs.RecordCount)) = 1 Or Len(str(rs.RecordCount)) = 0 Then
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
        
        .ctrl_ListObject.SkinPath = App.Path & "\Skins\Deco"
        .ctrl_ListObject.ForeColor = &H0&
        .ctrl_ListObject.MouseMoveColor = &H0&
        .ctrl_ListObject.MouseDownColor = &H0&
         .ctrl_ListObject.DrawMenu
         
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
        
         ctrl_ListObject.AddItem "Minimize"
         ctrl_ListObject.AddItem "Maximize"
         ctrl_ListObject.AddItem "Hide Form"
         ctrl_ListObject.AddItem "close form"
End With

End Sub


Private Sub Txtapre_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
                Case Asc("0") To Asc("9")
                Case Asc(".")
                    KeyAscii = IIf(Index = 5 Or Index = 1, 0, KeyAscii)
                Case KeyAscii = vbKeyBack
                Case Else
                    KeyAscii = 0
        End Select
End Sub

Private Sub Txtapre_LostFocus()
If Not IsNumeric(Txtapre) Then
d = MsgBox("Apprecaition must be number only", vbOKCancel, PROJ)
    Txtapre.SetFocus
End If
End Sub

Private Sub Txtcos_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
                Case Asc("0") To Asc("9")
                Case Asc(".")
                    KeyAscii = IIf(Index = 5 Or Index = 1, 0, KeyAscii)
                Case KeyAscii = vbKeyBack
                Case Else
                    KeyAscii = 0
        End Select
End Sub

Private Sub Txtcos_LostFocus()
If Not IsNumeric(Txtcos) Then
d = MsgBox("Cost must be number only", vbOKCancel, PROJ)
    Txtcos.SetFocus
End If
End Sub

Private Sub txtdep_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
                Case Asc("0") To Asc("9")
                Case Asc(".")
                    KeyAscii = IIf(Index = 5 Or Index = 1, 0, KeyAscii)
                Case KeyAscii = vbKeyBack
                Case Else
                    KeyAscii = 0
        End Select
End Sub

Private Sub txtdep_LostFocus()
If Not IsNumeric(txtdep) Then
d = MsgBox("Depreciation must be number only", vbOKCancel, PROJ)
    txtdep.SetFocus
End If
End Sub
