VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmcourses 
   BorderStyle     =   0  'None
   Caption         =   "Courses setup"
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5535
   LinkTopic       =   "Form2"
   ScaleHeight     =   7200
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtdes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   1440
      TabIndex        =   9
      Top             =   1680
      Width           =   3735
   End
   Begin VB.TextBox txtsub 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   3000
      TabIndex        =   7
      Top             =   1080
      Width           =   2175
   End
   Begin VB.ComboBox cbocla 
      Height          =   315
      ItemData        =   "frmcourses.frx":0000
      Left            =   960
      List            =   "frmcourses.frx":0022
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1080
      Width           =   1095
   End
   Begin Project1.ctrl_SkinableButton btnsave 
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   2160
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      Caption         =   "Add"
   End
   Begin Project1.ctrl_SkinableForm ctrl_SkinableForm 
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1508
      Caption         =   "Course Setup"
      CaptionTop      =   0
   End
   Begin Project1.ctrl_SkinableButton btnquit 
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   6600
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      Caption         =   "Quit"
   End
   Begin MSComctlLib.ListView listview1 
      Height          =   3795
      Left            =   240
      TabIndex        =   3
      ToolTipText     =   "Double click to Edit"
      Top             =   2640
      Width           =   5025
      _ExtentX        =   8864
      _ExtentY        =   6694
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "Imagelist1"
      SmallIcons      =   "Imagelist1"
      ForeColor       =   4194304
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Class"
         Object.Width           =   4762
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Subject"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Description"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList Imagelist1 
      Left            =   5160
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmcourses.frx":0094
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
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
      Left            =   360
      TabIndex        =   8
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
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
      Height          =   195
      Left            =   2280
      TabIndex        =   6
      Top             =   1080
      Width           =   660
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
      Left            =   360
      TabIndex        =   5
      Top             =   1080
      Width           =   465
   End
End
Attribute VB_Name = "frmcourses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnquit_Click()
Unload Me
End Sub

Private Sub btnsave_Click()
StayOnTop Me, 0
Dim rs As Recordset
Call link
Set rs = db.OpenRecordset("select * from courses where class='" & cbocla.Text & "' and subject='" & txtsub.Text & "'")

With rs
If .RecordCount < 1 Then
.AddNew
rs("description") = txtdes.Text
rs("class") = cbocla.Text
rs("subject") = txtsub.Text
.Update
Else
EDitA
End If
End With
rs.Close
With ListView1
            .ListItems.add , , cbocla.Text, 1, 1
            .ListItems(.ListItems.Count).SubItems(1) = txtsub.Text
            .ListItems(.ListItems.Count).SubItems(2) = txtdes.Text
End With
cbocla.Text = ""
txtsub.Text = ""
txtdes.Text = ""
StayOnTop Me, True
End Sub
Sub EDitA()
Dim rs As Recordset
Call link
Set rs = db.OpenRecordset("select * from courses where class='" & cbocla.Text & "' and subject='" & txtsub.Text & "'")

With rs
If .RecordCount > 0 Then
.Edit
rs("description") = txtdes.Text
rs("class") = cbocla.Text
rs("subject") = txtsub.Text
.Update
End If
End With
rs.Close
End Sub

Private Sub btnsave_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
btnsave.Refresh
End Sub

Private Sub Form_Load()
Me.Top = Val(Screen.Height - Me.Height) / 2
Me.Left = Val(Screen.Width - Me.Width) / 2
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
        
        .btnquit.SkinPath = App.Path & "\Skins\Deco"
        .btnquit.ForeColor = &H0&
        .btnquit.LoadSkin
        .btnquit.Refresh

End With

End Sub


