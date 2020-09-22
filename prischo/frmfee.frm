VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmfee 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4500
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   4500
   StartUpPosition =   3  'Windows Default
   Begin Project1.ctrl_SkinableButton btnupdate 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   4440
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "Add Fee"
   End
   Begin Project1.ctrl_SkinableButton ctrl_SkinableButton2 
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   4440
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "Quit"
   End
   Begin MSComctlLib.ListView ListItems 
      Height          =   3795
      Left            =   120
      TabIndex        =   2
      Top             =   330
      Width           =   4185
      _ExtentX        =   7382
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Fee Name"
         Object.Width           =   4762
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Amount"
         Object.Width           =   2117
      EndProperty
   End
   Begin MSComctlLib.ImageList Imagelist1 
      Left            =   3030
      Top             =   0
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
            Picture         =   "frmfee.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmfee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnupdate_Click()
frmfeeae.Show
Unload Me
End Sub

Private Sub ctrl_SkinableButton2_Click()
Unload Me

End Sub

Sub Filler()
Dim rs As Recordset
Call link
Set rs = db.OpenRecordset("select * from fee")
With rs
    While Not .EOF
    With ListItems
            .ListItems.add , , rs("Fee_Name"), 1, 1
            .ListItems(.ListItems.Count).SubItems(1) = Format$(IIf(IsNull(rs("Amount")), "", rs("Amount")), "#,##0.00")
End With
.MoveNext
    Wend

End With
End Sub
Sub inito()
With Me
.btnupdate.SkinPath = App.Path & "\Skins\Holograph"
        .btnupdate.ForeColor = &HFFFFFF
        .btnupdate.LoadSkin
        .btnupdate.Refresh
        
        .ctrl_SkinableButton2.SkinPath = App.Path & "\Skins\Holograph"
        .ctrl_SkinableButton2.ForeColor = &HFFFFFF
        .ctrl_SkinableButton2.LoadSkin
        .ctrl_SkinableButton2.Refresh
        
End With
End Sub

Private Sub Form_Load()
inito
StayOnTop Me, True
Me.Top = Val(Screen.Height - Me.Height) / 2
Me.Left = Val(Screen.Width - Me.Width) / 2
Filler
End Sub
