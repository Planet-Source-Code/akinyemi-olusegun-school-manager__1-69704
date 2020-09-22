VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmclevel 
   Caption         =   "List of Fees"
   ClientHeight    =   5475
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4860
   LinkTopic       =   "Form2"
   ScaleHeight     =   5475
   ScaleWidth      =   4860
   StartUpPosition =   3  'Windows Default
   Begin Project1.ctrl_SkinableButton btnadd 
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   5040
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      Caption         =   "Add New"
   End
   Begin Project1.ctrl_SkinableButton btnquit 
      Height          =   495
      Left            =   3480
      TabIndex        =   1
      Top             =   5040
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      Caption         =   "Quit"
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4275
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   7541
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
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
         Object.Width           =   5468
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Amount"
         Object.Width           =   1764
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
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
            Picture         =   "frmclevel.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblmessage 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   4800
      Width           =   45
   End
End
Attribute VB_Name = "frmclevel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ctrl_SkinableButton2_Click()

End Sub

Private Sub btnadd_Click()
StayOnTop Me, 0
frmaddyearlevel.Show
Unload Me

End Sub

Private Sub btnmodify_Click()
lblmessage.Caption = "Modify"
End Sub

Private Sub btnquit_Click()
Unload Me

End Sub

Private Sub Form_Load()
inito
StayOnTop Me, True
Me.Top = Val(Screen.Height - Me.Height) / 2
Me.Left = Val(Screen.Width - Me.Width) / 2
Filler
End Sub
Sub Filler()
Dim rs As Recordset
Call link
Set rs = db.OpenRecordset("select * from classlevel")
With rs
Do While Not .EOF
With listview1
            .ListItems.add , , rs("class"), 1, 1
            .ListItems(.ListItems.Count).SubItems(1) = Format$(IIf(IsNull(rs("Amount")), "", rs("Amount")), "#,##0.00")
End With
           .MoveNext
           
        Loop
        End With
End Sub

Sub inito()
With Me
.btnadd.SkinPath = App.Path & "\Skins\Holograph"
        .btnadd.ForeColor = &HFFFFFF
        .btnadd.LoadSkin
        .btnadd.Refresh
        
        .btnquit.SkinPath = App.Path & "\Skins\Holograph"
        .btnquit.ForeColor = &HFFFFFF
        .btnquit.LoadSkin
        .btnquit.Refresh
        
End With
End Sub
