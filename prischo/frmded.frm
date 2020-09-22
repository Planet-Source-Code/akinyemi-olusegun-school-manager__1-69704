VERSION 5.00
Begin VB.Form frmded 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   6225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4725
   LinkTopic       =   "Form2"
   ScaleHeight     =   6225
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cbomon 
      Height          =   315
      ItemData        =   "frmded.frx":0000
      Left            =   1680
      List            =   "frmded.frx":0028
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   960
      Width           =   1815
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   240
      TabIndex        =   9
      Top             =   3000
      Width           =   3975
   End
   Begin VB.ComboBox cboid 
      Height          =   315
      ItemData        =   "frmded.frx":008E
      Left            =   1680
      List            =   "frmded.frx":0090
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1440
      Width           =   1815
   End
   Begin VB.ComboBox cboded 
      Height          =   315
      ItemData        =   "frmded.frx":0092
      Left            =   1680
      List            =   "frmded.frx":0094
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox txtamo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Top             =   2400
      Width           =   1695
   End
   Begin Project1.ctrl_SkinableForm ctrl_SkinableForm 
      Height          =   855
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1508
      Caption         =   "Deduction"
      CaptionTop      =   0
   End
   Begin Project1.ctrl_SkinableButton btnquit 
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   5640
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "Close"
   End
   Begin Project1.ctrl_SkinableButton btnprint 
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   2400
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Caption         =   "&Add"
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Month"
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
      Left            =   120
      TabIndex        =   11
      Top             =   960
      Width           =   540
   End
   Begin VB.Label Label2 
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
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Deduction Type"
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
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   1365
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
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
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   645
   End
End
Attribute VB_Name = "frmded"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnprint_Click()
StayOnTop Me, 0
On Error Resume Next
Dim rs As Recordset
Call link
Set rs = db.OpenRecordset("select * from deduction2")
    With rs
        .AddNew
            rs("type") = cboded.Text
            rs("id") = cboid.Text
            rs("amount") = txtamo.Text
            rs("month") = cboded.Text
        
        .Update
    
    End With
List1.AddItem cboded.Text + "  " + txtamo.Text

StayOnTop Me, True
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

Private Sub cboded_Click()
On Error Resume Next
Dim rs1 As Recordset
Call link

   Set rs1 = db.OpenRecordset("select * from deduction where type='" & cboded.Text & "'")
           With rs1
           txtamo.Text = rs1("Amount")
                      End With
End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Top = Val(Screen.Height - Me.Height) / 2
Me.Left = Val(Screen.Width - Me.Width) / 2
inito
StayOnTop Me, True


Dim rs As Recordset
Dim rs1 As Recordset
Call link
Set rs = db.OpenRecordset("select * from staff")
    With rs
        While Not .EOF
                cboid.AddItem rs("id")
            .MoveNext
        Wend
    
    End With
    
    
    Set rs1 = db.OpenRecordset("select * from deduction")
        With rs1
        
            While Not .EOF
                    cboded.AddItem rs1("type")
                .MoveNext
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

.btnquit.SkinPath = App.Path & "\Skins\Deco"
        .btnquit.ForeColor = &H0&
        .btnquit.LoadSkin
        .btnquit.Refresh
        
        .btnprint.SkinPath = App.Path & "\Skins\Deco"
        .btnprint.ForeColor = &H0&
        .btnprint.LoadSkin
    .btnprint.Refresh
End With
End Sub
