VERSION 5.00
Begin VB.Form frmplaceholder 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Picture Chooser"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5190
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2760
      TabIndex        =   9
      Top             =   3240
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   3960
      TabIndex        =   7
      Top             =   4320
      Width           =   1095
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   480
      ScaleHeight     =   1305
      ScaleWidth      =   1665
      TabIndex        =   6
      Top             =   3240
      Width           =   1695
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Height          =   1785
      Left            =   2760
      Pattern         =   "*.gif"
      TabIndex        =   2
      Top             =   1200
      Width           =   2295
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      Height          =   1890
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   2175
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Preview Selection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   300
      Left            =   2520
      TabIndex        =   8
      Top             =   3840
      Width           =   2130
   End
   Begin VB.Label Label3 
      Caption         =   "File Name"
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Folder Name"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Disk"
      Height          =   195
      Left            =   480
      TabIndex        =   3
      Top             =   360
      Width           =   315
   End
End
Attribute VB_Name = "frmplaceholder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

 frmaddstud.Text1.Text = Text1
frmaddstud.pic.Picture = LoadPicture(Text1)
frmaddstud.WindowState = 2
frmstaffae.Text1 = Text1
frmstaffae.pic.Picture = LoadPicture(Text1)
frmstaffae.WindowState = 1
Unload Me

End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
Dim st As String
st = Dir1.Path + "\" + File1.filename
Text1 = st
pic.Picture = LoadPicture(Text1)
End Sub

Private Sub Form_Load()
Me.Top = Val(Screen.Height - Me.Height) / 2
Me.Left = Val(Screen.Width - Me.Width) / 2
StayOnTop Me, True
Dir1.Path = "c:\prischo\picture\"
End Sub
