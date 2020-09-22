VERSION 5.00
Begin VB.Form frmstaffpics 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Select Staff picture"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   4830
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   600
      TabIndex        =   5
      Top             =   360
      Width           =   1455
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      Height          =   1890
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   2175
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Height          =   1785
      Left            =   2400
      Pattern         =   "*.gif"
      TabIndex        =   3
      Top             =   1200
      Width           =   2295
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   120
      ScaleHeight     =   1305
      ScaleWidth      =   1665
      TabIndex        =   2
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   4440
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2400
      TabIndex        =   0
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Disk"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   360
      Width           =   315
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Folder Name"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "File Name"
      Height          =   195
      Left            =   2400
      TabIndex        =   7
      Top             =   960
      Width           =   705
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
      Left            =   2040
      TabIndex        =   6
      Top             =   3720
      Width           =   2130
   End
End
Attribute VB_Name = "frmstaffpics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Hide
 
frmstaffae.Text1 = Text1
frmstaffae.pic.Picture = LoadPicture(Text1)
frmstaffae.WindowState = 0
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

