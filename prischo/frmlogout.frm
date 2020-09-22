VERSION 5.00
Begin VB.Form frmlogout 
   ClientHeight    =   2130
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   6045
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   2130
   ScaleWidth      =   6045
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   5160
      Picture         =   "frmlogout.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   7
      Top             =   120
      Width           =   480
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2400
      Top             =   360
   End
   Begin Project1.ProgressGuage pv 
      Height          =   375
      Left            =   240
      Top             =   1320
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   661
   End
   Begin VB.Label perc 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   480
      TabIndex        =   6
      Top             =   1800
      Width           =   45
   End
   Begin VB.Label Label4 
      Caption         =   "Completed"
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label lll 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   2160
      TabIndex        =   4
      Top             =   960
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Shutting down components"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label ll 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   2040
      TabIndex        =   2
      Top             =   480
      Width           =   45
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Shutting down modules"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1650
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Loggin out. Please wait.........."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmlogout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
StayOnTop Me, True
Me.Top = Val(Screen.Height - Me.Height) / 2
Me.Left = Val(Screen.Width - Me.Width) / 2

End Sub

Private Sub Timer1_Timer()
pv.Value = pv.Value + 5
perc.Caption = str(pv.Value) + "%"
If pv.Value = 5 Then
ll.Caption = Mid(frmmain.mnubackup.Caption, 8, Len(frmmain.mnubackup.Caption) - 8)


frmmain.mnubackup.Visible = False
End If


If pv.Value = 10 Then
ll.Caption = frmmain.mnuviewrecord.Caption

frmmain.mnuviewrecord.Visible = False
End If

If pv.Value = 15 Then
ll.Caption = frmmain.mnutransaction.Caption

frmmain.mnutransaction.Visible = False
End If

If pv.Value = 20 Then
ll.Caption = frmmain.mnustudentdfgdfgd.Caption

frmmain.mnustudentdfgdfgd.Visible = False
End If

If pv.Value = 25 Then
ll.Caption = frmmain.mnustudentdfgdfgd.Caption
frmmain.Toolbar1.Enabled = False
frmmain.mnustudentdfgdfgd.Visible = False
End If

If pv.Value = 30 Then
ll.Caption = frmmain.mnureportser.Caption
frmmain.mnuchangeuser.Caption = "Log in"
frmmain.mnureportser.Visible = False
End If

If pv.Value = 35 Then
ll.Caption = frmmain.mnufinance.Caption

frmmain.mnufinance.Visible = False
End If

If pv.Value = 40 Then
ll.Caption = frmmain.mnusearch.Caption

frmmain.mnusearch.Visible = False
End If

If pv.Value = 45 Then
ll.Caption = frmmain.mnucurrentuser.Caption + " account"

frmmain.mnucurrentuser.Visible = False
'Unload Me
End If

If pv.Value = 50 Then
ll.Caption = Mid(frmmain.mnulogout.Caption, 9, Len(frmmain.mnulogout.Caption) - 8)

frmmain.mnulogout.Visible = False
'Unload Me
End If
If pv.Value = 55 Then
ll.Caption = frmmain.mnucreateuser.Caption

frmmain.mnucreateuser.Visible = False
frmmain.status.Panels(1).Text = "User:"

'Unload Me
End If


If pv.Value = 65 Then
lll.Caption = "closing staff modules"
Unload frmstaff
Unload frmstaffae
End If

If pv.Value = 70 Then
lll.Caption = "closing Asset modules"
Unload frmasset
Unload frmassetae
End If

If pv.Value = 75 Then
lll.Caption = "closing student performance modules"
Unload frmcard
Unload frmcourses
Unload frmgenerate
Unload frmresult
End If

If pv.Value = 80 Then
lll.Caption = "closing System GUIs handles"
DoEvents
'ReleaseDC()
frmmain.Label14.Visible = 1

End If

If pv.Value = 100 Then
Unload Me
End If

End Sub
