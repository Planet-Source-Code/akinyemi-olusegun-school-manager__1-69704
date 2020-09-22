VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmback 
   ClientHeight    =   5970
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   4575
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   5970
   ScaleWidth      =   4575
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Text            =   "*.mdb"
      Top             =   4200
      Width           =   2055
   End
   Begin VB.DirListBox Dir1 
      Height          =   990
      Left            =   2520
      TabIndex        =   6
      Top             =   3600
      Width           =   1815
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   720
      TabIndex        =   5
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   2040
      Top             =   3720
   End
   Begin Project1.ProgressGuage pv 
      Height          =   255
      Left            =   0
      Top             =   4680
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   450
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Abort"
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   5400
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   3120
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Backup"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   5400
      Width           =   1695
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   0
      Picture         =   "frmback.frx":0000
      ScaleHeight     =   2625
      ScaleWidth      =   4545
      TabIndex        =   4
      Top             =   480
      Width           =   4575
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Destination File"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   3960
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Folder/directory"
      Height          =   195
      Left            =   2760
      TabIndex        =   9
      Top             =   3360
      Width           =   1110
   End
   Begin VB.Label Label2 
      Caption         =   "Drive"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Backup database file"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   720
      TabIndex        =   3
      Top             =   0
      Width           =   2580
   End
   Begin VB.Label lblstatus 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   4920
      Width           =   4455
   End
End
Attribute VB_Name = "frmback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

On Error GoTo BackupError
Call link
 'Dialog.filename = ""
 
    'Dialog.Filter = "Backup files (*.mdb) |*.mdb|"
    'Dialog.ShowSave
    Timer1.Enabled = True
        If Text1 <> "" Then
        pv.Value = pv.Value + 5
            If FileLen(App.Path + "\data.mdb") > 1210000 And Mid(Text1, 1, 1) = "A" Then
           
                MsgBox "Database size is greater than diskette capacity.", vbOKOnly + vbCritical, PROJ
                lblstatus.Caption = "Backup terminated"
                Exit Sub
            Else
            pv.Value = pv.Value + 10
                     db.Close
               Command2.Enabled = False
                     pv.Value = pv.Value + 5
                  
                    FileCopy (App.Path & "\data.mdb"), Text1.Text
                     pv.Value = pv.Value + 40
                     If pv.Value = 60 Then
                     pv.Value = 100
                     End If
                    MsgBox "Backup Completed", vbOKOnly + vbInformation, PROJ
                    lblstatus.Caption = "Backup completed"
                    ''Set DB = OpenDatabase(App.Path + "\DATABASE\Video.mdb")
                    'Set db = OpenDatabase(App.Path + "\data.mdb", dbDriverComplete, False)
                   'UPDATE THE AUDIT LOG FILE
                    'pv.Value = 100
                    Timer1.Enabled = False
                    Command2.Enabled = 1
                  '  Me.Hide
            End If
        End If
    Exit Sub

BackupError:

If Err.Number = 3420 Then
   Resume Next
End If
MsgBox Err.Description + ", cannot backup at this time, try again later", vbOKOnly + vbCritical, PROJ
Command1.Enabled = False
Command2.Enabled = 1

Exit Sub
End Sub

Private Sub Command2_Click()
Unload Me

End Sub


Function GetFile(driv As DirListBox, txt As String) As String
If Len(driv.Path) = 3 Then
GetFile = driv.Path + txt
Else
GetFile = driv.Path + "\" + txt
End If
End Function
Function GetExtension(st As String) As String
If Right(st, 4) = ".mdb" Then
GetExtension = st
Else
GetExtension = st + ".mdb"
End If
End Function

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
StayOnTop Me, True
Me.Top = Val(Screen.Height - Me.Height) / 2
Me.Left = Val(Screen.Width - Me.Width) / 2
End Sub

Private Sub Text1_LostFocus()
Text1.Text = GetFile(Dir1, Text1.Text)
Text1.Text = GetExtension(Text1.Text)
End Sub

Private Sub Timer1_Timer()
If pv.Value = 100 Then
Timer1.Enabled = False
pv.Value = pv.Value + 5
End If
End Sub
