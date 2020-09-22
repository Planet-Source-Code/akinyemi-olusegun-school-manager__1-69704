VERSION 5.00
Begin VB.Form Frmopen 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   ClientHeight    =   6180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8070
   LinkTopic       =   "Form2"
   ScaleHeight     =   6180
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   7000
      Left            =   6600
      Top             =   2520
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Akinyemi Olusegun                    E-Mail: segzee20002001@yahoo.com Phone:+2348066405658"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1200
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   4740
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "About Macosoft"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   480
      TabIndex        =   1
      Top             =   480
      Width           =   2160
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "School Manager"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   600
      Left            =   2880
      TabIndex        =   0
      Top             =   4560
      Width           =   3705
   End
   Begin VB.Image Image1 
      Height          =   2385
      Left            =   3240
      Picture         =   "openfr.frx":0000
      Top             =   2040
      Width           =   3120
   End
End
Attribute VB_Name = "Frmopen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
Unload Me
frmlog.Show
End Sub
