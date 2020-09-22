VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9105
   LinkTopic       =   "Form1"
   ScaleHeight     =   6315
   ScaleWidth      =   9105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form1.frx":0000
      Height          =   3375
      Left            =   480
      TabIndex        =   1
      Top             =   1320
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   5953
      _Version        =   393216
      BackColor       =   4210752
      BorderStyle     =   0
      ForeColor       =   65280
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin Project1.ctrl_SkinableButton btnnew 
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   5400
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      Caption         =   "New Record"
   End
   Begin Project1.ctrl_SkinableForm ctrl_SkinableForm 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1508
      Caption         =   "Candidate Record"
      CaptionTop      =   0
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1200
      Top             =   4800
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\prischo\DATA.MDB;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\prischo\DATA.MDB;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "studinfo"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin Project1.ctrl_SkinableButton btnqui 
      Height          =   375
      Left            =   6120
      TabIndex        =   3
      Top             =   5400
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      Caption         =   "Quit"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub Command1_Click()
'Dim rs As Recordset
'Call link
'Set rs = db.OpenRecordset("select * from studinf where ID='" & Text2 & "'")
'With rs
'If .RecordCount > 0 Then
 '   ' add more codes later
   ' Else
  '      MsgBox "No record found", vbExclamation
    
Private Sub btnnew_Click()
frmaddstud.Show
Unload Me
End Sub

Private Sub btnnew_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
btnnew.Refresh
End Sub

   ' End If
'End With


Private Sub Command2_Click()


End Sub

Private Sub btnqui_Click()
Unload Me

End Sub

Private Sub btnqui_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
btnqui.Refresh
End Sub

Private Sub Form_Load()
StayOnTop Me, True
' Change values in the ReportsTo field to 5 for all
    ' employee records that currently have ReportsTo
    ' values of 2.
   ' Call link
    'db.Execute "UPDATE studinfo " _
        & "SET Enrolled = 'No' " _
        & "WHERE Enrolled ='Yes';"
        
    
'On Error GoTo ferror
Me.Top = Val(Screen.Height - Me.Height) / 2
Me.Left = Val(Screen.Width - Me.Width) / 2

inito
Dim rs As Recordset
Call link
Set rs = db.OpenRecordset("select * from studinfo Order By ID Desc")

   
    
ferror:
'd = MsgBox(Err.Description, vbCritical, PROJ)
    
End Sub
Sub inito()
With Form1
.ctrl_SkinableForm.SkinPath = App.Path & "\Skins\Deco"
        .ctrl_SkinableForm.BackColor = &HCECECE
        .ctrl_SkinableForm.CaptionTop = 300
        .ctrl_SkinableForm.CaptionColor = &H0&
        Call Me.ctrl_SkinableForm.LoadSkin(Me)
        
        .btnnew.SkinPath = App.Path & "\Skins\Deco"
        .btnnew.ForeColor = &H0&
        .btnnew.LoadSkin
        .btnnew.Refresh
        
        .btnqui.SkinPath = App.Path & "\Skins\Deco"
        .btnqui.ForeColor = &H0&
        .btnqui.LoadSkin
        .btnqui.Refresh
        
End With
End Sub
