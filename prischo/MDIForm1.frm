VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmmain 
   Caption         =   "Main Menu"
   ClientHeight    =   8190
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11880
   LinkTopic       =   "Form2"
   Picture         =   "MDIForm1.frx":0000
   ScaleHeight     =   8190
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar status 
      Align           =   2  'Align Bottom
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   7575
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1085
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7057
            MinWidth        =   7057
            Picture         =   "MDIForm1.frx":C4FB
            Text            =   "Logged User:"
            TextSave        =   "Logged User:"
            Object.ToolTipText     =   "logged on user"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
            Picture         =   "MDIForm1.frx":CDD5
            Text            =   "Event:"
            TextSave        =   "Event:"
            Object.ToolTipText     =   "report what is going on"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
            Picture         =   "MDIForm1.frx":EAAF
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer tmrUpdate 
      Interval        =   100
      Left            =   2400
      Top             =   1200
   End
   Begin MSComctlLib.ImageList IMG 
      Left            =   1800
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   65535
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   35
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":10789
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":11463
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":122B5
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":13107
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":139E1
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":142BB
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":14B95
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1555F
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":15E39
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":16153
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":16A2D
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":17307
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":17BE1
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":17EFB
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":187D5
            Key             =   "IMG15"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":190AF
            Key             =   "IMG16"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":19989
            Key             =   "IMG17"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1A263
            Key             =   "IMG18"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1AB3D
            Key             =   "IMG19"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1B417
            Key             =   "IMG20"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1BCF1
            Key             =   "IMG21"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1C5CB
            Key             =   "IMG22"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1CEA5
            Key             =   "IMG23"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1D77F
            Key             =   "IMG24"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1E059
            Key             =   "IMG25"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1E933
            Key             =   "IMG26"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1F20D
            Key             =   "IMG27"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1FAE7
            Key             =   "IMG28"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":203C1
            Key             =   "IMG29"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":20C9B
            Key             =   "IMG30"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":21551
            Key             =   "IMG31"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":21E2B
            Key             =   "IMG32"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2227D
            Key             =   "IMG33"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":226CF
            Key             =   "IMG34"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":24E81
            Key             =   "IMG35"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList IMGert 
      Left            =   2760
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   22
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":26103
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":26555
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":269A7
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":26DF9
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":26EA4
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":26FFE
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":28CD8
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":28D63
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":28E31
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":28EDA
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2910D
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2955F
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2AD21
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2C9FB
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2CD80
            Key             =   "IMG15"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2D052
            Key             =   "IMG16"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2D305
            Key             =   "IMG17"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2D3BA
            Key             =   "IMG18"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2D86D
            Key             =   "IMG19"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2DAAB
            Key             =   "IMG20"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2DBD3
            Key             =   "IMG21"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2F8AD
            Key             =   "IMG22"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1005
      ButtonWidth     =   1667
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      ImageList       =   "IMG"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "File"
            Description     =   "Issue Book"
            ImageIndex      =   10
            Style           =   5
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "View record"
            ImageIndex      =   21
            Style           =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Transaction"
            Description     =   "Return Book"
            ImageIndex      =   35
            Style           =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Student"
            Description     =   "Separator"
            ImageIndex      =   2
            Style           =   5
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Report"
            Description     =   "Book Records"
            ImageIndex      =   1
            Style           =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Finance"
            Description     =   "Member Records"
            ImageIndex      =   10
            Style           =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Search"
            Description     =   "Reports"
            ImageIndex      =   33
            Style           =   5
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList ImgList32 
      Left            =   3720
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":31587
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":32261
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":330B3
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":33D8D
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":34A67
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":35741
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":3641B
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":370F5
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Anounymous Mode"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   29.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   810
      Left            =   480
      TabIndex        =   0
      Top             =   3840
      Width           =   5790
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu dgdggdgd 
         Caption         =   "{SIDEBAR:TEXT|CAPTION:HOUSE KEEPING |Font:Arial|BOLD|Fsize:10|Fcolor:16777111|Bcolor:&HFFFFC0C0&|Gradient}"
      End
      Begin VB.Menu mnubackup 
         Caption         =   "{IMG:9}Backup Files"
      End
      Begin VB.Menu g 
         Caption         =   "-"
      End
      Begin VB.Menu mnulogout 
         Caption         =   "{IMG:20}Logout"
      End
      Begin VB.Menu n 
         Caption         =   "-"
      End
      Begin VB.Menu mnuchangeuser 
         Caption         =   "{IMG:19}Change User"
      End
      Begin VB.Menu jh 
         Caption         =   "-"
      End
      Begin VB.Menu mnucreateuser 
         Caption         =   "Create User"
      End
      Begin VB.Menu khjkhjk 
         Caption         =   "-"
      End
      Begin VB.Menu jhkkhjk 
         Caption         =   "-"
      End
      Begin VB.Menu jhhj 
         Caption         =   "-"
      End
      Begin VB.Menu fgfhgfh 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "{IMG:5}Exit"
      End
   End
   Begin VB.Menu mnuviewrecord 
      Caption         =   "View Record"
      Begin VB.Menu YUYU 
         Caption         =   "{SIDEBAR:TEXT|CAPTION:VIEW RECORDS|Font:Arial|BOLD|Fsize:10|Fcolor:16777215|Bcolor:&H00FFC0C0&|Gradient}"
      End
      Begin VB.Menu mnucandidate 
         Caption         =   "{IMG:2}Candidate"
      End
      Begin VB.Menu g1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuclasslevel 
         Caption         =   "{IMG:18}Basic Fee"
      End
      Begin VB.Menu sd 
         Caption         =   "-"
      End
      Begin VB.Menu mnufees 
         Caption         =   "{IMG:4}OtherFees"
      End
      Begin VB.Menu frrt 
         Caption         =   "-"
      End
      Begin VB.Menu mnudiscount 
         Caption         =   "{IMG:10}Discount"
      End
      Begin VB.Menu tgrt 
         Caption         =   "-"
      End
      Begin VB.Menu mnuacademic 
         Caption         =   "{IMG:22}Academic Year"
      End
      Begin VB.Menu mnudfs 
         Caption         =   "-"
      End
      Begin VB.Menu mnuteachers 
         Caption         =   "School"
         Begin VB.Menu fdsfsf 
            Caption         =   "{SIDEBAR:TEXT|CAPTION:OTHER TYPE |Font:Arial|BOLD|Fsize:10|Fcolor:16777255|Bcolor:&HFFFFC0C0&|Gradient}"
         End
         Begin VB.Menu mnustaffdfgdfg 
            Caption         =   "Staff"
         End
         Begin VB.Menu mnuassetfgdfg 
            Caption         =   "Asset"
         End
         Begin VB.Menu mnuclassrooms 
            Caption         =   "Class Rooms"
         End
      End
   End
   Begin VB.Menu mnutransaction 
      Caption         =   "Transaction"
      Begin VB.Menu mnustudentra 
         Caption         =   "{IMG:35}Student Payment"
      End
      Begin VB.Menu mnustaffimoo 
         Caption         =   "Staff"
         Begin VB.Menu mnudeduction 
            Caption         =   "Deduction"
         End
         Begin VB.Menu mnuscale 
            Caption         =   "Staff Salary Scale"
         End
         Begin VB.Menu mnustafftransaction 
            Caption         =   "Staff payment"
         End
      End
   End
   Begin VB.Menu mnustudentdfgdfgd 
      Caption         =   "Student"
      Begin VB.Menu mnutimetable 
         Caption         =   "Time Table"
      End
      Begin VB.Menu mnuacademics 
         Caption         =   "Academics"
         Begin VB.Menu RTHJFGH 
            Caption         =   "{SIDEBAR:TEXT|CAPTION:PERFORMANCE |Font:Arial|BOLD|Fsize:14|Fcolor:26777255|Bcolor:&H0080C0FF&|Gradient}"
         End
         Begin VB.Menu mnucoursesetup 
            Caption         =   "{IMG:8}Course Setup"
         End
         Begin VB.Menu mnuperformanceasset 
            Caption         =   "Examination and Assesment"
         End
         Begin VB.Menu mnugenerateresult 
            Caption         =   "Generate Result"
         End
         Begin VB.Menu mnuprintresuklt 
            Caption         =   "Print Result"
         End
      End
   End
   Begin VB.Menu mnureportser 
      Caption         =   "Report"
      Begin VB.Menu mnurepstudent 
         Caption         =   "Student"
         Begin VB.Menu mnuclass 
            Caption         =   "Class"
         End
         Begin VB.Menu mnupayed 
            Caption         =   "Payed"
         End
         Begin VB.Menu mnudebtors 
            Caption         =   "Debtors"
         End
         Begin VB.Menu mnusexfdgdfg 
            Caption         =   "Sex"
         End
      End
      Begin VB.Menu mnustaffreport 
         Caption         =   "Staff"
         Begin VB.Menu mnustafreport 
            Caption         =   "Staff by Sex"
         End
         Begin VB.Menu mnuhatis 
            Caption         =   "Staff by Hatis"
         End
         Begin VB.Menu mnudesignation 
            Caption         =   "Staff by Designation"
         End
      End
   End
   Begin VB.Menu mnufinance 
      Caption         =   "Finance"
      Begin VB.Menu mnuincome 
         Caption         =   "Income"
      End
      Begin VB.Menu mnuexpenditure 
         Caption         =   "Expenditures"
      End
   End
   Begin VB.Menu mnusearch 
      Caption         =   "Search"
      Begin VB.Menu mnustaff 
         Caption         =   "Staff"
      End
      Begin VB.Menu mnustudent 
         Caption         =   "Student"
      End
      Begin VB.Menu mnuasset 
         Caption         =   "Asset"
      End
      Begin VB.Menu mnucourse 
         Caption         =   "Course"
      End
   End
   Begin VB.Menu mnucurrentuser 
      Caption         =   "Current User"
      Visible         =   0   'False
      Begin VB.Menu mnuuser 
         Caption         =   "user"
      End
      Begin VB.Menu mnulogout2 
         Caption         =   "Logout"
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'##############################################
'#          Coded by Akinyemi Olusegun        #
'#                                            #
'#                                            #
'#                                            #
'#    description :  About the Author         #
'#         e-mail :  segzee20002001@yahoo.com #
'#    url :  http://maxisoft.bravehost.com    #
'#                                            #
'##############################################



Private Sub ctrl_ListObject_Click(Index As Integer)
Select Case Index
Case 1
frmfee.Show
Case 2
frmstaff.Show
Case 3
frmasset.Show
Case 4
frmpay.Show
Case 5
Case 6
'change user
Case 7
'logout
Case 8
frmback.Show
Case 9
Form1.Show
Case 10
frmyear.Show

End Select

End Sub

'Private Sub ctrl_PullDownMenu_Click(Index As Integer)
'Select Case Index

'Case 1
'            PopupMenu mnufile, , Me.ctrl_PullDownMenu.Left + Me.ctrl_PullDownMenu.pSelectionLeft, Me.ctrl_PullDownMenu.Top + Me.ctrl_PullDownMenu.pSelectionBottom

'Case 2
 '           PopupMenu mnuviewrecord, , Me.ctrl_PullDownMenu.Left + Me.ctrl_PullDownMenu.pSelectionLeft, Me.ctrl_PullDownMenu.Top + Me.ctrl_PullDownMenu.pSelectionBottom

'Case 3
 '           PopupMenu mnutransaction, , Me.ctrl_PullDownMenu.Left + Me.ctrl_PullDownMenu.pSelectionLeft, Me.ctrl_PullDownMenu.Top + Me.ctrl_PullDownMenu.pSelectionBottom

'Case 4
 '           PopupMenu mnustudent, , Me.ctrl_PullDownMenu.Left + Me.ctrl_PullDownMenu.pSelectionLeft, Me.ctrl_PullDownMenu.Top + Me.ctrl_PullDownMenu.pSelectionBottom
'Case 5
  '         PopupMenu mnufinance, Me.ctrl_PullDownMenu.Left + Me.ctrl_PullDownMenu.pSelectionLeft, Me.ctrl_PullDownMenu.Top + Me.ctrl_PullDownMenu.pSelectionBottom
'Case 6
 '          PopupMenu mnusearch, Me.ctrl_PullDownMenu.Left + Me.ctrl_PullDownMenu.pSelectionLeft, Me.ctrl_PullDownMenu.Top + Me.ctrl_PullDownMenu.pSelectionBottom
'Case 7
 '          PopupMenu mnureportser, Me.ctrl_PullDownMenu.Left + Me.ctrl_PullDownMenu.pSelectionLeft, Me.ctrl_PullDownMenu.Top + Me.ctrl_PullDownMenu.pSelectionBottom

'End Select
'End Sub

Sub inito()
With Me




' .ctrl_PullDownMenu.BackColor = &H2E2E32
 '       .ctrl_PullDownMenu.ForeColor = &HFFFFFF
  '      .ctrl_PullDownMenu.Refresh
        
        
             
    'Call ctrl_PullDownMenu.AddItem("File")
    'Call ctrl_PullDownMenu.AddItem("Record")
    'Call ctrl_PullDownMenu.AddItem("Transaction")
    'Call ctrl_PullDownMenu.AddItem("Student")
   'Call ctrl_PullDownMenu.AddItem("Finance")
  ' Call ctrl_PullDownMenu.AddItem("Search")
   'Call ctrl_PullDownMenu.AddItem("Report")
End With
End Sub

Private Sub Form_Load()
inito
'StayOnTop Me, True
Label14.Visible = 0
frmlog.Hide
Unload frmlog

SetMenus hwnd, IMG



End Sub

Private Sub mnuacademic_Click()
frmyear.Show
End Sub

Private Sub mnuasset_Click()
frmasset.Show
End Sub

Private Sub mnuassetfgdfg_Click()
frmasset.Show
End Sub

Private Sub mnubackup_Click()
frmback.Show
End Sub

Private Sub mnucandidate_Click()
Form1.Show
End Sub

Private Sub mnuchangeuser_Click()
frmlog.ctrl_SkinableForm.Caption = "Switch user"
frmlog.Show
End Sub

Private Sub mnuclasslevel_Click()
frmclevel.Show
End Sub

Private Sub mnucoursesetup_Click()
frmcourses.Show
End Sub

Private Sub mnucreateuser_Click()
frmuser.Show
End Sub

Private Sub mnudeduction_Click()
frmdeduction.Show
End Sub

Private Sub mnudiscount_Click()
frmdiscount.Show
End Sub

Private Sub mnuexit_Click()
f = MsgBox("Do you really want to quit ? [yes/No]", vbYesNo + vbQuestion, "Closing")
    If f = vbYes Then
        Unload Me
    Else
        Exit Sub
    End If
End Sub

Private Sub mnufees_Click()
frmfee.Show
End Sub

Private Sub mnuperfandasse_Click()
frmcourses.Show
End Sub

Private Sub mnugenerateresult_Click()
frmgenerate.Show
End Sub

Private Sub mnulogout_Click()
f = MsgBox("Do you really want to log out ?", vbYesNo + vbQuestion, "Confirmation")
If f = vbYes Then
frmlogout.Show
Else
Exit Sub
End If
End Sub

Private Sub mnulogout2_Click()
Call mnulogout_Click

End Sub

Private Sub mnuperformanceasset_Click()
frmcard.Show
End Sub

Private Sub mnuprintresuklt_Click()
frmresult.Show
End Sub

Private Sub mnuscale_Click()
frmscale.Show
End Sub

Private Sub mnustaffdfgdfg_Click()
frmstaff.Show
End Sub

Private Sub mnustafftransaction_Click()
frmpayroll.Show
End Sub

Private Sub mnustudentra_Click()

frmpay.Show
End Sub


Private Function AssignColour(nPercent As Integer) As Long
    Select Case nPercent
        Case Is < 10
            AssignColour = vbRed
        Case Is < 25
            AssignColour = vbMagenta
        Case Else
            AssignColour = vbBlue
    End Select
End Function

Private Sub mnuuser_Click()
DoEvents


End Sub

Private Sub status_PanelClick(ByVal Panel As MSComctlLib.Panel)
Select Case Panel.Index
Case 1
PopupMenu mnucurrentuser, , status.Panels(1).Left - 70, status.Top - 5


End Select
End Sub

Private Sub tmrUpdate_Timer()
status.Panels(3).Text = Time

End Sub

Private Sub Toolbar1_ButtonDropDown(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
        Case 1
        PopupMenu mnufile
        Case 2
        PopupMenu mnuviewrecord
        Case 3
        PopupMenu mnutransaction
        Case 4
        PopupMenu mnustudentdfgdfgd
        Case 5
        PopupMenu mnureportser
        Case 6
        PopupMenu mnufinance
        Case 7
        PopupMenu mnusearch
End Select
End Sub

