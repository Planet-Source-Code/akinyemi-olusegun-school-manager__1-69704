VERSION 5.00
Begin VB.Form frmpayroll 
   BorderStyle     =   0  'None
   Caption         =   "payroll"
   ClientHeight    =   7065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6825
   LinkTopic       =   "Form2"
   ScaleHeight     =   7065
   ScaleWidth      =   6825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstded 
      Appearance      =   0  'Flat
      Height          =   2565
      Left            =   3480
      TabIndex        =   22
      Top             =   2520
      Width           =   1935
   End
   Begin VB.ListBox lstamod 
      Appearance      =   0  'Flat
      Height          =   2565
      Left            =   5400
      TabIndex        =   21
      Top             =   2520
      Width           =   1215
   End
   Begin VB.ComboBox cbomon 
      Height          =   315
      ItemData        =   "frmpayroll.frx":0000
      Left            =   2640
      List            =   "frmpayroll.frx":0028
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   960
      Width           =   1815
   End
   Begin VB.ListBox lstamo 
      Appearance      =   0  'Flat
      Height          =   2565
      Left            =   2160
      TabIndex        =   12
      Top             =   2520
      Width           =   1215
   End
   Begin VB.ListBox lstall 
      Appearance      =   0  'Flat
      Height          =   2565
      Left            =   240
      TabIndex        =   9
      Top             =   2520
      Width           =   1935
   End
   Begin VB.TextBox txtdes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1920
      Width           =   2055
   End
   Begin VB.ComboBox cboid 
      Height          =   315
      Left            =   720
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox txtnam 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1440
      Width           =   2535
   End
   Begin Project1.ctrl_SkinableForm ctrl_SkinableForm 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1508
      Caption         =   "Staff Payroll"
      CaptionTop      =   0
   End
   Begin Project1.ctrl_SkinableButton btnquit 
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   6240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "Close"
   End
   Begin Project1.ctrl_SkinableButton btnprint 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   6240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "&Save"
   End
   Begin VB.Label Label18 
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
      Left            =   5640
      TabIndex        =   32
      Top             =   2280
      Width           =   645
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Deductions"
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
      Left            =   3840
      TabIndex        =   31
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label16 
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
      Left            =   2400
      TabIndex        =   30
      Top             =   2280
      Width           =   645
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Allowances"
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
      Left            =   600
      TabIndex        =   29
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   2880
      TabIndex        =   28
      Top             =   5400
      Width           =   150
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Deductions"
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
      Left            =   3240
      TabIndex        =   27
      Top             =   5520
      Width           =   975
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   1440
      TabIndex        =   26
      Top             =   5400
      Width           =   240
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Allowances"
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
      Left            =   1800
      TabIndex        =   25
      Top             =   5520
      Width           =   975
   End
   Begin VB.Label lll 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   6000
      TabIndex        =   24
      Top             =   5160
      Width           =   435
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
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
      Left            =   5400
      TabIndex        =   23
      Top             =   5160
      Width           =   450
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Print Slip"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   5280
      TabIndex        =   20
      Top             =   5880
      Width           =   780
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
      Left            =   1680
      TabIndex        =   19
      Top             =   960
      Width           =   540
   End
   Begin VB.Label lbltotal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000"
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
      Left            =   4920
      TabIndex        =   17
      Top             =   5520
      Width           =   435
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   4440
      TabIndex        =   16
      Top             =   5400
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Basic Salary"
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
      Left            =   240
      TabIndex        =   15
      Top             =   5520
      Width           =   1065
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
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
      Left            =   2160
      TabIndex        =   14
      Top             =   5160
      Width           =   450
   End
   Begin VB.Label lbltot 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   2760
      TabIndex        =   13
      Top             =   5160
      Width           =   435
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Basic Salary"
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
      Left            =   3480
      TabIndex        =   11
      Top             =   1920
      Width           =   1065
   End
   Begin VB.Label lblsalary 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   4800
      TabIndex        =   10
      Top             =   1920
      Width           =   435
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID:"
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
      Left            =   240
      TabIndex        =   7
      Top             =   1440
      Width           =   270
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Designation"
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
      Left            =   240
      TabIndex        =   5
      Top             =   1920
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
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
      Left            =   2760
      TabIndex        =   4
      Top             =   1440
      Width           =   495
   End
End
Attribute VB_Name = "frmpayroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnprint_Click()
StayOnTop Me, 0
Dim rs As Recordset
Call link
Set rs = db.OpenRecordset("select * from salary where ID='" & cboid.Text & "' and month='" & cbomon.Text & "'")
With rs
If .RecordCount < 1 Then
.AddNew
rs("month") = cbomon.Text
rs("salary") = lblTotal.Caption
rs("ID") = cboid.Text
.Update
Else
d = MsgBox("Salary for this month already exist." + vbCrLf + "Do you want to edit it ?", vbYesNo + vbQuestion, PROJ)
If d = vbYes Then
   EditO
   f = MsgBox("Staff Payment information saved", vbOKOnly, PROJ)
   
   Else
Exit Sub
End If
 End If


End With
StayOnTop Me, True
End Sub
Sub EditO()
Dim rs As Recordset
Call link
Set rs = db.OpenRecordset("select * from salary where ID='" & cboid.Text & "' and month='" & cbomon.Text & "'")
With rs
.Edit
rs("month") = cbomon.Text
rs("salary") = lblTotal.Caption
rs("ID") = cboid.Text
.Update
End With
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

Private Sub cboid_Click()
StayOnTop Me, 0
On Error Resume Next
Dim rs As Recordset
Dim rs1 As Recordset
Dim rs3 As Recordset
Dim rs4 As Recordset
lstall.Clear
lstamo.Clear
lstamod.Clear
lstded.Clear
Call link
Set rs = db.OpenRecordset("select * from staff where ID='" & cboid.Text & "'")
    With rs
            d = rs("Designation")
            txtnam.Text = rs("Name")
            txtdes.Text = rs("Designation")
    End With
    
    
 Set rs1 = db.OpenRecordset("select * from scale1 where Designation='" & d & "'")
    With rs1
        lblsalary.Caption = rs1("salary")
        g = rs1("salary")
    End With
    
 Set rs3 = db.OpenRecordset("select * from scale2 where Category='" & d & "'")
        With rs3
            While Not .EOF
                lstall.AddItem rs3("Type")
                lstamo.AddItem rs3("Amount")
                aj = aj + rs3("Amount")
                lbltot.Caption = aj
                            .MoveNext
            Wend
        
        
        End With
    Set rs4 = db.OpenRecordset("select * from deduction2 where ID='" & cboid.Text & "'")
    
        With rs4
            While Not .EOF
                lstded.AddItem rs4("Type")
                lstamod.AddItem rs4("Amount")
                a = Val(a) + Val(rs4("Amount"))
                lll.Caption = ""
                lll.Caption = a
                            .MoveNext
            Wend
            End With
           ' lbltotal.Caption = Val(lbltot.Caption + lblsalary.Caption - lll.Caption)
            lblTotal.Caption = ""
           lblTotal.Caption = Val((Val(g) + Val(aj)) - Val(a))
StayOnTop Me, True
End Sub

Private Sub Form_Load()
Me.Top = Val(Screen.Height - Me.Height) / 2
Me.Left = Val(Screen.Width - Me.Width) / 2
inito
StayOnTop Me, True


Dim rs As Recordset
Call link
Set rs = db.OpenRecordset("select * from staff")
    With rs
        While Not .EOF
                cboid.AddItem rs("id")
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
        
        .btnprint.SkinPath = App.Path & "\Skins\Deco"
        .btnprint.ForeColor = &H0&
        .btnprint.LoadSkin
        .btnprint.Refresh
        
        
        .btnquit.SkinPath = App.Path & "\Skins\Deco"
        .btnquit.ForeColor = &H0&
        .btnquit.LoadSkin
        .btnquit.Refresh
End With
End Sub
