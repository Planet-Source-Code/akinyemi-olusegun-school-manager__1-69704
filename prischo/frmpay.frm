VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmpay 
   BorderStyle     =   0  'None
   ClientHeight    =   7590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9750
   LinkTopic       =   "Form2"
   ScaleHeight     =   7590
   ScaleWidth      =   9750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Project1.ctrl_SkinableButton btnpay 
      Height          =   375
      Left            =   3720
      TabIndex        =   25
      Top             =   6600
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Caption         =   "Save/Receipt"
   End
   Begin VB.TextBox txtacc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   7080
      Width           =   1455
   End
   Begin VB.TextBox txtamo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   6720
      TabIndex        =   17
      Top             =   4560
      Width           =   1695
   End
   Begin Project1.ctrl_SkinableButton btn 
      Height          =   375
      Left            =   7800
      TabIndex        =   10
      Top             =   3480
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "Fees"
   End
   Begin VB.TextBox txtmod 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox txtnam 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1680
      Width           =   2655
   End
   Begin VB.ComboBox cboid 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1680
      Width           =   1095
   End
   Begin VB.ComboBox cbocla 
      Height          =   315
      ItemData        =   "frmpay.frx":0000
      Left            =   2160
      List            =   "frmpay.frx":0022
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1680
      Width           =   1095
   End
   Begin Project1.ctrl_SkinableForm ctrl_SkinableForm 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      Caption         =   "Student Account"
      CaptionTop      =   0
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3195
      Left            =   240
      TabIndex        =   9
      Top             =   3360
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   5636
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "Imagelist1"
      SmallIcons      =   "Imagelist1"
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Account NO."
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ID"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Year"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Total Amount"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Amount Paid"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Balance"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Class"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
   End
   Begin Project1.ctrl_SkinableButton btndis 
      Height          =   375
      Left            =   7800
      TabIndex        =   18
      Top             =   3960
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "Discount"
   End
   Begin Project1.ctrl_SkinableButton btncalc 
      Height          =   375
      Left            =   6480
      TabIndex        =   21
      Top             =   5640
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "Calculate"
   End
   Begin Project1.ctrl_SkinableButton btnquit 
      Height          =   375
      Left            =   7800
      TabIndex        =   26
      Top             =   7080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "Close"
   End
   Begin Project1.ctrl_SkinableButton btnclear 
      Height          =   375
      Left            =   6000
      TabIndex        =   27
      Top             =   7080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "Clear"
   End
   Begin MSComctlLib.ImageList Imagelist1 
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
            Picture         =   "frmpay.frx":0094
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Project1.ctrl_SkinableButton ctrl_SkinableButton1 
      Height          =   375
      Left            =   3720
      TabIndex        =   32
      Top             =   7080
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Caption         =   "Print Reciept"
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
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
      Left            =   5520
      TabIndex        =   31
      Top             =   6240
      Width           =   615
   End
   Begin VB.Label lblstatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   300
      Left            =   6480
      TabIndex        =   30
      Top             =   6240
      Width           =   90
   End
   Begin VB.Label lblfinal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   6120
      TabIndex        =   29
      Top             =   6600
      Width           =   75
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Final:"
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
      Left            =   5520
      TabIndex        =   28
      Top             =   6600
      Width           =   480
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Account Number:"
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
      Left            =   480
      TabIndex        =   24
      Top             =   7080
      Width           =   1485
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Candidate Payment Screen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   960
      TabIndex        =   22
      Top             =   1080
      Width           =   3285
   End
   Begin VB.Label lbldebt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   6480
      TabIndex        =   20
      Top             =   5160
      Width           =   75
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Balance:"
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
      Left            =   5520
      TabIndex        =   19
      Top             =   5160
      Width           =   765
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Discount:"
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
      Left            =   5520
      TabIndex        =   16
      Top             =   3960
      Width           =   825
   End
   Begin VB.Label lblpay 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   6480
      TabIndex        =   15
      Top             =   3960
      Width           =   75
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   9240
      TabIndex        =   14
      Top             =   2760
      Width           =   75
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount Paid:"
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
      Left            =   5520
      TabIndex        =   13
      Top             =   4560
      Width           =   1140
   End
   Begin VB.Label lbltot 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   6600
      TabIndex        =   12
      Top             =   3480
      Width           =   75
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Fee:"
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
      Left            =   5520
      TabIndex        =   11
      Top             =   3480
      Width           =   885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   4
      X1              =   0
      X2              =   9120
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mode:"
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
      Left            =   1560
      TabIndex        =   8
      Top             =   2280
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
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
      TabIndex        =   6
      Top             =   1680
      Width           =   555
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
      Left            =   3600
      TabIndex        =   3
      Top             =   1680
      Width           =   270
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Class"
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
      Left            =   1560
      TabIndex        =   2
      Top             =   1680
      Width           =   465
   End
End
Attribute VB_Name = "frmpay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_Click()
frmfee.Show
End Sub

Private Sub btn_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
btn.Refresh
End Sub

Private Sub btncalc_Click()
    Dim a, B, c, d, e, f As Integer
    a = Val(Me.lbltot.Caption)
    B = Val(Me.lblpay.Caption)
    c = Val(a - B)
    d = Val(txtamo)
    e = Val(c - d)
    Me.lbldebt.Caption = e
    'Format$(IIf(IsNull(e), "", e), "#,##0.00")
    lblfinal.Caption = c
    'Format$(IIf(IsNull(c), "", c), "#,##0.00")

End Sub

Private Sub btncalc_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
btncalc.Refresh
End Sub

Private Sub btnclear_Click()
    Unload Me
    Me.Show
End Sub

Private Sub btnclear_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
btnclear.Refresh
End Sub

Private Sub btndis_Click()
    frmdis.Show
End Sub

Private Sub btndis_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
btndis.Refresh
End Sub

Private Sub btnpay_Click()
On Error Resume Next
Call StayOnTop(Me, True)
Dim rs As Recordset
Dim rs2 As Recordset
Call link
Set rs = db.OpenRecordset("select * from account where ACCNumber='" & txtacc.Text & "'")
With rs
If .RecordCount > 0 Then
    If Val(rs("debt")) > 0 Then
    qs = "Owing"
    lblstatus.Caption = qs
        d = MsgBox("This candidate has paid but still owing", vbOKCancel, PROJ)
        Else
        qs = "Paid"
            lblstatus.Caption = qs

    End If

    Call Edit
Else
    .AddNew
    rs("ACCNumber") = txtacc.Text
    rs("ID") = cboid.Text
    rs("Year") = str(Val(Year(Now) - 1)) + "/" + str(Year(Now))
    rs("Total_Amount") = lbltot.Caption
    rs("Amount_payed") = txtamo.Text
    rs("Debt") = lbldebt.Caption
    rs("class") = cbocla.Text
    If Val(txtamo) = Val(lblfinal.Caption) Then
    rs("status") = "Paid"
    Frmreciept.lblaccnum.Caption = txtacc.Text
    Frmreciept.lblamo.Caption = lblfinal.Caption
    Frmreciept.lblclas.Caption = cbocla.Text
    Frmreciept.lblid.Caption = cboid.Text
    Frmreciept.lblnam.Caption = txtnam.Text
    Frmreciept.lblpay.Caption = txtamo.Text
    Frmreciept.lblnum.Caption = "R" + str(txtacc.Text)
Me.WindowState = 1

    Frmreciept.Show
    Else
    rs("status") = "Owing"
    End If
    .Update
    
        d = MsgBox("Information saved.Reciept will be printed shortly" + vbCrLf + " Make sure there is paper in the printer tray.", vbInformation, PROJ)
      
        Me.WindowState = 1

'Unload Me
'Me.Show
doGet
End If
'End If
End With


End Sub
Sub doGet()
Dim rs As Recordset
Call link
Set rs = db.OpenRecordset("select * from account where class='" & cbocla.Text & "'")
With rs
    While Not .EOF
        With ListView1
        
            .ListItems.add , , rs("ACCNumber"), 1, 1
            .ListItems(.ListItems.Count).SubItems(1) = rs("ID")
            .ListItems(.ListItems.Count).SubItems(2) = rs("year")
            .ListItems(.ListItems.Count).SubItems(3) = rs("Total_Amount")
            'Format$(IIf(IsNull(rs("Total_Amount")), "", rs("Total_Amount")), "#,##0.00")
            .ListItems(.ListItems.Count).SubItems(4) = rs("Amount_payed")
            'Format$(IIf(IsNull(rs("Amount_payed")), "", rs("Amount_payed")), "#,##0.00")
            .ListItems(.ListItems.Count).SubItems(5) = rs("Debt")
            'Format$(IIf(IsNull(rs("Debt")), "", rs("Debt")), "#,##0.00")
            .ListItems(.ListItems.Count).SubItems(6) = rs("class")
            .ListItems(.ListItems.Count).SubItems(7) = rs("status")
        End With
            .MoveNext
    Wend

End With
End Sub
Sub Edit()
Dim rs As Recordset
Dim rs2 As Recordset
Call link
Set rs = db.OpenRecordset("select * from account where ACCNumber='" & txtacc.Text & "'")
With rs
If .RecordCount > 0 Then

If rs("debt") > 0 Then
    qs = "Owing"
        lblstatus.Caption = qs

        d = MsgBox("This candidate has paid but still owing", vbOKCancel, PROJ)
        Else
        
        qs = "Paid"
            lblstatus.Caption = qs

    End If
.Edit
rs("ACCNumber") = txtacc.Text
rs("ID") = cboid.Text
rs("Year") = str(Val(Year(Now) - 1)) + "/" + str(Year(Now))
rs("Total_Amount") = lbltot.Caption
rs("Amount_payed") = txtamo.Text
rs("Debt") = lbldebt.Caption
rs("class") = cbocla.Text
rs("status") = qs
.Update
If qs = "Paid" Then
Frmreciept.lblaccnum.Caption = txtacc.Text
    Frmreciept.lblamo.Caption = lblfinal.Caption
    Frmreciept.lblclas.Caption = cbocla.Text
    Frmreciept.lblid.Caption = cboid.Text
    Frmreciept.lblnam.Caption = txtnam.Text
    Frmreciept.lblpay.Caption = txtamo.Text
    Frmreciept.lblnum.Caption = "R" + str(txtacc.Text)
    Frmreciept.Show
End If
'Unload Me
'Me.Show
End If
End With
End Sub

Private Sub btnpay_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
btnpay.Refresh
End Sub

Private Sub btnquit_Click()
Unload Me

End Sub

Private Sub btnquit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
btnquit.Refresh
End Sub

Private Sub cbocla_Click()
cboid.Clear

Dim rs As Recordset
Call link
Set rs = db.OpenRecordset("select * from studinfo where class='" & cbocla.Text & "'")
With rs
While Not .EOF

cboid.AddItem rs("ID")

.MoveNext
Wend


End With
doGet
End Sub

Private Sub cboid_Click()
StayOnTop frmpay, False
y = Year(Now)
txtacc = cboid.Text
Dim rs, rs1, rs2 As Recordset
Call link
Set rs2 = db.OpenRecordset("select * from year")
With rs2
y = rs2("year")
End With

Set rs = db.OpenRecordset("select * from studinfo where ID='" & cboid.Text & "'")
With rs
txtnam.Text = rs("name")
txtmod = rs("mode")
End With

Set rs1 = db.OpenRecordset("select * from account where id='" & cboid.Text & "' and year='" & y & "'")
With rs1
If .RecordCount < 1 Then
DoEvents
DoEvents
p = MsgBox("This candidate has not yet paid")
lblstatus.Caption = "Not Paid"
DoEvents
DoEvents
DoEvents
Call Get_pay
DoEvents

Else
Call Get_pay

End If

End With
StayOnTop frmpay, True

End Sub

Sub Get_pay()
'Dim d As Integer
Dim rs, rs1 As Recordset
Call link
Set rs = db.OpenRecordset("select * from fee")

With rs
    While Not .EOF
        d = d + Val(rs("amount"))
.MoveNext
lbltot.Caption = d
Wend

End With

Set rs1 = db.OpenRecordset("select * from tution where type='" & txtmod.Text & "'")
With rs1

d = d + Val(rs1("amount"))
lbltot.Caption = d

'Format$(IIf(IsNull(d), "", d), "#,##0.00")

End With

End Sub
Sub inito()
With Me
.ctrl_SkinableForm.SkinPath = App.Path & "\Skins\Deco"
        .ctrl_SkinableForm.BackColor = &HCECECE
        .ctrl_SkinableForm.CaptionTop = 300
        .ctrl_SkinableForm.CaptionColor = &H0&
        Call Me.ctrl_SkinableForm.LoadSkin(Me)
        
         .btncalc.SkinPath = App.Path & "\Skins\Deco"
       .btncalc.ForeColor = &H0&
        .btncalc.LoadSkin
        .btncalc.Refresh
                
        .btnpay.SkinPath = App.Path & "\Skins\Deco"
       .btnpay.ForeColor = &H0&
        .btnpay.LoadSkin
       .btnpay.Refresh
                
             
             .ctrl_SkinableButton1.SkinPath = App.Path & "\Skins\Deco"
       .ctrl_SkinableButton1.ForeColor = &H0&
        .ctrl_SkinableButton1.LoadSkin
       .ctrl_SkinableButton1.Refresh
             
                .btnclear.SkinPath = App.Path & "\Skins\Deco"
      .btnclear.ForeColor = &H0&
       .btnclear.LoadSkin
       .btnclear.Refresh
                
               .btnquit.SkinPath = App.Path & "\Skins\Deco"
      .btnquit.ForeColor = &H0&
      .btnquit.LoadSkin
       .btnquit.Refresh
                
        .btndis.SkinPath = App.Path & "\Skins\Deco"
       .btndis.ForeColor = &H0&
        .btndis.LoadSkin
        .btndis.Refresh
        
       
        .btn.SkinPath = App.Path & "\Skins\Deco"
       .btn.ForeColor = &H0&
        .btn.LoadSkin
        .btn.Refresh
        
End With
End Sub

Private Sub ctrl_SkinableButton1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
ctrl_SkinableButton1.Refresh
End Sub

Private Sub Form_Load()
StayOnTop Me, True
inito


Me.lblpay.Caption = 0
Me.Top = Val(Screen.Height - Me.Height) / 2
Me.Left = Val(Screen.Width - Me.Width) / 2
End Sub

Private Sub txtamo_Change()
'f = txtamo
'txtamo = Format$(IIf(IsNull(f), "", f), "#,##0.00")
End Sub

Private Sub txtamo_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
                Case Asc("0") To Asc("9")
                Case Asc(".")
                    KeyAscii = IIf(Index = 5 Or Index = 1, 0, KeyAscii)
                Case KeyAscii = vbKeyBack
                Case Else
                    KeyAscii = 0
        End Select
End Sub
