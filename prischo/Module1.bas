Attribute VB_Name = "Module1"
'Database connectivity module
'Using MS DAO 3.51 Library
'
'
'
'


Public db As Database
Public rs As Recordset

Public Const RECERROR As String = "Record Not found"

Public Const PROJ As String = "Student Information system"
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Public Sub StayOnTop(frmForm As Form, fOnTop As Boolean)
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2
    
    Dim lState As Long
    Dim iLeft As Integer, iTop As Integer, iWidth As Integer, iHeight As Integer

    With frmForm
        iLeft = .Left / Screen.TwipsPerPixelX
        iTop = .Top / Screen.TwipsPerPixelY
        iWidth = .Width / Screen.TwipsPerPixelX
        iHeight = .Height / Screen.TwipsPerPixelY
    End With
    
    If fOnTop Then
        lState = HWND_TOPMOST
    Else
        lState = HWND_NOTOPMOST
    End If
    SetWindowPos frmForm.hwnd, lState, iLeft, iTop, iWidth, iHeight, 80
End Sub
 Sub link()
Set db = OpenDatabase(App.Path + "\data.mdb")
End Sub





