Attribute VB_Name = "modTools"
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

Global Today As Variant
Global filename As String
Global CmdType As String
Global FindType As String
Global TableType As String
Global FieldType As String
Global ListType As Integer
Global DescType As Integer
Global FormType As String
Global EditMode As Boolean
Global blnAuto As Boolean
'Global db As Database
'Global rst As Recordset
'Dim dummy As Recordset

' PAD STRING
Function Pad_Str(str As String, val_to_pad As String, strlength As Integer, Right As Boolean) As String
    Dim s1 As String
    s1 = ""
    For i = 1 To strlength - Len(str) Step 1
        s1 = s1 & val_to_pad
    Next i
    If Right Then
        Pad_Str = str & s1
    Else
        Pad_Str = s1 & str
    End If
End Function

' SEARCH VALIDATION


' CONVERT STRING TO NUMERIC


'
' SQL QUERY DATA FOR FIND BOX USE ONLY



' COPY PICTURE FILE TO A FIELDNAME
Function CopyFileToField(filename As String, fd As DAO.Field)
    Dim ChunkSize As Long
    Dim FileNum As Integer
    Dim Buffer()  As Byte
    Dim BytesNeeded As Long
    Dim Buffers As Long
    Dim Remainder As Long
    Dim i As Long
    If Len(filename) = 0 Then
        Exit Function
    End If
    If Dir(filename) = "" Then
        Err.Raise vbObjectError, , "File not found: """ & filename & """"
    End If
    ChunkSize = 65536
    FileNum = FreeFile
    Open filename For Binary As #FileNum
    BytesNeeded = LOF(FileNum)
    Buffers = BytesNeeded \ ChunkSize
    Remainder = BytesNeeded Mod ChunkSize
    For i = 0 To Buffers - 1
        ReDim Buffer(ChunkSize)
        Get #FileNum, , Buffer
        fd.AppendChunk Buffer
    Next
    ReDim Buffer(Remainder)
    Get #FileNum, , Buffer
    fd.AppendChunk Buffer
    Close #FileNum
End Function
