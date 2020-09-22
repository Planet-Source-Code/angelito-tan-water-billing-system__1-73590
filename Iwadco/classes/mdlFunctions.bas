Attribute VB_Name = "mdlFunctions"

Option Explicit
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public CN               As New ADODB.Connection
Public rs               As New ADODB.Recordset
Public lst              As ListItem
Public empID            As String
Public empType          As Long
Public lstview          As New DataViewer
Public FrmName          As String
Public sql              As String
Public CoorID           As Long
Public AreaID           As Long
Public txtVal           As String
Public txt              As TextStream
Public fso              As New FileSystemObject
Public fileforopen      As File
Public sqlserverdata()  As String
Public enc_md5          As MD5
Public smonth           As Long
Public syear            As Long
Public conName          As String
Dim maxRec
Public z
Public FK_ID            As Integer
Public tmpIDreadings    As Long
Public tmpIDconsumer    As String
Public CONSUMERID       As Variant
Public formBoolean      As Boolean
Public SelectedForm     As String
Public CN1 As New ADODB.Connection
Public aa1 As String 'para sa month -frmselcoor
Public aa2 As String 'para sa year-frmselcoor
Public aa3 As Integer  'para sa area-frmselcoor
Public asdasd As String

Function getMachineName() As String
Dim sBuffer As String
Dim lAns As Long
sBuffer = Space$(255)
lAns = GetComputerName(sBuffer, 255)
If lAns <> 0 Then
    getMachineName = Left$(sBuffer, InStr(sBuffer, Chr(0)) - 1)
End If
End Function

Public Function DBConnect(sqlServer As String, sqlUser As String, sqlPass As String, sqlDB As String) As Boolean
On Error GoTo errtrap
If CN.State = adStateOpen Then CN.Close
CN.ConnectionString = "DRIVER=SQL Server;SERVER=" & sqlServer & ";UID=" & sqlUser & ";PWD=" & sqlPass & ";APP=Visual Basic;DATABASE=" & sqlDB
CN.Open
If DataEnvironment1.Connection1.State = adStateOpen Then DataEnvironment1.Connection1.Close
DataEnvironment1.Connection1.ConnectionString = "DRIVER=SQL Server;SERVER=" & sqlServer & ";UID=" & sqlUser & ";PWD=" & sqlPass & ";APP=Visual Basic;DATABASE=" & sqlDB
DBConnect = True
errtrap:
Select Case err.Number
Case 0
Case Else
    Debug.Print err.Number
    DBConnect = False
    Exit Function
End Select
End Function

Public Function rsCheck()
If rs.State = 1 Then rs.Close
End Function

Public Function logInSucceeded(usrName As String, pssWord As String) As Boolean
rsCheck
rs.Open "SELECT * FROM iwadco_user WHERE username = '" & usrName & "' AND password='" & LCase(enc_md5.DigestStrToHexStr(pssWord)) & "' AND status='E'", CN, adOpenStatic, adLockOptimistic
If rs.RecordCount <> 0 Then
    logInSucceeded = True
    empID = rs(0)
    empType = rs(3)
    Exit Function
End If
logInSucceeded = False
End Function

Public Function LoadCombo(query As String, src As ComboBox)
rsCheck
rs.Open query, CN, adOpenStatic, adLockOptimistic
src.Clear
Do While rs.EOF <> True
    src.AddItem rs(0)
    rs.MoveNext
Loop
End Function

Public Function selectCombo(str As Long, src As ComboBox)
On Error Resume Next
Dim X As Long
For X = 1 To src.ListCount
    SendKeys "{DOWN}"
    'Debug.Print src.Text & "=" & str
    If str = X Then
        Exit For
    End If
Next
End Function

Public Function user_priv(priv_type As String) As Boolean
Dim r_priv() As String
Dim X As Long
rsCheck
rs.Open "SELECT iwadco_actype.priv FROM iwadco_actype INNER JOIN iwadco_user ON iwadco_actype.id = iwadco_user.type WHERE (iwadco_user.id = " & empType & ")"
r_priv = Split(rs(0), ",")
For X = 0 To UBound(r_priv)
    If r_priv(X) = priv_type Then
        user_priv = True
        Exit Function
    Else
        user_priv = False
    End If
Next
End Function

Public Sub FormShow(frm As Form, edit As Boolean)
Load frm
If edit = True Then
    frm.Caption = frm.Caption & " - Edit"
Else
    frm.Caption = frm.Caption & " - Add"
End If
frm.Show 1
End Sub

Public Function loadtxt(sFile As String) As Boolean
    Dim i As Integer, zxc As Integer
    Dim wordlist As String
    
    Open sFile For Input As #1
    While Not EOF(1)
    Input #1, wordlist
    txtVal = wordlist
    Wend
    Close #1
    sqlserverdata = Split(DeCode(txtVal), ":")
   
    If sqlserverdata(0) <> "" And sqlserverdata(3) <> "" Then
        loadtxt = True
    Else
        loadtxt = False
    End If
    Exit Function
errtrap:
Select Case err.Number
Case Else
    MsgBox err.Number
    MsgBox "Error has occured in loading database!" & vbCrLf & "Please contact the database administrator!", vbCritical, "Critical Error!"
    End
End Select
End Function
Public Function Unique_ID(CorID As Variant, AreaID As Variant, ByVal CoorID As Double, ByVal Area_ID) As String
Dim cntRec As Long, ConsID As Variant
'Dim GetID As Variant
    rsCheck
    rs.Open "SELECT * FROM iwadco_cons WHERE coor_id ='" & CoorID & "'AND area_id ='" & Area_ID & "'"
    maxRec = rs.RecordCount + 1
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        Do
            If maxRec < 10 Then
               ConsID = "0000"
               ElseIf maxRec >= 10 And maxRec <= 99 Then
               ConsID = "00000"
               ElseIf maxRec >= 99 And maxRec <= 999 Then
               ConsID = "000000"
               ElseIf maxRec >= 999 And maxRec <= 9999 Then
               ConsID = "0000000"
               ElseIf maxRec >= 9999 And maxRec <= 99999 Then
               ConsID = "00000000"
               ElseIf maxRec >= 99999 And maxRec <= 999999 Then
               ConsID = "000000000"
               ElseIf maxRec >= 999999 And maxRec <= 9999999 Then
               ConsID = "0000000000"
               ElseIf maxRec >= 9999999 And maxRec <= 99999999 Then
               ConsID = "00000000000"
               ElseIf maxRec >= 99999999 And maxRec <= 999999999 Then
               ConsID = "000000000000"
               ElseIf maxRec >= 9999999999# And maxRec <= 9999999999# Then
               ConsID = "0000000000000"
               'store up to ten billion
            End If
            
            Mid(ConsID, Len(maxRec)) = StrReverse(Mid(maxRec, 1))
            Unique_ID = CorID & "-" & AreaID & "-" & StrReverse(Mid(ConsID, Len(maxRec)))
            rsCheck
            rs.Open "SELECT * FROM iwadco_cons WHERE id ='" & Unique_ID & "'"
            If rs.RecordCount > 0 Then
                   maxRec = maxRec + 1
                   rsCheck
                   rs.Open "SELECT * FROM iwadco_cons WHERE coor_id ='" & CoorID & "'AND area_id ='" & Area_ID & "'"
                   rs.MoveFirst
                   Else
                   Unique_ID = CorID & "-" & AreaID & "-" & StrReverse(Mid(ConsID, Len(maxRec)))
                   Exit Function
            End If
            rs.MoveNext
        Loop Until rs.EOF
   Else
        ConsID = "0000"
        Mid(ConsID, Len(maxRec)) = StrReverse(Mid(maxRec, 1))
        Unique_ID = CorID & "-" & AreaID & "-" & StrReverse(Mid(ConsID, Len(maxRec)))
   End If
End Function

Public Function str_Filter(Text As TextBox, ascKey1 As Integer, ascKey2 As Integer, ascKey3 As Integer)
On Error Resume Next
'-----function dump all strings except
Dim Delimeter As String
Dim X As Long
Dim intStr As Variant
For X = 1 To Len(Text.Text)         'asckey1                                asckey2                              asckey3
   If Asc(Mid(Text.Text, X, 1)) >= ascKey1 And Asc(Mid(Text.Text, X, 1)) <= ascKey2 Or Asc(Mid(Text.Text, X, 1)) = ascKey3 Then
   Else
   Delimeter = Chr(Asc(Mid(Text.Text, X, 1)))
   End If
Next
intStr = ""
For X = 1 To Len(Text.Text)
    If Mid(Text.Text, X, 1) <> Delimeter Then
        intStr = intStr & Mid(Text.Text, X, 1)
        Else
        SendKeys "{end}"
    End If
Next
str_Filter = intStr
End Function

Public Function UnloadAllExceptOne(FormToStay As String)
Dim oFrm As Form
For Each oFrm In Forms
    If oFrm.Name <> FormToStay And Not _
             (TypeOf oFrm Is frmMain) And Not (TypeOf oFrm Is frmShortCuts) Then Unload oFrm
Next
'get formtostay
SelectedForm = UCase(FormToStay)
End Function

