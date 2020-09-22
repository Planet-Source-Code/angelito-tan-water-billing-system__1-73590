VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim lst As ListItem

Public Function lstDatabase(ByVal query As String, lstview As ListView, imageindex As Integer) As Long
On Error Resume Next
Dim xSize As Integer
Dim lstCnt As Integer
Dim subCnt As Integer
lstview.ListItems.Clear
rsCheck
rs.Open query, CN, adOpenStatic, adLockOptimistic
lstview.Refresh
lstview.ColumnHeaders.Clear
For lstCnt = 1 To rs.Fields.Count
    lstview.ColumnHeaders.ADD lstCnt, , rs(lstCnt - 1).Name
    xSize = lstview.Width / rs.Fields.Count
    If xSize > 1440 Then
        lstview.ColumnHeaders.Item(lstCnt).Width = xSize
    Else
        lstview.ColumnHeaders.Item(lstCnt).Width = 2000
    End If
Next
If rs.RecordCount <> 0 Then
    rs.MoveFirst
    rs.MoveLast
    rs.MoveFirst
    For lstCnt = 1 To rs.RecordCount
        If rs(0).Type = adCurrency Then
            Set lst = lstview.ListItems.ADD(, , Format(CDbl(rs(0)), "##,###,###,##0.00"))
        Else
            Set lst = lstview.ListItems.ADD(, , rs(0), , imageindex)
        End If
        For subCnt = 1 To rs.Fields.Count - 1
            If rs(0) = Null Then
                lst.SubItems(subCnt) = "-----"
            Else
                If rs(subCnt).Type = adCurrency Then
                    lst.SubItems(subCnt) = Format(CDbl(rs(subCnt)), "##,###,###,##0.00")
                Else
                    lst.SubItems(subCnt) = rs(subCnt)
                End If
            End If
        Next subCnt
        rs.MoveNext
    Next lstCnt
    lstDatabase = rs.RecordCount
    Exit Function
End If
lstDatabase = rs.RecordCount
End Function

