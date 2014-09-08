VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Table"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Copyright https://github.com/ftrippel

Option Explicit

Private target As Worksheet

' Normal Columns
Private cols As New Dictionary
Private dataCol As Range
Private valueCol As Range

' Linked Cells (lcs)
Private lcs As New Dictionary

Public Sub AddColumn(col As Range, colType As ColumnType)
    If colType = ColumnType.Normal Then
        cols.Add cols.Count + 1, col
    ElseIf colType = ColumnType.data Then
        Set dataCol = col
    ElseIf colType = ColumnType.value Then
        Set valueCol = col
    End If
End Sub

Public Sub Build(sheetName As String)

    Set target = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    target.Name = sheetName
    
    Speed (True)
    
    Dim r As Range
    Dim c As Range
    Dim i As Integer
    Dim v As Variant
    
    Dim rows As Range
    Set rows = Range(cols.Items(1).Cells(2, 1), cols.Items(1).Cells(2, 1).End(xlDown))
    
    Dim values As Range
    Set values = Range(valueCol.Cells(2, 1), valueCol.Cells(rows.Count, 1))
    values.AdvancedFilter Action:=xlFilterCopy, CopyToRange:=target.Cells(1, 1), Unique:=True
    
    Dim valuesTmp As Range
    Dim cell As cell
    Set valuesTmp = Range(target.Cells(1, 1), target.Cells(1, 1).End(xlDown))
    valuesTmp.Sort key1:=valuesTmp, order1:=xlAscending, Header:=xlNo
    Dim valueLookup As New Dictionary
    Dim cellLookup As New Dictionary
    i = 1
    For Each c In valuesTmp
        Set cell = New cell
        cell.value = c.value
        cell.numberFormat = c.numberFormat
        valueLookup.Add cell.value, i
        cellLookup.Add i, cell
        i = i + 1
    Next
    target.Cells.Clear
        
    i = 1
    For Each v In cols.Items
        target.Cells(1, i).value = v.Cells(1, 1).value
        i = i + 1
    Next v
    For Each v In valueLookup.Keys
            i = valueLookup.item(v)
            Set cell = cellLookup.item(i)
            target.Cells(1, cols.Count + i).value = cell.value
            target.Cells(1, cols.Count + i).numberFormat = cell.numberFormat
    Next
    
    lcs.RemoveAll
    Dim lc As LinkedCell
    Dim k As Integer
    For k = 2 To rows.Count + 1
        i = 1
        For Each v In cols.Items
            Set c = v.Cells(k, 1)
            target.Cells(k, i).value = c.value
            target.Cells(k, i).numberFormat = c.numberFormat
            i = i + 1
        Next v
        
        Dim j As Integer
        j = valueLookup.item(valueCol.Cells(k, 1).value)
        
        target.Cells(k, cols.Count + j).Locked = False
        
        Set lc = New LinkedCell
        lc.source = dataCol.Cells(k, 1)
        lc.target = target.Cells(k, cols.Count + j)
        If Not lcs.Exists(lc.id) Then
            lcs.Add lc.id, lc
        End If
            
    Next k
    
    PushValues
    
    Range(target.Cells(1, 1), target.Cells(1, cols.Count)).AutoFilter
    target.Columns.AutoFit
    
    Speed (False)

End Sub

Public Sub Protect(pwd As String)
    target.Protect password:=pwd, AllowFiltering:=True, AllowSorting:=True
End Sub

Public Sub PushValues()
    Speed (True)
    Dim k As Variant
    For Each k In lcs.Items
        k.SourceToTarget
    Next
    Speed (False)
End Sub

Public Sub PullValues()
    Speed (True)
    Dim k As Variant
    For Each k In lcs.Items
        k.TargetToSource
    Next
    Speed (False)
End Sub

Public Sub Destroy()
    Speed True
    Application.DisplayAlerts = False
    target.Delete
    Application.DisplayAlerts = True
End Sub
