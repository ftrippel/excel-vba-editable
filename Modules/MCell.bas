Attribute VB_Name = "MCell"

' Copyright https://github.com/ftrippel

Option Explicit

Public Function CreateCell(r As Range) As cell
    Set CreateCell = New cell
    CreateCell.Init r
End Function
