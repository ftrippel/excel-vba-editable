Attribute VB_Name = "MLinkedCell"

' Copyright https://github.com/ftrippel

Option Explicit

Public Function CreateLinkedCell(source As Range, target As Range) As LinkedCell
    Set CreateLinkedCell = New LinkedCell
    CreateLinkedCell.Init source, target
End Function

