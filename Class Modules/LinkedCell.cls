VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LinkedCell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Copyright https://github.com/ftrippel

Option Explicit

Private xsource As Range
Private xtarget As Range
    
Public Property Get id() As String
    id = xsource.Address & ":" & xtarget.Address
End Property

Public Property Get source() As Range
    Set source = xsource
End Property

Public Property Let source(value As Range)
    Set xsource = value
End Property

Public Property Get target() As Range
    Set target = xtarget
End Property

Public Property Let target(value As Range)
    Set xtarget = value
End Property

Public Sub TargetToSource()
    xsource.value = xtarget.value
    xsource.numberFormat = xtarget.numberFormat
End Sub

Public Sub SourceToTarget()
    xtarget.value = xsource.value
    xtarget.numberFormat = xsource.numberFormat
End Sub

