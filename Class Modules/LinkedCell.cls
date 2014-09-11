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

Private m_source As Range
Private m_target As Range

Public Sub Init(source As Range, target As Range)
    Set m_source = source
    Set m_target = target
End Sub

Public Property Get id() As String
    id = m_source.Address & ":" & m_target.Address
End Property

Public Property Get source() As Range
    Set source = m_source
End Property

Public Property Let source(value As Range)
    Set m_source = value
End Property

Public Property Get target() As Range
    Set target = m_target
End Property

Public Property Let target(value As Range)
    Set m_target = value
End Property

Public Sub TargetToSource()
    m_source.value = m_target.value
    m_source.numberFormat = m_target.numberFormat
End Sub

Public Sub SourceToTarget()
    m_target.value = m_source.value
    m_target.numberFormat = m_source.numberFormat
End Sub

