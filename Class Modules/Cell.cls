VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Cell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' Copyright https://github.com/ftrippel

Option Explicit

Private m_value As Variant
Private m_numberFormat As Variant

Public Sub Init(r As Range)
    m_value = r.value
    m_numberFormat = r.numberFormat
End Sub

Public Property Get value() As Variant
    value = m_value
End Property

Public Property Let value(v As Variant)
    m_value = v
End Property

Public Property Get numberFormat() As Variant
    numberFormat = m_numberFormat
End Property

Public Property Let numberFormat(v As Variant)
    m_numberFormat = v
End Property

