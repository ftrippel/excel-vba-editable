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

Private xvalue As Variant
Private xnumberFormat As String

Public Property Get value() As Variant
    value = xvalue
End Property

Public Property Let value(v As Variant)
    xvalue = v
End Property

Public Property Get numberFormat() As String
    numberFormat = xnumberFormat
End Property

Public Property Let numberFormat(v As String)
    xnumberFormat = v
End Property


