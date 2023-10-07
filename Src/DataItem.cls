VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Value As String
Public RawValue As String
Public NewValue As String
Public FromAll As Boolean

Public Sub Class_Initialize()
    Value = ""
    RawValue = ""
    NewValue = ""
    FromAll = False
End Sub
