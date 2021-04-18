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

Public value As String
Public rawValue As String
Public newValue As String
Public fromAll As Boolean

Public Sub Class_Initialize()

  value = ""
  rawValue = ""
  newValue = ""
  fromAll = False
    
End Sub
