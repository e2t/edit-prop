Attribute VB_Name = "Settings"
Option Explicit

Sub SaveSetting2(ByRef key As String, ByRef value As String)
    SaveSetting macroName, macroSection, key, value
End Sub

Sub SaveIntSetting(ByRef key As String, value As Integer)
    SaveSetting2 key, str(value)
End Sub

Sub SaveBoolSetting(ByRef key As String, value As Boolean)
    SaveSetting2 key, BoolToStr(value)
End Sub

Function GetSetting2(ByRef key As String) As String
    GetSetting2 = GetSetting(macroName, macroSection, key, "0")
End Function

Function GetBoolSetting(ByRef key As String) As Boolean
    GetBoolSetting = StrToBool(GetSetting2(key))
End Function

Function GetIntSetting(ByRef key As String) As Integer
    GetIntSetting = StrToInt(GetSetting2(key))
End Function

Function StrToInt(ByRef value As String) As Integer
    If IsNumeric(value) Then
        StrToInt = CInt(value)
    Else
        StrToInt = 0
    End If
End Function

Function StrToBool(ByRef value As String) As Boolean
    If IsNumeric(value) Then
        StrToBool = CInt(value)
    Else
        StrToBool = False
    End If
End Function

Function BoolToStr(value As Boolean) As String
    BoolToStr = str(CInt(value))
End Function
