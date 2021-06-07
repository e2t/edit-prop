Attribute VB_Name = "Settings"
Option Explicit

Sub SaveSetting2(ByRef Key As String, ByRef value As String)

  SaveSetting macroName, macroSection, Key, value
    
End Sub


Sub SaveIntSetting(ByRef Key As String, value As Integer)

  SaveSetting2 Key, str(value)
  
End Sub


Sub SaveBoolSetting(ByRef Key As String, value As Boolean)

  SaveSetting2 Key, BoolToStr(value)
  
End Sub


Function GetSetting2(ByRef Key As String) As String

  GetSetting2 = GetSetting(macroName, macroSection, Key, "0")
  
End Function


Function GetBoolSetting(ByRef Key As String) As Boolean

  GetBoolSetting = StrToBool(GetSetting2(Key))
  
End Function


Function GetIntSetting(ByRef Key As String) As Integer

  GetIntSetting = StrToInt(GetSetting2(Key))
  
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
