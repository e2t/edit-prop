Attribute VB_Name = "Settings"
Option Explicit

Const kLangTranslate = "LangTranslate"

Sub SaveSetting2(ByRef Key As String, ByRef Value As String)

  SaveSetting MacroName, MacroSection, Key, Value
    
End Sub

Sub SaveIntSetting(ByRef Key As String, Value As Integer)

  SaveSetting2 Key, Str(Value)
  
End Sub

Sub SaveBoolSetting(ByRef Key As String, Value As Boolean)

  SaveSetting2 Key, BoolToStr(Value)
  
End Sub

Function GetSetting2(ByRef Key As String) As String

  GetSetting2 = GetSetting(MacroName, MacroSection, Key, "0")
  
End Function

Function GetBoolSetting(ByRef Key As String) As Boolean

  GetBoolSetting = StrToBool(GetSetting2(Key))
  
End Function

Function GetIntSetting(ByRef Key As String) As Integer

  GetIntSetting = StrToInt(GetSetting2(Key))
  
End Function

Function StrToInt(ByRef Value As String) As Integer

  If IsNumeric(Value) Then
    StrToInt = CInt(Value)
  Else
    StrToInt = 0
  End If
  
End Function

Function StrToBool(ByRef Value As String) As Boolean

  If IsNumeric(Value) Then
    StrToBool = CInt(Value)
  Else
    StrToBool = False
  End If
  
End Function

Function BoolToStr(Value As Boolean) As String

  BoolToStr = Str(CInt(Value))
    
End Function

Function GetLangTranslateSetting() As Integer

  GetLangTranslateSetting = GetIntSetting(kLangTranslate)

End Function

Sub SaveLangTranslate(Index As Integer)

  SaveIntSetting kLangTranslate, Index

End Sub
