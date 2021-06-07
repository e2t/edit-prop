VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "Редактор свойств"
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15030
   OleObjectBlob   =   "MainForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CodeBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

  ExitByKey KeyCode, Shift

End Sub

Private Sub lenLab_Click()

  Me.lenBox.text = ""
    
End Sub

Private Sub MiniSignBox_Change()

  Dim Key As String
  Dim I As Variant
  
  CodeBox.Clear
  Key = Me.MiniSignBox.text
  If UserDrawingTypes.Exists(Key) Then
    For Each I In UserDrawingTypes(Key)
      CodeBox.AddItem I
    Next
    CodeBox.text = CodeBox.List(0)
  End If

End Sub

Private Sub widLab_Click()

  Me.widBox.text = ""
    
End Sub

Private Sub MassChk_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

  SetShiftStatus Shift
    
End Sub

Private Sub SignChk_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

  SetShiftStatus Shift
  
End Sub


Private Sub NameChk_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

  SetShiftStatus Shift
    
End Sub

Private Sub BlankChk_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  
  SetShiftStatus Shift
  
End Sub

Private Sub SizeChk_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  
  SetShiftStatus Shift
  
End Sub

Private Sub FormatChk_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  
  SetShiftStatus Shift
  
End Sub

Private Sub lenChk_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  
  SetShiftStatus Shift
  
End Sub

Private Sub widChk_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

  SetShiftStatus Shift
  
End Sub

Private Sub NoteChk_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  
  SetShiftStatus Shift
  
End Sub

Private Sub DevelChk_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  
  SetShiftStatus Shift
  
End Sub

Private Sub CloseBut_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  
  SetShiftStatus Shift
  
End Sub

Private Sub BlankLab_Click()
  
  SetValueInBox BlankBox, 0
  
End Sub

Private Sub DevelLab_Click()
  
  SetValueInBox DevelBox, 0
  
End Sub

Private Sub DraftLab_Click()
  
  SetValueInBox DraftBox, 0
  
End Sub

Private Sub CheckingLab_Click()
  
  SetValueInBox CheckingBox, 0
  
End Sub

Private Sub FormatLab_Click()
  
  SetValueInBox FormatBox, 0
  
End Sub

Private Sub MaterialLab_Click()
  
  SetValueInBox MaterialBox, 1
  
End Sub

Private Sub NameLab_Click()

  If gIsDrawing Then
    DrawNameBox.value = DrawNameBox.value & " " & NameBox.value
  End If
    
End Sub

Private Sub NoteLab_Click()
  
  SetValueInBox NoteBox, 0
  
End Sub

Private Sub MassLab_Click()
  
  SetValueInBox MassBox, 0
  
End Sub

Private Sub MiniSignLab_Click()
  
  SetValueInBox MiniSignBox, 1
  
End Sub

Private Sub OrgLab_Click()
  
  SetValueInBox OrgBox, 0
  
End Sub

Private Sub SignBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  
  ExitByKey KeyCode, Shift
  
End Sub

Private Sub NameBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  
  ExitByKey KeyCode, Shift
  
End Sub

Private Sub NameBoxEN_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  
  ExitByKey KeyCode, Shift
  
End Sub

Private Sub NameBoxPL_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  
  ExitByKey KeyCode, Shift
  
End Sub

Private Sub NameBoxUA_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  
  ExitByKey KeyCode, Shift
  
End Sub

Private Sub BlankBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  
  ExitByKey KeyCode, Shift
  
End Sub

Private Sub SizeBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  
  ExitByKey KeyCode, Shift
  
End Sub

Private Sub MaterialBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  
  ExitByKey KeyCode, Shift
  
End Sub

Private Sub FormatBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  
  ExitByKey KeyCode, Shift
  
End Sub

Private Sub RealFormatBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  
  ExitByKey KeyCode, Shift
  
End Sub

Private Sub lenBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  
  ExitByKey KeyCode, Shift
  
End Sub

Private Sub widBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  
  ExitByKey KeyCode, Shift
  
End Sub

Private Sub MiniSignBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  
  ExitByKey KeyCode, Shift
  
End Sub

Private Sub NoteBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  
  ExitByKey KeyCode, Shift
  
End Sub

Private Sub MassBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  
  ExitByKey KeyCode, Shift
  
End Sub

Private Sub OrgBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  
  ExitByKey KeyCode, Shift
  
End Sub

Private Sub DevelBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  
  ExitByKey KeyCode, Shift
  
End Sub

Private Sub DraftBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  
  ExitByKey KeyCode, Shift
  
End Sub

Private Sub ModelNameBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  
  ExitByKey KeyCode, Shift
  
End Sub

Private Sub DrawNameBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  
  ExitByKey KeyCode, Shift
  
End Sub

Private Sub ConfBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  
  ExitByKey KeyCode, Shift
  
End Sub

Private Sub SizeLab_Click()
  
  SetValueInBox SizeBox, 0
  
End Sub

Private Sub SettingBut_Click()

  OpenSettingsFile
    
End Sub

Private Sub ModelNameLab_Click()

  RewriteNameAndSign ModelNameBox.text, ConfBox.text
  
End Sub

Private Sub DrawNameLab_Click()

  If gIsDrawing Then
    RewriteNameAndSign DrawNameBox.text, ConfBox.text
  End If
    
End Sub

Private Sub UserForm_Initialize()

  Set gItems = New Dictionary
  isShiftPressed = False
  readOldAfterChecked = True
  InitWidgets
  ReadProp gModelManager, commonSpace, modelProps
  If gIsDrawing Then
    ReadProp gDrawManager, commonSpace, drawProps
  End If
  If Not gIsAssembly Then
    Me.IsFastenerChk.value = GetIsFastener
  End If
  Me.ConfBox.text = gCurConf
    
End Sub

Private Sub ConfBox_Change()

  Dim part As PartDoc
  
  If ConfBox.text = "" Then Exit Sub
      
  If gItems.Exists(gCurConf) Then 'запись старой конфигурации
    ReadForm gCurConf
  End If

  gCurConf = ConfBox.text  'до этого в gCurConf записана старая конфигурация
  
  If Not gItems.Exists(gCurConf) Then
    gModel.ShowConfiguration2 gCurConf 'ускоряет чтение свойств
    ReadProp gModelExt.CustomPropertyManager(gCurConf), gCurConf, modelProps
  End If
  ReloadForm gCurConf
    
End Sub

Private Sub SignChk_Change()

  TrySetPropToAll SignBox, SignChk, pDesignation
  Me.ConfBox.SetFocus
  
End Sub

Private Sub NameChk_Change()

  TrySetPropToAll NameBox, NameChk, pName
  TrySetPropToAll NameBoxEN, NameChk, pNameEN
  TrySetPropToAll NameBoxPL, NameChk, pNamePL
  TrySetPropToAll NameBoxUA, NameChk, pNameUA
  Me.ConfBox.SetFocus
  
End Sub

Private Sub BlankChk_Change()

  TrySetPropToAll BlankBox, BlankChk, pBlank
  Me.ConfBox.SetFocus
  
End Sub

Private Sub FormatChk_Change()

  TrySetPropToAll FormatBox, FormatChk, pFormat
  Me.ConfBox.SetFocus
  
End Sub

Private Sub NoteChk_Change()

  TrySetPropToAll NoteBox, NoteChk, pNote
  Me.ConfBox.SetFocus
  
End Sub

Private Sub DevelChk_Change()

  TrySetPropToAll DevelBox, DevelChk, pDesigner
  Me.ConfBox.SetFocus
  
End Sub

Private Sub SizeChk_Change()

  TrySetPropToAll SizeBox, SizeChk, pSize
  Me.ConfBox.SetFocus
  
End Sub

Private Sub MassChk_Change()

  TrySetPropToAll MassBox, MassChk, pMass
  Me.ConfBox.SetFocus
  
End Sub
  
Private Sub lenChk_Change()

  TrySetPropToAll lenBox, lenChk, pLen
  Me.ConfBox.SetFocus
  
End Sub

Private Sub widChk_Change()

  TrySetPropToAll widBox, widChk, pWid
  Me.ConfBox.SetFocus
  
End Sub

Private Sub CloseBut_Click()

  Dim options As swSaveAsOptions_e, errors As swFileSaveError_e, warnings As swFileSaveWarning_e
  
  If isShiftPressed Then
    gDoc.Save3 options, errors, warnings  ' отсутствует проверка сохранения
    gApp.CloseDoc (gDoc.GetPathName)
  End If
  ExitApp
    
End Sub

Private Sub ApplyBut_Click()

  Execute
  If isShiftPressed Then
    ExitApp
  End If
    
End Sub
