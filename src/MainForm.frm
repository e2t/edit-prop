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


'Written in 2014-2017 by Eduard E. Tikhenko <aquaried@gmail.com>
'
'To the extent possible under law, the author(s) have dedicated all copyright
'and related and neighboring rights to this software to the public domain
'worldwide. This software is distributed without any warranty.
'You should have received a copy of the CC0 Public Domain Dedication along
'with this software.
'If not, see <http://creativecommons.org/publicdomain/zero/1.0/>

Option Explicit

Dim PosConfItem As Integer
Private readOldAfterChecked As Boolean
Dim gItems As Dictionary
Dim gCutItems As Dictionary 'of Dictionary
Dim isShiftPressed As Boolean
Dim gCodeRegexPattern As String

Private Sub chkUpdateStd_Click()

  SaveBoolSetting "UpdateStd", Me.chkUpdateStd.value
  
End Sub

Private Sub lenLab_Click()

  Me.lenBox.text = ""
    
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

Private Sub tabConfAndCuts_Change()

  Me.ConfBox.Clear
  Select Case Me.tabConfAndCuts.value
    Case tabNumberConf
      InitWidgetFrom Me.ConfBox, gModelConfNames
      Me.ConfBox.ListIndex = indexLastConf
    Case tabNumberCuts
      InitWidgetFrom Me.ConfBox, gModelCutsNames
      Me.ConfBox.ListIndex = indexLastCut
  End Select
    
End Sub

Private Sub tabConfAndCuts_MouseDown(ByVal Index As Long, ByVal Button As Integer, _
                                     ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
                                     
  Select Case Me.tabConfAndCuts.value
    Case tabNumberConf
      indexLastConf = Me.ConfBox.ListIndex
    Case tabNumberCuts
      indexLastCut = Me.ConfBox.ListIndex
  End Select
    
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

Sub SetShiftStatus(Shift As Integer)
  
  isShiftPressed = Shift And 1
  
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

Private Sub ExitByKey(KeyCode As MSForms.ReturnInteger, Shift As Integer)

  If Shift = 1 And KeyCode = vbKeyReturn Then
    Execute
    Unload Me
  End If
    
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

Private Sub SetValueInBox(ByRef box As ComboBox, Index As Integer)

  If 0 <= Index And Index < box.ListCount Then
    box.text = box.list(Index)
  End If
    
End Sub

Private Sub SettingBut_Click()

  OpenSettingsFile
    
End Sub

Sub RewriteNameAndSign(source As String, conf As String)

  Dim Designation As String
  Dim Name As String
  Dim Code As String
  Dim I As Integer
  Dim IsCodeFound As Boolean
  
  Designation = ""
  Name = ""
  SplitNameAndSign source, conf, Designation, Name, Code
  SignBox.text = Designation
  NameBox.text = Name
  If gIsDrawing Then
    IsCodeFound = False
    I = 0
    While (I < MiniSignBox.ListCount) And (Not IsCodeFound)
      IsCodeFound = (StrComp(MiniSignBox.list(I), Code, vbTextCompare) = 0)
      If IsCodeFound Then
        MiniSignBox.ListIndex = I
      End If
      I = I + 1
    Wend
  End If
  
End Sub

Private Sub ModelNameLab_Click()

  RewriteNameAndSign ModelNameBox.text, ConfBox.text
End Sub

Private Sub DrawNameLab_Click()

  If gIsDrawing Then
    RewriteNameAndSign DrawNameBox.text, ConfBox.text
  End If
    
End Sub

' Без точек "." в наименовании
Sub SplitNameAndSign(line As String, conf As String, ByRef Designation As String, _
                     ByRef Name As String, ByRef Code As String)
                     
  Const flat As String = "SM-FLAT-PATTERN"
  Dim regexAsm As RegExp
  Dim regexPrt As RegExp
  Dim matches As Object
  Dim z As Variant
  
  Designation = line
  Name = line
  Code = ""
  
  Set regexAsm = New RegExp
  regexAsm.Pattern = "(.*\..*[0-9] *)(" + gCodeRegexPattern + ") ([^.]+)"
  regexAsm.IgnoreCase = True
  regexAsm.Global = True
  
  Set regexPrt = New RegExp
  regexPrt.Pattern = "(.*\.[^ ]+) ([^.]+)"
  regexPrt.IgnoreCase = True
  regexPrt.Global = True
  
  If regexAsm.Test(line) Then
    Set matches = regexAsm.Execute(line)
    Designation = Trim(matches(0).SubMatches(0))
    Code = matches(0).SubMatches(1)
    Name = Trim(matches(0).SubMatches(2))
  ElseIf regexPrt.Test(line) Then
    Set matches = regexPrt.Execute(line)
    Designation = Trim(matches(0).SubMatches(0))
    Name = Trim(matches(0).SubMatches(1))
  End If
  
  If conf Like "*" & flat Then
    conf = Left(conf, Len(conf) - Len(flat))
  End If
  Select Case conf
    Case "00", "По умолчанию"
      'pass
    Case Else
      SignChk.value = False ' running event
      Designation = Designation & "-" & conf
  End Select
End Sub

Private Sub UserForm_Activate()

  Set gItems = New Dictionary
  Set gCutItems = New Dictionary
  PosConfItem = -1
  isShiftPressed = False
  readOldAfterChecked = True
  InitWidgets
  ReadProp gModelManager, commonSpace, modelProps
  If gIsDrawing Then
    ReadProp gDrawManager, commonSpace, drawProps
  End If
  ConfBox.text = gCurConf
    
End Sub

' Устанавливает значения gItems из свойств, игнорируя существующие
Sub ReadProp(manager As CustomPropertyManager, conf As String, props() As String)

  Dim items As Dictionary
  Set items = SelectItems(conf)
  
  If items.Exists(conf) Then
  Else
    items.Add conf, New Dictionary
  End If
  
  Dim prop_v As Variant
  For Each prop_v In props
    Dim prop As String
    prop = prop_v
    If items(conf).Exists(prop) Then
    Else
      Dim item As DataItem
      Set item = New DataItem
       
      Dim raw As String: raw = ""
      Dim val As String: val = ""
      manager.Get4 prop, False, raw, val
      item.rawValue = raw
      item.value = val
      
      If conf <> commonSpace Then
        item.fromAll = (item.rawValue = "" And items(commonSpace)(prop).rawValue <> "")
      Else
        item.fromAll = True
      End If
       
      If prop = pMaterial Then
        item.newValue = item.value
      Else
        item.newValue = item.rawValue
      End If
      items(conf).Add prop, item
    End If
  Next
    
End Sub

Sub SetBoxValue2(chk As CheckBox, prop As String, conf As String)

  Dim items As Dictionary
  Set items = SelectItems(conf)

  Dim item As DataItem
  Set item = items(conf)(prop)
  If Not chk Is Nothing Then
    If chk.value <> item.fromAll Then
      chk.value = item.fromAll
    Else
      ChangeChecked prop
    End If
  Else
    ChangeChecked prop
  End If
    
End Sub

Sub ReloadForm(conf As String)

  readOldAfterChecked = False

  If gIsDrawing Then
    SetBoxValue2 Nothing, pShortDrawingType, commonSpace
    SetBoxValue2 Nothing, pOrganization, commonSpace
    SetBoxValue2 Nothing, pDrafter, commonSpace
    SetBoxValue2 Nothing, pChecking, commonSpace
  End If
  
  SetBoxValue2 DevelChk, pDesigner, conf
  SetBoxValue2 SignChk, pDesignation, conf
  SetBoxValue2 NameChk, pName, conf
  SetBoxValue2 FormatChk, pFormat, conf
  SetBoxValue2 NoteChk, pNote, conf
  SetBoxValue2 MassChk, pMass, conf
  
  If Not gIsAssembly Then
    SetBoxValue2 BlankChk, pBlank, conf
    SetBoxValue2 SizeChk, pSize, conf
    SetBoxValue2 Nothing, pMaterial, conf
    SetBoxValue2 lenChk, pLen, conf
    SetBoxValue2 widChk, pWid, conf
  End If
  readOldAfterChecked = True
    
End Sub

Private Sub ConfBox_Change()
    '''''''''Refactoring
    If ConfBox.text = "" Then Exit Sub
        
    If Me.tabConfAndCuts.value = tabNumberConf Then
        Dim items As Dictionary
        Set items = SelectItems(gCurConf)
        If items.Exists(gCurConf) Then
            ReadForm gCurConf
        End If

        gCurConf = ConfBox.text
        ReadProp gModelExt.CustomPropertyManager(gCurConf), gCurConf, modelProps
        ReloadForm gCurConf
    Else
        gCurConf = ConfBox.text
        ' gModel is PartDoc if the cuts
        Dim part As PartDoc
        Set part = gModel
        Dim cut As Feature
        Set cut = part.FeatureByName(gCurConf)
        ReadProp cut.CustomPropertyManager, commonSpace, modelProps
        ReloadForm commonSpace
    End If
End Sub

Function ExistsInCombo(box As ComboBox, value As String)

  ExistsInCombo = False
  Dim I As Variant
  For Each I In box.list
    If I = value Then
      ExistsInCombo = True
      Exit For
    End If
  Next
    
End Function

Sub FromAllChecked(chk As CheckBox, box As Object, prop As String, conf As String, _
                   fromAll As Boolean, SetFirstItem As Boolean)
                   
  Dim items As Dictionary
  Set items = SelectItems(conf)
  Dim cmb As ComboBox
  
  If readOldAfterChecked Then
    ReadBox box, chk, conf, prop, False
  End If
  If prop = pSize Then
    ChangeSizeEqual (conf)
  ElseIf prop = pMass Then
    ChangeMassEqual (conf)
  End If

  Dim value As String
  If fromAll Then
    value = items(commonSpace)(prop).newValue
  Else
    value = items(conf)(prop).newValue
  End If
  If SetFirstItem And value = "" And TypeOf box Is ComboBox Then
    Set cmb = box
    If cmb.ListCount > 0 Then
      value = cmb.list(0)
    End If
  End If
  
  If box.Enabled Then
    If TypeOf box Is ComboBox Then
      Set cmb = box
      If cmb.Style = fmStyleDropDownList Then
        If ExistsInCombo(cmb, value) Then
          SetComboInExistValue box, value
        ElseIf cmb.ListCount > 0 Then
          cmb.ListIndex = 0
        End If
      Else
        cmb.text = value
      End If
    Else
      box.text = value
    End If
  End If
  
End Sub

Sub SetComboInExistValue(ByRef box As Object, value As String)

  On Error Resume Next  ''''ПОДАВЛЕНИЕ ОШИБКИ для Гордиенко
  box.text = value
    
End Sub

Function SelectItems(conf As String) As Dictionary

  If Me.tabConfAndCuts.value = tabNumberCuts Then
    If gCutItems.Exists(conf) Then
    Else
      gCutItems.Add conf, New Dictionary
      gCutItems(conf).Add commonSpace, New Dictionary
    End If
    Set SelectItems = gCutItems(conf)
  Else
    Set SelectItems = gItems
  End If
    
End Function

' Устанавливает значения gItems из формы
' conf - конфигурация ИЛИ элемент списка вырезов
Sub ReadBox(box As Object, chk As CheckBox, conf As String, prop As String, forward As Boolean)
    Dim items As Dictionary
    Set items = SelectItems(conf)
    '''''''''''''''''''''''''''''''''
    If Not items.Exists(conf) Then
        items.Add conf, New Dictionary
    End If
    
    If Not items(conf).Exists(prop) Then
        items(conf).Add prop, New DataItem
    End If
    
    If chk Is Nothing And conf = commonSpace Then
        items(commonSpace)(prop).fromAll = True
        items(commonSpace)(prop).newValue = box.text
    ElseIf prop = pMaterial Then
        items(conf)(prop).fromAll = False
        items(conf)(prop).newValue = box.text   'уравнение MaterialEqual устанавливается в SetProp2
    Else
        items(conf)(prop).fromAll = chk.value
        If forward Then
            If chk.value Then
                items(commonSpace)(prop).newValue = box.text
            Else
                items(conf).item(prop).newValue = box.text
            End If
        Else
            If chk.value Then
                items(conf)(prop).newValue = box.text
            Else
                items(commonSpace)(prop).newValue = box.text
            End If
        End If
    End If
End Sub

Sub SetPropToAll(box As Object, chk As CheckBox, property As String)
    Dim I As Variant
    Dim conf As String
    
    ReadBox box, Nothing, commonSpace, property, True
    For Each I In gModelConfNames
        conf = I
        ReadBox box, chk, conf, property, True
    Next
End Sub

Private Sub ReadForm(conf As String)

  ReadBox NameBox, NameChk, conf, pName, True
  ReadBox NameBoxEN, NameChk, conf, pNameEN, True
  ReadBox NameBoxPL, NameChk, conf, pNamePL, True
  ReadBox NameBoxUA, NameChk, conf, pNameUA, True
  
  ReadBox DevelBox, DevelChk, conf, pDesigner, True
  ReadBox SignBox, SignChk, conf, pDesignation, True
  ReadBox FormatBox, FormatChk, conf, pFormat, True
  ReadBox NoteBox, NoteChk, conf, pNote, True
  ReadBox MassBox, MassChk, conf, pMass, True
  
  If gIsDrawing Then
    ReadBox MiniSignBox, Nothing, commonSpace, pShortDrawingType, True
    ReadBox OrgBox, Nothing, commonSpace, pOrganization, True
    ReadBox DraftBox, Nothing, commonSpace, pDrafter, True
    ReadBox CheckingBox, Nothing, commonSpace, pChecking, True
  End If
  
  If Not gIsAssembly Then
    ReadBox BlankBox, BlankChk, conf, pBlank, True
    ReadBox SizeBox, SizeChk, conf, pSize, True
    ReadBox MaterialBox, Nothing, conf, pMaterial, True
    ReadBox lenBox, lenChk, conf, pLen, True
    ReadBox widBox, widChk, conf, pWid, True
  End If
  
End Sub

Private Sub ChangeChecked(prop As String)

  Select Case prop
    Case pDesignation
      FromAllChecked SignChk, SignBox, pDesignation, gCurConf, SignChk.value, False
    Case pName
      FromAllChecked NameChk, NameBox, pName, gCurConf, NameChk.value, False
    Case pNameEN
      FromAllChecked NameChk, NameBoxEN, pNameEN, gCurConf, NameChk.value, False
    Case pNamePL
      FromAllChecked NameChk, NameBoxPL, pNamePL, gCurConf, NameChk.value, False
    Case pNameUA
      FromAllChecked NameChk, NameBoxUA, pNameUA, gCurConf, NameChk.value, False
    Case pBlank
      FromAllChecked BlankChk, BlankBox, pBlank, gCurConf, BlankChk.value, False
    Case pFormat
      FromAllChecked FormatChk, FormatBox, pFormat, gCurConf, FormatChk.value, False
    Case pNote
      FromAllChecked NoteChk, NoteBox, pNote, gCurConf, NoteChk.value, False
    Case pDesigner
      FromAllChecked DevelChk, DevelBox, pDesigner, gCurConf, DevelChk.value, True
    Case pSize
      FromAllChecked SizeChk, SizeBox, pSize, gCurConf, SizeChk.value, False
    Case pMass
      FromAllChecked MassChk, MassBox, pMass, gCurConf, MassChk.value, True
    Case pMaterial
      FromAllChecked Nothing, MaterialBox, pMaterial, gCurConf, False, False
    Case pOrganization
      FromAllChecked Nothing, OrgBox, pOrganization, gCurConf, True, True
    Case pDrafter
      FromAllChecked Nothing, DraftBox, pDrafter, gCurConf, True, True
    Case pChecking
      FromAllChecked Nothing, CheckingBox, pChecking, gCurConf, True, True
    Case pShortDrawingType
      FromAllChecked Nothing, MiniSignBox, pShortDrawingType, gCurConf, True, False
    Case pLen
      FromAllChecked lenChk, lenBox, pLen, gCurConf, lenChk.value, False
    Case pWid
      FromAllChecked widChk, widBox, pWid, gCurConf, widChk.value, False
  End Select
   
End Sub

Sub TrySetPropToAll(box As Object, chk As CheckBox, property As String)

  If isShiftPressed And Not chk.value Then
    isShiftPressed = False
    
    readOldAfterChecked = False
    chk.value = True
    readOldAfterChecked = True
    
    If MsgBox("Связать все конфигурации со значением?", vbYesNo) = vbYes Then
      SetPropToAll box, chk, property
    End If
  Else
    ChangeChecked property
  End If
  
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

Private Sub CreateBaseDesignation()

  Dim mainDesignation As String
  Dim resolvedValue As String
  Dim rawValue As String
  Dim wasResolved As Boolean
  
  If gModel.Extension.CustomPropertyManager(gMainConf).Get5(pDesignation, False, rawValue, resolvedValue, wasResolved) = swCustomInfoGetResult_NotPresent Then
    gModel.Extension.CustomPropertyManager("").Get5 pDesignation, False, rawValue, resolvedValue, wasResolved
  End If
  
  gBaseDesignation = GetBaseDesignation(resolvedValue)
  
End Sub

Private Sub Execute()

  ReadForm gCurConf
  
  gModel.SetReadOnlyState False  'must be first!
  
  WriteModelProperties
  ChangeMassUnits
  If gIsDrawing Then
    CreateBaseDesignation
    WriteDrawingProperties
    SetSpeedformat
    If Me.chkUpdateStd.value Then
      ReloadStandard
    End If
    ChangeLineStyles
  End If
  gDoc.ForceRebuild3 True
  TryRenameDraft DrawNameBox.text
    
End Sub

Private Sub ChangeMassEqual(conf As String)

  Dim items As Dictionary
  Set items = SelectItems(conf)
  
  If Not gIsUnnamed Then
    If Me.tabConfAndCuts.value = tabNumberCuts Then
      '''''''''''''''''''''MassBox.list(0) = Equal("SW-Mass", items(commonSpace)(pMass).fromAll, commonSpace, gNameModel)
    Else
      MassBox.list(0) = Equal("SW-Mass", items(conf)(pMass).fromAll, conf, gNameModel)
    End If
  End If
    
End Sub

Private Sub ChangeSizeEqual(conf As String)

  Dim items As Dictionary
  Set items = SelectItems(conf)
  
  If Not gIsUnnamed And Not gIsAssembly Then
    If Me.tabConfAndCuts.value = tabNumberCuts Then
      '''''''''''''''''''''
    Else
      SizeBox.list(0) = GetEquationThickness(conf, items(conf)(pSize).fromAll, gNameModel)
    End If
  End If
  
End Sub

Private Function MaterialEqual(conf As String) As String

  If Not gIsUnnamed And Not gIsAssembly Then
    MaterialEqual = Equal(pTrueMaterial, False, conf, gNameModel)
  Else
    MaterialEqual = ""
  End If
  
End Function

Private Sub SetMaterial(conf As String)

  Dim new_material As String
  new_material = gItems(conf)(pMaterial).newValue
  If new_material <> sEmpty And new_material <> "" Then
    gModel.SetMaterialPropertyName2 conf, materialdb, new_material  'it's method of PartDoc
  End If
    
End Sub

Private Sub SetSpeedformat()

  If RealFormatBox.text <> RealFormatBox.list(0) Then
    ReloadSheet (RealFormatBox.text)
  End If
    
End Sub

Private Sub OutputTypeAndName()

  ModelNameBox.text = ShortFileName(gNameModel)
  If gIsDrawing Then
    DrawNameBox.Enabled = True
    DrawNameLab.Enabled = True
    DrawNameBox.text = ShortFileName(ShortFileNameExt(gDoc.GetPathName))
  End If
  If gIsAssembly Then
    Controls("ModelNameLab").Caption = "Файл сборки"
  Else
    Controls("ModelNameLab").Caption = "Файл детали"
  End If
    
End Sub

Private Function CreateCodeRegexPattern() As String
  
  If userDrawingTypes.count > 0 Then
    CreateCodeRegexPattern = Join(userDrawingTypes.Keys, "|")
  Else
    CreateCodeRegexPattern = "СБ|МЧ|УЧ|РСБ"
  End If
  
End Function

Private Sub InitWidgets()

  OutputTypeAndName
  GetConfNames  'set gModelConfNames
  GetCutsNamesAndCount  'set gModelCutsNames and gModelCutsCount
  InitWidgetFrom ConfBox, gModelConfNames
  InitWidgetFrom DevelBox, userDesigner
  InitWidgetFrom FormatBox, userFormat
  InitWidgetFrom NameBox, userName
  InitWidgetFrom NoteBox, userNote
  
  If gIsUnnamed Then
    MassLab.Enabled = False
    MassBox.Enabled = False
    MassChk.Enabled = False
  Else
    MassBox.AddItem ("")
    InitWidgetFrom MassBox, userMass
  End If
  
  If gIsAssembly Then
    BlankBox.Enabled = False
    BlankLab.Enabled = False
    BlankChk.Enabled = False
    lenBox.Enabled = False
    lenLab.Enabled = False
    lenChk.Enabled = False
    widBox.Enabled = False
    widLab.Enabled = False
    widChk.Enabled = False
  Else
    InitWidgetFrom BlankBox, userBlank
    InitWidgetFrom lenBox, userLen
    InitWidgetFrom widBox, userWid
  End If

  If gIsUnnamed Or gIsAssembly Then
    SizeLab.Enabled = False
    SizeBox.Enabled = False
    SizeChk.Enabled = False
    MaterialBox.Enabled = False
    MaterialLab.Enabled = False
  Else
    SizeBox.AddItem ("")  ' for Equation
    InitWidgetFrom SizeBox, userSize

    Dim baseMaterials() As String
    baseMaterials = ReadMaterialNames("Материалы.sldmat")
    Dim resultMaterials() As String
    Dim I As Variant
    Dim k As Integer: k = 0
    
    For Each I In userMaterials
      Dim count As Integer
      count = IndexInArray(I, baseMaterials)
      If count <> -1 Then
        ReDim Preserve resultMaterials(k)
        resultMaterials(k) = I
        k = k + 1
        baseMaterials(count) = ""
      End If
    Next
    
    For Each I In baseMaterials   ' maybe nonunique items
      If I <> "" Then
        ReDim Preserve resultMaterials(k)
        resultMaterials(k) = I
        k = k + 1
      End If
    Next
    MaterialBox.AddItem sEmpty
    InitWidgetFrom MaterialBox, resultMaterials
  End If

  If gIsDrawing Then
    Me.chkUpdateStd.value = GetBoolSetting("UpdateStd")
    MiniSignBox.AddItem ""
    InitWidgetFrom MiniSignBox, userDrawingTypes.Keys
    InitWidgetFrom OrgBox, userOrganization
    InitWidgetFrom DraftBox, userDrafter
    InitWidgetFrom CheckingBox, userChecking
    InitRealFormatBox '''установка основных надписей
    gCodeRegexPattern = CreateCodeRegexPattern
  Else
    chkUpdateStd.Enabled = False
    MiniSignBox.Enabled = False
    MiniSignLab.Enabled = False
    OrgLab.Enabled = False
    OrgBox.Enabled = False
    DraftLab.Enabled = False
    DraftBox.Enabled = False
    RealFormatBox.Enabled = False
    RealFormatLab.Enabled = False
  End If
  
  Me.tabConfAndCuts.Tabs(0).Caption = "Конфигурации " & gModel.GetConfigurationCount
  If gModelCutsCount = 0 Then
  Else
    Me.tabConfAndCuts.Tabs.Add , "Вырезы " & gModelCutsCount
  End If
    
End Sub

Private Sub GetCutsNamesAndCount()

  Dim f As Feature
  Dim subf As Feature
  Dim storage As Dictionary
  Dim key As Variant
  Dim I As Integer
  
  If gIsAssembly Then
    gModelCutsCount = 0
  Else
    Set storage = New Dictionary
    Set f = gModel.FirstFeature
    Do
      If f Is Nothing Then
        Exit Do
      Else
        CheckTypeCut f, storage
        Set subf = f.GetFirstSubFeature
        Do
          If subf Is Nothing Then
            Exit Do
          Else
            CheckTypeCut subf, storage
            Set subf = subf.GetNextSubFeature
          End If
        Loop
        Set f = f.GetNextFeature
      End If
    Loop
    
    gModelCutsCount = storage.count
    If gModelCutsCount > 0 Then
      ReDim gModelCutsNames(gModelCutsCount - 1)
      I = 0
      For Each key In storage.Keys
        gModelCutsNames(I) = key
        I = I + 1
      Next
    End If
  End If
    
End Sub

Private Sub CheckTypeCut(aFeature As Feature, ByRef storage As Dictionary)

  If storage.Exists(aFeature.Name) Then
  ElseIf aFeature.GetTypeName2 = "CutListFolder" Then
    storage.Add aFeature.Name, 0
  End If
    
End Sub

Private Sub GetConfNames()

  Dim conf As Variant
  Dim I As Integer
  
  'FIX: gModel.GetConfigurationCount crash if flexible confs
  ReDim gModelConfNames(UBound(gModel.GetConfigurationNames) - LBound(gModel.GetConfigurationNames))
  
  I = 0
  For Each conf In BubbleSort(gModel.GetConfigurationNames)  'configurations list is not sorted
    gModelConfNames(I) = conf
    I = I + 1
  Next
    
End Sub

Private Function InitRealFormatBox() 'mask for button

  Dim filename As String
  Dim names() As String
  Dim I As Long
  
  RealFormatBox.AddItem ("<данная>")
  RealFormatBox.text = RealFormatBox.list(0)
  I = -1
  filename = Dir(gConfigPath & "*.SLDDRT")
  While filename <> ""
    I = I + 1
    ReDim Preserve names(0 To I)
    names(I) = ShortFileName(filename)
    filename = Dir()
  Wend
  names = SortSpeedFormats(names)
  While I >= 0
    RealFormatBox.AddItem names(I)
    I = I - 1
  Wend
    
End Function

Function SortSpeedFormats(names() As String) As String()

  Dim majorNames() As String
  Dim minorNames() As String
  Dim name_ As Variant
  Dim Name As String
  Dim n As Integer
  Dim j As Integer
  Dim I As Integer
  
  n = -1
  j = -1
  If Not IsArrayEmpty(names) Then
    For Each name_ In names
      Name = name_
      If Name Like "[aAаА]# *" Or Name Like "[aAаА]#" Then
        j = j + 1
        ReDim Preserve majorNames(j)
        majorNames(j) = Name
      Else
        n = n + 1
        ReDim Preserve minorNames(n)
        minorNames(n) = Name
      End If
    Next
    For I = 0 To n
      names(LBound(names) + I) = minorNames(I)
    Next
    For I = j To 0 Step -1
      names(UBound(names) - j + I) = majorNames(I)
    Next
  End If
  SortSpeedFormats = names
    
End Function

Sub SetModelProp(conf As String, prop As String, item As DataItem)

  Dim confManager As CustomPropertyManager
  Set confManager = gModelExt.CustomPropertyManager(conf)
  If conf <> commonSpace And item.fromAll Then
    confManager.Delete (prop)
  Else
    SetProp2 confManager, prop, item, conf
  End If
    
End Sub

Function SetProp2(manager As CustomPropertyManager, prop As String, item As DataItem, _
                  Optional conf As String = commonSpace) As Boolean
                  
  Dim result As Boolean
  result = False
  
  If prop = pMaterial Then
    If item.newValue <> sEmpty Then
      result = SetProp(manager, prop, MaterialEqual(conf))
    Else
      gModelExt.CustomPropertyManager(conf).Delete2 pMaterial
      gModelManager.Delete2 pMaterial
    End If
  Else
    result = SetProp(manager, prop, item.newValue)
  End If
  SetProp2 = result
    
End Function

Private Sub WriteModelProperties()

  Dim I As Variant
  Dim j As Variant
  Dim conf As String
  Dim prop As String
  Dim item As DataItem
  
  For Each I In gItems.Keys
    conf = I
    For Each j In modelProps
      prop = j
      Set item = gItems(conf)(prop)
      
      Select Case prop
        Case pBlank, pSize, pLen, pWid
          If Not gIsAssembly Then
            SetModelProp conf, prop, item
          End If
        Case pMaterial
          If Not gIsAssembly Then
            If Not gIsUnnamed Then
              SetModelProp conf, prop, item
            End If
            SetMaterial conf
          End If
        Case Else
          SetModelProp conf, prop, item
      End Select
      
    Next
  Next
    
End Sub

Private Sub WriteDrawingProperties()

  Dim toAll As Boolean: toAll = True
  Dim item As Dictionary: Set item = gItems(commonSpace)

  'см. массив drawProps
  SetProp2 gDrawManager, pShortDrawingType, item(pShortDrawingType)
  SetProp2 gDrawManager, pOrganization, item(pOrganization)
  SetProp2 gDrawManager, pDrafter, item(pDrafter)
  SetProp2 gDrawManager, pChecking, item(pChecking)  'before: userChecking(0)
  SetProp gDrawManager, pApprover, userApprover(0)
  SetProp gDrawManager, pNormControl, userNormControl(0)
  SetProp gDrawManager, pTechControl, userTechControl(0)
  SetProp gDrawManager, pLongDrawingType, userDrawingTypes(MiniSignBox.text)
  SetProp gDrawManager, pBaseDesignation, gBaseDesignation
    
End Sub

Private Sub CloseBut_Click()

  Dim options As swSaveAsOptions_e, errors As swFileSaveError_e, warnings As swFileSaveWarning_e
  If isShiftPressed Then
    gDoc.Save3 options, errors, warnings  ' отсутствует проверка сохранения
    gApp.CloseDoc (gDoc.GetPathName)
  End If
  Unload Me
    
End Sub

Private Sub ApplyBut_Click()

  Execute
  If isShiftPressed Then
    Unload Me
  End If
    
End Sub
