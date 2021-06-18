Attribute VB_Name = "Main"
Option Explicit

Public Const macroName As String = "EditProp"
Public Const macroSection As String = "Main"

'Свойства модели, записаны в массиве modelProps
Public Const pDesignation As String = "Обозначение"
Public Const pMaterial As String = "Материал"
Public Const pName As String = "Наименование"
Public Const pNameEN As String = "Наименование EN"
Public Const pNamePL As String = "Наименование PL"
Public Const pNameUA As String = "Наименование UA"
Public Const pBlank As String = "Заготовка"
Public Const pSize As String = "Типоразмер"
Public Const pNote As String = "Примечание"
Public Const pDesigner As String = "Разработал"
Public Const pFormat As String = "Формат"
Public Const pMass As String = "Масса"
Public Const pLen As String = "Длина"
Public Const pWid As String = "Ширина"
'Специальное свойство для получения материала
Public Const pTrueMaterial As String = "SW-Material"
'Специальное свойство деталей
Public Const pIsFastener As String = "IsFastener"
Public Const IsFastenerTrue = "1"
'Свойства чертежа, записаны в массиве drawProps
Public Const pDrafter As String = "Начертил"
Public Const pShortDrawingType As String = "Пометка"
Public Const pLongDrawingType As String = "Тип документа"
Public Const pOrganization As String = "Организация"
Public Const pChecking As String = "Проверил"
Public Const pApprover As String = "Утвердил"
Public Const pTechControl As String = "Техконтроль"
Public Const pNormControl As String = "Нормоконтроль"
Public Const pBaseDesignation As String = "Базовое обозначение"

Public Const materialdb As String = "Материалы"
Public Const commonSpace As String = ""
Public Const Separator As String = ";"
Public Const Separator2 As String = "="
Public Const Separator3 As String = ","
Public Const SettingsFile As String = "Настройки.txt"
Public Const ppExclude As String = "Исключить"
Public Const sEmpty = " "

Enum ErrorCode
  ok = 0
  emptyView = 1
  emptySheet = 2
End Enum

Public gApp As Object
Public gDoc As ModelDoc2
Public gConfigPath As String
Public gModel As ModelDoc2
Public gModelConfNames() As String
Public gDrawing As DrawingDoc
Public gModelExt As ModelDocExtension
Public gDrawExt As ModelDocExtension
Public gSheet As Sheet
Public gCurConf As String 'выбранная в списке конфигурация
Public gMainConf As String 'основная конфигурация на чертеже
Public gBaseDesignation As String
Public gChangeNumber As Long
Public gIsAssembly As Boolean
Public gIsUnnamed As Boolean
Public gIsDrawing As Boolean
Public gNameModel As String
Public gModelManager As CustomPropertyManager
Public gDrawManager As CustomPropertyManager
Public gFSO As FileSystemObject

Public userName() As String
Public UserDrawingTypes As Dictionary
Public userBlank() As String
Public userSize() As String
Public userFormat() As String
Public userNote() As String
Public userMass() As String
Public userOrganization() As String
Public userDesigner() As String
Public userDrafter() As String
Public userChecking() As String
Public userApprover() As String
Public userTechControl() As String
Public userNormControl() As String
Public userLen() As String
Public userWid() As String
Public userMaterials() As String
Public userPreExclude() As String
Public modelProps(13) As String
Public drawProps(8) As String

Public gItems As Dictionary
Public readOldAfterChecked As Boolean
Public isShiftPressed As Boolean

Dim gCodeRegexPattern As String
Dim gRegexMaterial As RegExp

Sub Main()

  Init
  Set gDoc = gApp.ActiveDoc
  If gDoc Is Nothing Then
    MsgBox "Нет открытых документов."
  Else
    EditorRun
  End If
    
End Sub

Function Init() As Boolean

  Set gApp = Application.SldWorks
  Set gFSO = New FileSystemObject
  
  ReDim userName(0)
  userName(0) = ""
  
  Set UserDrawingTypes = New Dictionary
  
  ReDim userBlank(0)
  userBlank(0) = ""
  ReDim userSize(0)
  userSize(0) = ""
  ReDim userFormat(0)
  userFormat(0) = ""
  ReDim userNote(0)
  userNote(0) = ""
  ReDim userMass(0)
  userMass(0) = ""
  ReDim userOrganization(0)
  userOrganization(0) = ""
  ReDim userDesigner(0)
  userDesigner(0) = ""
  ReDim userDrafter(0)
  userDrafter(0) = ""
  ReDim userChecking(0)
  userChecking(0) = ""
  ReDim userApprover(0)
  userApprover(0) = ""
  ReDim userTechControl(0)
  userTechControl(0) = ""
  ReDim userNormControl(0)
  userNormControl(0) = ""
  ReDim userWid(0)
  userWid(0) = ""
  ReDim userLen(0)
  userLen(0) = ""
  ReDim userMaterials(0)
  userMaterials(0) = ""
  ReDim userPreExclude(0)
  userPreExclude(0) = ""
  
  modelProps(0) = pDesignation
  modelProps(1) = pMaterial
  modelProps(2) = pName
  modelProps(3) = pBlank
  modelProps(4) = pSize
  modelProps(5) = pNote
  modelProps(6) = pDesigner
  modelProps(7) = pFormat
  modelProps(8) = pMass
  modelProps(9) = pLen
  modelProps(10) = pWid
  modelProps(11) = pNameEN
  modelProps(12) = pNamePL
  modelProps(13) = pNameUA
  
  drawProps(0) = pDrafter
  drawProps(1) = pShortDrawingType
  drawProps(2) = pLongDrawingType
  drawProps(3) = pOrganization
  drawProps(4) = pChecking
  drawProps(5) = pApprover
  drawProps(6) = pTechControl
  drawProps(7) = pNormControl
  drawProps(8) = pBaseDesignation
  
  SetConfigFolder
  ReadSettings  'only after SetConfigFolder
  
  Set gRegexMaterial = New RegExp
  gRegexMaterial.Global = True
  gRegexMaterial.MultiLine = True
  gRegexMaterial.IgnoreCase = True
  gRegexMaterial.Pattern = "material name=""([^""]+)"""
    
End Function

Function SetConfigFolder() As Boolean

  gConfigPath = gApp.GetCurrentMacroPathFolder() + "\config\"
  
End Function

Function IsDrawing(doc As ModelDoc2) As Boolean

  IsDrawing = CBool(doc.GetType = swDocumentTypes_e.swDocDRAWING)
    
End Function

Function IsPropertyExist(manager As CustomPropertyManager, prop As String) As Boolean

  Dim Names As Variant
  
  Names = manager.GetNames()
  IsPropertyExist = False
  If Not IsEmpty(Names) Then
    If IndexInArray(prop, Names) <> -1 Then
      IsPropertyExist = True
    End If
  End If
  
End Function

Function SetProp(manager As CustomPropertyManager, prop As String, value As String) As Boolean

  manager.Add2 prop, swCustomInfoText, ""
  SetProp = Not CBool(manager.Set(prop, Trim(value)))
  
End Function

Function IsArrayEmpty(anArray As Variant) As Boolean

  Dim I As Integer

  On Error GoTo ArrayIsEmpty
  IsArrayEmpty = LBound(anArray) > UBound(anArray)
  Exit Function
  
ArrayIsEmpty:
  IsArrayEmpty = True

End Function

Sub InitWidgetFrom(widget As Object, values As Variant)

  Dim I As Variant
  
  If Not IsArrayEmpty(values) Then
    For Each I In values
      widget.AddItem I
    Next
  End If
  
End Sub

Function InitWidgets() 'hide

  Dim baseMaterials As Dictionary
  Dim resultMaterials() As String
  Dim I As Variant
  Dim k As Integer

  OutputTypeAndName
  GetConfNames  'set gModelConfNames
  InitWidgetFrom MainForm.ConfBox, gModelConfNames
  InitWidgetFrom MainForm.DevelBox, userDesigner
  InitWidgetFrom MainForm.FormatBox, userFormat
  InitWidgetFrom MainForm.NameBox, userName
  InitWidgetFrom MainForm.NoteBox, userNote
  
  If gIsUnnamed Then
    MainForm.MassLab.Enabled = False
    MainForm.MassBox.Enabled = False
    MainForm.MassChk.Enabled = False
  Else
    MainForm.MassBox.AddItem ("")
    InitWidgetFrom MainForm.MassBox, userMass
  End If
  
  If gIsAssembly Then
    MainForm.BlankBox.Enabled = False
    MainForm.BlankLab.Enabled = False
    MainForm.BlankChk.Enabled = False
    MainForm.lenBox.Enabled = False
    MainForm.lenLab.Enabled = False
    MainForm.lenChk.Enabled = False
    MainForm.widBox.Enabled = False
    MainForm.widLab.Enabled = False
    MainForm.widChk.Enabled = False
    MainForm.IsFastenerChk.Enabled = False
  Else
    InitWidgetFrom MainForm.BlankBox, userBlank
    InitWidgetFrom MainForm.lenBox, userLen
    InitWidgetFrom MainForm.widBox, userWid
  End If

  If gIsUnnamed Or gIsAssembly Then
    MainForm.SizeLab.Enabled = False
    MainForm.SizeBox.Enabled = False
    MainForm.SizeChk.Enabled = False
    MainForm.MaterialBox.Enabled = False
    MainForm.MaterialLab.Enabled = False
  Else
    MainForm.SizeBox.AddItem ("")  ' for Equation
    InitWidgetFrom MainForm.SizeBox, userSize
    Set baseMaterials = ReadMaterialNames("Материалы.sldmat")
    If baseMaterials.Count > 0 Then
      ReDim resultMaterials(baseMaterials.Count - 1)
      k = 0
      For Each I In userMaterials
        If baseMaterials.Exists(I) Then
          resultMaterials(k) = I
          k = k + 1
          baseMaterials.Remove I
        End If
      Next
      
      For Each I In baseMaterials.Keys
        resultMaterials(k) = I
        k = k + 1
      Next
    End If
    MainForm.MaterialBox.AddItem sEmpty
    InitWidgetFrom MainForm.MaterialBox, resultMaterials
  End If

  If gIsDrawing Then
    MainForm.MiniSignBox.AddItem ""
    InitWidgetFrom MainForm.MiniSignBox, UserDrawingTypes.Keys
    InitWidgetFrom MainForm.OrgBox, userOrganization
    InitWidgetFrom MainForm.DraftBox, userDrafter
    InitWidgetFrom MainForm.CheckingBox, userChecking
    InitRealFormatBox '''установка основных надписей
    gCodeRegexPattern = CreateCodeRegexPattern
  Else
    MainForm.MiniSignBox.Enabled = False
    MainForm.CodeBox.Enabled = False
    MainForm.MiniSignLab.Enabled = False
    MainForm.OrgLab.Enabled = False
    MainForm.OrgBox.Enabled = False
    MainForm.DraftLab.Enabled = False
    MainForm.DraftBox.Enabled = False
    MainForm.RealFormatBox.Enabled = False
    MainForm.RealFormatLab.Enabled = False
    MainForm.CheckingLab.Enabled = False
    MainForm.CheckingBox.Enabled = False
  End If
  
  MainForm.tabConf.Tabs(0).Caption = "Конфигурации " & gModel.GetConfigurationCount
    
End Function

Function Equal(swProp As String, toAll As Boolean, Conf As String, nameModel As String) As String

  Dim confText As String
  
  confText = ""
  If Conf <> commonSpace And Not toAll Then
    confText = "@@" + Conf
  End If
  Equal = """" + swProp + confText + "@" + nameModel + """"
    
End Function

Sub ReadHeaderValues(ByRef userStr() As String, ByRef Count As Long, Lines() As String, EndLines As Long)

  If Count < EndLines Then
    Count = Count + 1
    If Lines(Count) <> "" Then
      userStr = Split(Lines(Count), Separator)
    End If
  End If
    
End Sub

Sub ReadDrawingTypes(ByRef Count As Long, Lines() As String, EndLines As Long)

  Dim X As Variant
  Dim IndexSeparator As Integer
  Dim Names() As String
  Dim ShortName As String
  Dim LongNames() As String
  
  If Count < EndLines Then
    Count = Count + 1
    If Lines(Count) <> "" Then
      For Each X In Split(Lines(Count), Separator)  'String()
        IndexSeparator = InStr(X, Separator2)  'равен нулю, если разделитель не найден
        If IndexSeparator > 0 Then
          Names = Split(X, Separator2)
          ShortName = Names(0)
          LongNames = Split(Names(1), Separator3)
        Else
          ShortName = X
          Erase LongNames
        End If
        If Not UserDrawingTypes.Exists(ShortName) Then
          UserDrawingTypes.Add ShortName, LongNames
        End If
      Next
    End If
  End If
    
End Sub

Sub MsgDebug(msg As Variant)

  If MsgBox(msg, vbOKCancel) = vbCancel Then ExitApp
  
End Sub

Function ReadSettings() As Boolean

  Dim Lines() As String
  Dim EndLines As Long
  Dim I As Long
  Dim FStream As TextStream
  Dim FullPath As String
  
  FullPath = gConfigPath + SettingsFile
  If gFSO.FileExists(FullPath) Then
    Set FStream = gFSO.OpenTextFile(FullPath, ForReading, False, TristateTrue)
    Lines = Split(FStream.ReadAll, vbNewLine)
    FStream.Close
  
    EndLines = UBound(Lines)
    I = 0
    While I <= EndLines
      Select Case Lines(I)
        Case HeaderInFile(pName)
          ReadHeaderValues userName, I, Lines, EndLines
        Case HeaderInFile(pShortDrawingType)
          ReadDrawingTypes I, Lines, EndLines
        Case HeaderInFile(pSize)
          ReadHeaderValues userSize, I, Lines, EndLines
        Case HeaderInFile(pBlank)
          ReadHeaderValues userBlank, I, Lines, EndLines
        Case HeaderInFile(pDesigner)
          ReadHeaderValues userDesigner, I, Lines, EndLines
        Case HeaderInFile(pDrafter)
          ReadHeaderValues userDrafter, I, Lines, EndLines
        Case HeaderInFile(pFormat)
          ReadHeaderValues userFormat, I, Lines, EndLines
        Case HeaderInFile(pOrganization)
          ReadHeaderValues userOrganization, I, Lines, EndLines
        Case HeaderInFile(pMass)
          ReadHeaderValues userMass, I, Lines, EndLines
        Case HeaderInFile(pNote)
          ReadHeaderValues userNote, I, Lines, EndLines
        Case HeaderInFile(pChecking)
          ReadHeaderValues userChecking, I, Lines, EndLines
        Case HeaderInFile(pApprover)
          ReadHeaderValues userApprover, I, Lines, EndLines
        Case HeaderInFile(pTechControl)
          ReadHeaderValues userTechControl, I, Lines, EndLines
        Case HeaderInFile(pNormControl)
          ReadHeaderValues userNormControl, I, Lines, EndLines
        Case HeaderInFile(pLen)
          ReadHeaderValues userLen, I, Lines, EndLines
        Case HeaderInFile(pWid)
          ReadHeaderValues userWid, I, Lines, EndLines
        Case HeaderInFile(pMaterial)
          ReadHeaderValues userMaterials, I, Lines, EndLines
        Case HeaderInFile(ppExclude)
          ReadHeaderValues userPreExclude, I, Lines, EndLines
      End Select
      I = I + 1
    Wend
  End If
    
End Function

Function HeaderInFile(prop As String) As String

  HeaderInFile = "[" + prop + "]"
  
End Function

Function OpenSettingsFile() As Boolean

  Dim cmd As String
  Dim filename As String
  Dim text As String
  Dim fsT As Object
  Dim DrawingCodes(13) As String
  
  DrawingCodes(0) = ".AD=Сборочный чертеж,Assembly Drawing,Складальний кресленик"
  DrawingCodes(1) = ".ID=Монтажный чертеж,Installation Drawing,Монтажний кресленик"
  DrawingCodes(2) = ".DD=Габаритный чертеж,Dimension Drawing,Габаритний кресленик"
  DrawingCodes(3) = ".GA=Чертеж общего вида,General Arrangement Drawing,Кресленик загального виду"
  DrawingCodes(4) = ".TD=Чертеж 3D,3D-Drawing"
  DrawingCodes(5) = ".MD=Чертеж компонента,Component Drawing"
  DrawingCodes(6) = ".ND=Компоновочный чертеж,Arrangement Drawing"
  DrawingCodes(7) = ".CD=Концептуальный чертеж,Concept Drawing"
  DrawingCodes(8) = ".LD=Чертеж размещения,Layout Drawing"
  DrawingCodes(9) = ".ED=Разнесенный чертеж,Exploded-view Drawing"
  DrawingCodes(10) = "СБ=Сборочный чертеж,Складальний кресленик"
  DrawingCodes(11) = "МЧ=Монтажный чертеж,Монтажний кресленик"
  DrawingCodes(12) = "ГЧ=Габаритный чертеж,Габаритний кресленик"
  DrawingCodes(13) = "ВО=Чертеж общего вида,Кресленик загального виду"
  
  filename = gConfigPath + SettingsFile
  If Not gFSO.FileExists(filename) Then
    If Not gFSO.FolderExists(gConfigPath) Then
      gFSO.CreateFolder gConfigPath
    End If
      
    text = _
      HeaderInFile(pBlank) + vbNewLine + ";;;" + vbNewLine + vbNewLine + _
      HeaderInFile(pMass) + vbNewLine + "см. табл." + vbNewLine + vbNewLine + _
      HeaderInFile(pName) + vbNewLine + ";;;" + vbNewLine + vbNewLine + _
      HeaderInFile(pNormControl) + vbNewLine + "Юриков" + vbNewLine + vbNewLine + _
      HeaderInFile(pDrafter) + vbNewLine + ";;;" + vbNewLine + vbNewLine + _
      HeaderInFile(pOrganization) + vbNewLine + "ООО ""Эко-Инвест"";ЗАО НПФ ""Экотон""" + vbNewLine + vbNewLine + _
      HeaderInFile(pShortDrawingType) + vbNewLine + _
      Join(DrawingCodes, Separator) + vbNewLine + vbNewLine + _
      HeaderInFile(pNote) + vbNewLine + ";;;" + vbNewLine + vbNewLine + _
      HeaderInFile(pChecking) + vbNewLine + "Юриков" + vbNewLine + vbNewLine + _
      HeaderInFile(pDesigner) + vbNewLine + ";;;" + vbNewLine + vbNewLine + _
      HeaderInFile(pTechControl) + vbNewLine + "Гуменный" + vbNewLine + vbNewLine + _
      HeaderInFile(pSize) + vbNewLine + ";;;" + vbNewLine + vbNewLine + _
      HeaderInFile(pApprover) + vbNewLine + "Гуменный" + vbNewLine + vbNewLine + _
      HeaderInFile(pFormat) + vbNewLine + "А4;А3;А2;А1;А0;БЧ;*;А4х3;А4х4;А3х3;А3х4" + vbNewLine + vbNewLine + _
      HeaderInFile(pMaterial) + vbNewLine + "AISI 304;Ст.3;EPDM" + vbNewLine + vbNewLine + _
      HeaderInFile(pLen) + vbNewLine + ";;;" + vbNewLine + vbNewLine + _
      HeaderInFile(pWid) + vbNewLine + ";;;" + vbNewLine + vbNewLine + _
      HeaderInFile(ppExclude) + vbNewLine + ";;;" + vbNewLine
      
    Set fsT = CreateObject("ADODB.Stream")
    fsT.Type = 2  'we want to save text/string data.
    fsT.Charset = "utf-16le"
    fsT.Open
    fsT.WriteText text
    fsT.SaveToFile filename, 2
  End If
  
  cmd = "notepad """ + filename + """"
  Shell cmd, vbNormalFocus
    
End Function

Function SetModelFromActiveDoc() 'mask for button

  Set gModel = gDoc
  gCurConf = gModel.GetActiveConfiguration.Name
    
End Function

Function EditorRun() As Boolean

  Dim isSelectedComp As Boolean
  Dim form As MainForm
  Dim haveErrors As ErrorCode
  Dim aView As View
  Dim selected As Component2
  
  isSelectedComp = False
  haveErrors = ErrorCode.ok
  gIsDrawing = IsDrawing(gDoc)
  If gIsDrawing Then
    Set gDrawing = gDoc
    Set gSheet = gDrawing.GetCurrentSheet
    If gSheet Is Nothing Or IsEmpty(gSheet.GetViews) Then
      haveErrors = ErrorCode.emptySheet
    Else
      Set aView = SelectView
      Set gModel = aView.ReferencedDocument
      If gModel Is Nothing Then
        haveErrors = ErrorCode.emptyView
      Else
        gCurConf = aView.ReferencedConfiguration
      End If
    End If
  Else
    If gDoc.GetType = swDocASSEMBLY Then
      Set selected = GetSelectedComponent
      If selected Is Nothing Then
        SetModelFromActiveDoc
      Else
        Set gModel = selected.GetModelDoc2
        gCurConf = selected.ReferencedConfiguration
        isSelectedComp = True
      End If
    Else
      SetModelFromActiveDoc
    End If
  End If
  If haveErrors = ErrorCode.ok Then
    gMainConf = gCurConf
    Set gModelExt = gModel.Extension
    gNameModel = gModel.GetPathName
    gIsAssembly = CBool(gModel.GetType = swDocASSEMBLY)
    gIsUnnamed = CBool(gNameModel = "")
    Set gModelManager = gModelExt.CustomPropertyManager("")
    If gIsDrawing Then
      Set gDrawExt = gDrawing.Extension
      Set gDrawManager = gDrawExt.CustomPropertyManager("")
    End If
  End If
  Select Case haveErrors
    Case ErrorCode.ok
      MainForm.labWarningSelected.Visible = isSelectedComp
      MainForm.Show
    Case ErrorCode.emptyView
      MsgBox ("Пустой вид. Нет ссылки на модель.")
    Case ErrorCode.emptySheet
      MsgBox ("Пустой лист. Модель не обнаружена.")
  End Select
  
End Function

Function GetSelectedComponent() As Component2

  Set GetSelectedComponent = gDoc.SelectionManager.GetSelectedObjectsComponent3(1, -1)
  
End Function

Sub SetLineStyle(object_type As swUserPreferenceIntegerValue_e, value As Integer)

  gDrawExt.SetUserPreferenceInteger object_type, swDetailingNoOptionSpecified, value
  
End Sub

Function ChangeLineStyles() 'mask for button

  SetLineStyle swLineFontVisibleEdgesStyle, swLineCONTINUOUS
  SetLineStyle swLineFontVisibleEdgesThickness, swLW_NORMAL
  
  SetLineStyle swLineFontHiddenEdgesStyle, swLineHIDDEN
  SetLineStyle swLineFontHiddenEdgesThickness, swLW_THIN
  
  SetLineStyle swLineFontSketchCurvesStyle, swLineCONTINUOUS
  SetLineStyle swLineFontSketchCurvesThickness, swLW_THIN
  
  SetLineStyle swLineFontConstructionCurvesStyle, swLinePHANTOM
  SetLineStyle swLineFontConstructionCurvesThickness, swLW_THIN
  
  SetLineStyle swLineFontCrosshatchStyle, swLineCONTINUOUS
  SetLineStyle swLineFontCrosshatchThickness, swLW_THIN
  
  SetLineStyle swLineFontTangentEdgesStyle, swLinePHANTOM
  SetLineStyle swLineFontTangentEdgesThickness, swLW_THIN
  
  SetLineStyle swLineFontCosmeticThreadStyle, swLineCONTINUOUS
  SetLineStyle swLineFontCosmeticThreadThickness, swLW_THIN
  
  SetLineStyle swLineFontHideTangentEdgeStyle, swLineHIDDEN
  SetLineStyle swLineFontHideTangentEdgeThickness, swLW_THIN
  
  SetLineStyle swLineFontExplodedLinesStyle, swLineCHAINTHICK
  SetLineStyle swLineFontExplodedLinesThickness, swLW_THICK
  
  SetLineStyle swLineFontBreakLineStyle, swLineCONTINUOUS
  SetLineStyle swLineFontBreakLineThickness, swLW_THIN
  
  SetLineStyle swLineFontSpeedPakDrawingsModelEdgesStyle, swLineCONTINUOUS
  SetLineStyle swLineFontSpeedPakDrawingsModelEdgesThickness, swLW_NORMAL
  
  SetLineStyle swLineFontAdjoiningComponentStyle, swLineCENTER
  SetLineStyle swLineFontAdjoiningComponent, swLW_THIN
  
  SetLineStyle swLineFontBendLineUpStyle, swLinePHANTOM
  SetLineStyle swLineFontBendLineUpThickness, swLW_THIN
  
  SetLineStyle swLineFontBendLineDownStyle, swLinePHANTOM
  SetLineStyle swLineFontBendLineDownThickness, swLW_THIN
  
  SetLineStyle swLineFontEnvelopeComponentStyle, swLineCONTINUOUS
  SetLineStyle swLineFontEnvelopeComponentThickness, swLW_THIN
    
End Function

Sub ReloadSheet(format As String)

  Const isFirstAngle As Boolean = True
  Dim fullFormatName As String
  Dim oldSheetProp() As Double
  Dim width As Double, height As Double
  Dim scale1 As Double, scale2 As Double
      
  fullFormatName = gConfigPath + format + ".SLDDRT"
  If gFSO.FileExists(fullFormatName) Then
    oldSheetProp = gSheet.GetProperties
    scale1 = oldSheetProp(2)
    scale2 = oldSheetProp(3)
    gSheet.GetSize width, height
    
    'Set the sheet format to None
    gDrawing.SetupSheet5 gSheet.GetName, swDwgPapersUserDefined, _
                         swDwgTemplateNone, scale1, scale2, _
                         isFirstAngle, "", width, height, _
                         gSheet.CustomPropertyView, True
                         
    'Reload the sheet format from the specified location
    gDrawing.SetupSheet5 gSheet.GetName, swDwgPapersUserDefined, _
                         swDwgTemplateCustom, scale1, scale2, _
                         isFirstAngle, fullFormatName, width, height, _
                         gSheet.CustomPropertyView, True
                         
    'gDoc.ViewZoomtofit2
    ZoomToSheet
  Else
    MsgBox ("Файл " + fullFormatName + " не найден.")
  End If
    
End Sub

Function IntOrNul(str As String) As Long

  IntOrNul = 0
  If IsNumeric(str) Then
    IntOrNul = CInt(str)
  End If
    
End Function

Function ReadMaterialNames(filename As String) As Dictionary

  Dim Result As Dictionary
  Dim fullFilename As String
  Dim I As Variant
  Dim Matches As MatchCollection
  Dim FStream As TextStream
  Dim TextAll As String
  
  Set Result = New Dictionary
  fullFilename = gConfigPath + filename
  If gFSO.FileExists(fullFilename) Then
    Set FStream = gFSO.OpenTextFile(fullFilename, ForReading, False, TristateTrue)
    TextAll = FStream.ReadAll
    FStream.Close
    If gRegexMaterial.Test(TextAll) Then
      Set Matches = gRegexMaterial.Execute(TextAll)
      For Each I In Matches
        Result.Add I.SubMatches(0), ""
      Next
    End If
  End If
  Set ReadMaterialNames = Result
    
End Function

Function FirstItem(values As Variant) As String

  FirstItem = ""
  If Not IsEmpty(values) Then
    FirstItem = values(0)
  End If
    
End Function

Function FindView() As View

  Dim propView As String
  Dim firstView As View
  
  propView = gSheet.CustomPropertyView
  Set firstView = gDrawing.GetFirstView.GetNextView
  Set FindView = firstView
  Do While FindView.GetName2 <> propView
    Set FindView = FindView.GetNextView
    If FindView Is Nothing Then
      Set FindView = firstView
      Exit Do
    End If
  Loop
    
End Function

Function SelectView() As View

  Dim selected As Object
  
  Set selected = gDrawing.SelectionManager.GetSelectedObject5(1)
  If selected Is Nothing Then
    Set SelectView = FindView()
  ElseIf Not TypeOf selected Is View Then
    Set SelectView = FindView()
  Else
    Set SelectView = selected
  End If
    
End Function

Function IndexInArray(valueToFind As Variant, arr As Variant) As Integer

  Dim I As Integer
  
  IndexInArray = -1
  For I = LBound(arr) To UBound(arr)
    If arr(I) = valueToFind Then
      IndexInArray = I
      Exit Function
    End If
  Next
    
End Function

Sub SaveAsMy(NewName As String, OldName As String)

  Dim error As Long, warning As Long

  If gFSO.FileExists(NewName) Then
    If MsgBox("Файл с таким именем существует. Заменить?", vbOKCancel) = vbCancel Then
      Exit Sub
    End If
  End If
  If gDoc.Extension.SaveAs(NewName, swSaveAsCurrentVersion, _
                           swSaveAsOptions_AvoidRebuildOnSave, _
                           Nothing, error, warning) Then
      If OldName <> "" Then
        Kill (OldName)
      End If
  Else
    MsgBox ("Не удалось сохранить файл")
  End If
    
End Sub

Sub TryRenameDraft(sName As String)

  Dim NewName As String, OldName As String

  If sName = "" Then
    Exit Sub
  End If
  OldName = gDoc.GetPathName
  If OldName = "" Then
    NewName = gFSO.BuildPath(gFSO.GetParentFolderName(gModel.GetPathName), sName + ".SLDDRW")
    SaveAsMy NewName, OldName
  Else
    NewName = gFSO.BuildPath(gFSO.GetParentFolderName(OldName), sName + ".SLDDRW")
    If OldName <> NewName Then
      SaveAsMy NewName, OldName
    End If
  End If
    
End Sub

'Only for drawings!
Function ZoomToSheet()  'mask for button

  Dim width As Double
  Dim height As Double

  gDrawing.GetCurrentSheet.GetSize width, height
  gDoc.ViewZoomTo2 0, 0, 0, width, height, 0
    
End Function

Function GetEquationThickness(Conf As String, toAll As Boolean, nameModel As String) As String

  Const temp As String = "__temp__"
  Dim mgr As CustomPropertyManager
  Dim thName_ As Variant
  Dim thName As String
  Dim resolvedValue As String
  Dim rawValue As String
  Dim variantThicknessName(1) As String
  
  GetEquationThickness = ""
  Set mgr = gModel.Extension.CustomPropertyManager(Conf)
  variantThicknessName(0) = "Толщина"
  variantThicknessName(1) = "Thickness"
  'variantThicknessName(2) = "Grubos'c'"
  For Each thName_ In variantThicknessName
    thName = thName_
    SetProp mgr, temp, Equal(thName, toAll, Conf, gNameModel)
    mgr.Get5 temp, False, rawValue, resolvedValue, False
    If IsNumeric(StrConv(resolvedValue, vbUnicode)) Then  'IsNumeric make error with raw 'resolvedValue' without StrConv
      GetEquationThickness = rawValue
      Exit For
    End If
  Next
  mgr.Delete2 temp
    
End Function

Function GetBaseDesignation(Designation As String) As String

  Dim lastFullstopPosition As Integer
  Dim firstHyphenPosition As Integer
  
  GetBaseDesignation = Designation
  lastFullstopPosition = InStrRev(Designation, ".")
  If lastFullstopPosition > 0 Then
    firstHyphenPosition = InStr(lastFullstopPosition, Designation, "-")
    If firstHyphenPosition > 0 Then
      GetBaseDesignation = Left(Designation, firstHyphenPosition - 1)
    End If
  End If
  
End Function

Function ExitApp() 'hide

  Unload MainForm
  End

End Function

Function GetIsFastener() As Boolean

  Dim value As String
  Dim rawValue As String
  Dim wasResolved As Boolean

  GetIsFastener = False
  If gModelManager.Get5(pIsFastener, False, rawValue, value, wasResolved) <> swCustomInfoGetResult_NotPresent Then
    GetIsFastener = (value = IsFastenerTrue)
  End If

End Function

Function SetIsFastener() 'hide

  gModelManager.Delete2 pIsFastener
  gModelManager.Add3 pIsFastener, swCustomInfoNumber, IsFastenerTrue, swCustomPropertyDeleteAndAdd

End Function

Sub RewriteNameAndSign(source As String, Conf As String)

  Dim Designation As String
  Dim Name As String
  Dim Code As String
  Dim I As Integer
  Dim IsCodeFound As Boolean
  
  Designation = ""
  Name = ""
  SplitNameAndSign source, Conf, Designation, Name, Code
  MainForm.SignBox.text = Designation
  MainForm.NameBox.text = Name
  If gIsDrawing Then
    IsCodeFound = False
    I = 0
    While (I < MainForm.MiniSignBox.ListCount) And (Not IsCodeFound)
      IsCodeFound = (StrComp(MainForm.MiniSignBox.List(I), Code, vbTextCompare) = 0)
      If IsCodeFound Then
        MainForm.MiniSignBox.ListIndex = I
      End If
      I = I + 1
    Wend
  End If
  
End Sub

' Без точек "." в наименовании
Sub SplitNameAndSign(Line As String, Conf As String, ByRef Designation As String, _
                     ByRef Name As String, ByRef Code As String)
                     
  Const flat As String = "SM-FLAT-PATTERN"
  Dim regexAsm As RegExp
  Dim regexPrt As RegExp
  Dim Matches As Object
  Dim z As Variant
  
  Designation = Line
  Name = Line
  Code = ""
  
  Set regexAsm = New RegExp
  regexAsm.Pattern = "(.*\..*[0-9] *)(" + gCodeRegexPattern + ") ([^.]+)"
  regexAsm.IgnoreCase = True
  regexAsm.Global = True
  
  Set regexPrt = New RegExp
  regexPrt.Pattern = "(.*\.[^ ]+) ([^.]+)"
  regexPrt.IgnoreCase = True
  regexPrt.Global = True
  
  If regexAsm.Test(Line) Then
    Set Matches = regexAsm.Execute(Line)
    Designation = Trim(Matches(0).SubMatches(0))
    Code = Matches(0).SubMatches(1)
    Name = Trim(Matches(0).SubMatches(2))
  ElseIf regexPrt.Test(Line) Then
    Set Matches = regexPrt.Execute(Line)
    Designation = Trim(Matches(0).SubMatches(0))
    Name = Trim(Matches(0).SubMatches(1))
  End If
  
  If Conf Like "*" & flat Then
    Conf = Left(Conf, Len(Conf) - Len(flat))
  End If
  Select Case Conf
    Case "00", "По умолчанию"
      'pass
    Case Else
      MainForm.SignChk.value = False ' running event
      Designation = Designation & "-" & Conf
  End Select
  
End Sub

' Устанавливает значения gItems из свойств, игнорируя существующие
Sub ReadProp(manager As CustomPropertyManager, Conf As String, props() As String)

  Dim I As Variant
  Dim prop As String
  Dim Item As DataItem
  Dim raw As String
  Dim val As String
  
  If Not gItems.Exists(Conf) Then
    gItems.Add Conf, New Dictionary
  End If
  
  For Each I In props
    prop = I
    
    If Not gItems(Conf).Exists(prop) Then
      gItems(Conf).Add prop, New DataItem
    End If
    
    Set Item = gItems(Conf)(prop)
    raw = ""
    val = ""
    manager.Get4 prop, False, raw, val
    Item.rawValue = raw
    Item.value = val
    
    If Conf <> commonSpace Then
      Item.fromAll = (Item.rawValue = "") And (gItems(commonSpace)(prop).rawValue <> "")
    Else
      Item.fromAll = True
    End If
    
    If prop = pMaterial Then
      Item.newValue = Item.value
    Else
      Item.newValue = Item.rawValue
    End If
  Next
    
End Sub

Sub SetBoxValue2(Chk As CheckBox, prop As String, Conf As String)

  Dim Item As DataItem
  
  Set Item = gItems(Conf)(prop)

  If Not Chk Is Nothing Then
    If Chk.value <> Item.fromAll Then
      Chk.value = Item.fromAll
    Else
      ChangeChecked prop
    End If
  Else
    ChangeChecked prop
  End If
    
End Sub

Sub ReloadForm(Conf As String)

  readOldAfterChecked = False

  If gIsDrawing Then
    SetBoxValue2 Nothing, pShortDrawingType, commonSpace
    SetBoxValue2 Nothing, pLongDrawingType, commonSpace
    SetBoxValue2 Nothing, pOrganization, commonSpace
    SetBoxValue2 Nothing, pDrafter, commonSpace
    SetBoxValue2 Nothing, pChecking, commonSpace
  End If
  
  SetBoxValue2 MainForm.DevelChk, pDesigner, Conf
  SetBoxValue2 MainForm.SignChk, pDesignation, Conf
  
  SetBoxValue2 MainForm.NameChk, pName, Conf
  ChangeChecked pNameEN
  ChangeChecked pNamePL
  ChangeChecked pNameUA
  
  SetBoxValue2 MainForm.FormatChk, pFormat, Conf
  SetBoxValue2 MainForm.NoteChk, pNote, Conf
  SetBoxValue2 MainForm.MassChk, pMass, Conf
  
  If Not gIsAssembly Then
    SetBoxValue2 MainForm.BlankChk, pBlank, Conf
    SetBoxValue2 MainForm.SizeChk, pSize, Conf
    SetBoxValue2 Nothing, pMaterial, Conf
    SetBoxValue2 MainForm.lenChk, pLen, Conf
    SetBoxValue2 MainForm.widChk, pWid, Conf
  End If
  readOldAfterChecked = True
    
End Sub

Function ExistsInCombo(Box As ComboBox, value As String)

  Dim I As Variant

  ExistsInCombo = False
  If Box.ListCount > 0 Then
    For Each I In Box.List
      If I = value Then
        ExistsInCombo = True
        Exit For
      End If
    Next
  End If
    
End Function

Sub FromAllChecked(Chk As CheckBox, Box As Object, prop As String, Conf As String, _
                   fromAll As Boolean, SetFirstItem As Boolean)
                   
  Dim cmb As ComboBox
  Dim value As String
  
  If readOldAfterChecked Then
    ReadBox Box, Chk, Conf, prop, False
  End If
  If prop = pSize Then
    ChangeSizeEqual (Conf)
  ElseIf prop = pMass Then
    ChangeMassEqual (Conf)
  End If
  
  If fromAll Then
    value = gItems(commonSpace)(prop).newValue
  Else
    value = gItems(Conf)(prop).newValue
  End If
  If SetFirstItem And value = "" And TypeOf Box Is ComboBox Then
    Set cmb = Box
    If cmb.ListCount > 0 Then
      value = cmb.List(0)
    End If
  End If
  
  If Box.Enabled Then
    If TypeOf Box Is ComboBox Then
      Set cmb = Box
      If cmb.Style = fmStyleDropDownList Then
        If ExistsInCombo(cmb, value) Then
          SetComboInExistValue Box, value
        ElseIf cmb.ListCount > 0 Then
          cmb.ListIndex = 0
        End If
      Else
        cmb.text = value
      End If
    Else
      Box.text = value
    End If
  End If
  
End Sub

Sub SetComboInExistValue(ByRef Box As Object, value As String)

  On Error Resume Next  ''''ПОДАВЛЕНИЕ ОШИБКИ для Гордиенко
  Box.text = value
    
End Sub

' Устанавливает значения gItems из формы
' conf - конфигурация ИЛИ элемент списка вырезов
Sub ReadBox(Box As Object, Chk As CheckBox, Conf As String, prop As String, forward As Boolean)

  Dim TargetConf As String
  
  If Not gItems.Exists(Conf) Then
    gItems.Add Conf, New Dictionary
  End If
  If Not gItems(Conf).Exists(prop) Then
    gItems(Conf).Add prop, New DataItem
  End If
  
  If Chk Is Nothing And Conf = commonSpace Then
    gItems(commonSpace)(prop).fromAll = True
    gItems(commonSpace)(prop).newValue = Box.text
  ElseIf prop = pMaterial Then
    gItems(Conf)(prop).fromAll = False
    gItems(Conf)(prop).newValue = Box.text   'уравнение MaterialEqual устанавливается в SetProp2
  Else
    gItems(Conf)(prop).fromAll = Chk.value
    If forward Then
      If Chk.value Then
        TargetConf = commonSpace
      Else
        TargetConf = Conf
      End If
    Else
      If Chk.value Then
        TargetConf = Conf
      Else
        TargetConf = commonSpace
      End If
    End If
    gItems(TargetConf)(prop).newValue = Box.text
  End If
  
End Sub

Sub SetPropToAll(Box As Object, Chk As CheckBox, Property As String)

  Dim I As Variant
  Dim Conf As String
  Dim ConfManager As CustomPropertyManager
  
  For Each I In gModelConfNames
    Conf = I
    Set ConfManager = gModelExt.CustomPropertyManager(Conf)
    ConfManager.Delete2 Property
  Next
  ReadBox Box, Nothing, commonSpace, Property, True
    
End Sub

Sub ReadForm(Conf As String)

  ReadBox MainForm.NameBox, MainForm.NameChk, Conf, pName, True
  ReadBox MainForm.NameBoxEN, MainForm.NameChk, Conf, pNameEN, True
  ReadBox MainForm.NameBoxPL, MainForm.NameChk, Conf, pNamePL, True
  ReadBox MainForm.NameBoxUA, MainForm.NameChk, Conf, pNameUA, True
  
  ReadBox MainForm.DevelBox, MainForm.DevelChk, Conf, pDesigner, True
  ReadBox MainForm.SignBox, MainForm.SignChk, Conf, pDesignation, True
  ReadBox MainForm.FormatBox, MainForm.FormatChk, Conf, pFormat, True
  ReadBox MainForm.NoteBox, MainForm.NoteChk, Conf, pNote, True
  ReadBox MainForm.MassBox, MainForm.MassChk, Conf, pMass, True
  
  If gIsDrawing Then
    ReadBox MainForm.MiniSignBox, Nothing, commonSpace, pShortDrawingType, True
    ReadBox MainForm.CodeBox, Nothing, commonSpace, pLongDrawingType, True
    ReadBox MainForm.OrgBox, Nothing, commonSpace, pOrganization, True
    ReadBox MainForm.DraftBox, Nothing, commonSpace, pDrafter, True
    ReadBox MainForm.CheckingBox, Nothing, commonSpace, pChecking, True
  End If
  
  If Not gIsAssembly Then
    ReadBox MainForm.BlankBox, MainForm.BlankChk, Conf, pBlank, True
    ReadBox MainForm.SizeBox, MainForm.SizeChk, Conf, pSize, True
    ReadBox MainForm.MaterialBox, Nothing, Conf, pMaterial, True
    ReadBox MainForm.lenBox, MainForm.lenChk, Conf, pLen, True
    ReadBox MainForm.widBox, MainForm.widChk, Conf, pWid, True
  End If
  
End Sub

Sub ChangeChecked(prop As String)

  Select Case prop
    Case pDesignation
      FromAllChecked MainForm.SignChk, MainForm.SignBox, pDesignation, gCurConf, MainForm.SignChk.value, False
    Case pName
      FromAllChecked MainForm.NameChk, MainForm.NameBox, pName, gCurConf, MainForm.NameChk.value, False
    Case pNameEN
      FromAllChecked MainForm.NameChk, MainForm.NameBoxEN, pNameEN, gCurConf, MainForm.NameChk.value, False
    Case pNamePL
      FromAllChecked MainForm.NameChk, MainForm.NameBoxPL, pNamePL, gCurConf, MainForm.NameChk.value, False
    Case pNameUA
      FromAllChecked MainForm.NameChk, MainForm.NameBoxUA, pNameUA, gCurConf, MainForm.NameChk.value, False
    Case pBlank
      FromAllChecked MainForm.BlankChk, MainForm.BlankBox, pBlank, gCurConf, MainForm.BlankChk.value, False
    Case pFormat
      FromAllChecked MainForm.FormatChk, MainForm.FormatBox, pFormat, gCurConf, MainForm.FormatChk.value, False
    Case pNote
      FromAllChecked MainForm.NoteChk, MainForm.NoteBox, pNote, gCurConf, MainForm.NoteChk.value, False
    Case pDesigner
      FromAllChecked MainForm.DevelChk, MainForm.DevelBox, pDesigner, gCurConf, MainForm.DevelChk.value, True
    Case pSize
      FromAllChecked MainForm.SizeChk, MainForm.SizeBox, pSize, gCurConf, MainForm.SizeChk.value, False
    Case pMass
      FromAllChecked MainForm.MassChk, MainForm.MassBox, pMass, gCurConf, MainForm.MassChk.value, True
    Case pMaterial
      FromAllChecked Nothing, MainForm.MaterialBox, pMaterial, gCurConf, False, False
    Case pOrganization
      FromAllChecked Nothing, MainForm.OrgBox, pOrganization, gCurConf, True, True
    Case pDrafter
      FromAllChecked Nothing, MainForm.DraftBox, pDrafter, gCurConf, True, True
    Case pChecking
      FromAllChecked Nothing, MainForm.CheckingBox, pChecking, gCurConf, True, True
    Case pShortDrawingType
      FromAllChecked Nothing, MainForm.MiniSignBox, pShortDrawingType, gCurConf, True, False
    Case pLongDrawingType
      FromAllChecked Nothing, MainForm.CodeBox, pLongDrawingType, gCurConf, True, False
    Case pLen
      FromAllChecked MainForm.lenChk, MainForm.lenBox, pLen, gCurConf, MainForm.lenChk.value, False
    Case pWid
      FromAllChecked MainForm.widChk, MainForm.widBox, pWid, gCurConf, MainForm.widChk.value, False
  End Select
   
End Sub

Sub TrySetPropToAll(Box As Object, Chk As CheckBox, Property As String)

  If isShiftPressed And Not Chk.value Then
    isShiftPressed = False
    
    readOldAfterChecked = False
    Chk.value = True
    readOldAfterChecked = True
    
    If MsgBox("Связать все конфигурации со значением?", vbYesNo) = vbYes Then
      SetPropToAll Box, Chk, Property
    End If
  Else
    ChangeChecked Property
  End If
  
End Sub

Function CreateBaseDesignation() 'hide

  Dim mainDesignation As String
  Dim resolvedValue As String
  Dim rawValue As String
  Dim wasResolved As Boolean
  
  If gModel.Extension.CustomPropertyManager(gMainConf).Get5(pDesignation, False, rawValue, resolvedValue, wasResolved) = swCustomInfoGetResult_NotPresent Then
    gModel.Extension.CustomPropertyManager("").Get5 pDesignation, False, rawValue, resolvedValue, wasResolved
  End If
  
  gBaseDesignation = GetBaseDesignation(resolvedValue)
  
End Function

Function Execute() 'hide

  ReadForm gCurConf
  
  gModel.SetReadOnlyState False  'must be first!
  
  WriteModelProperties
  
  If Not gIsAssembly Then
    If MainForm.IsFastenerChk.value Then
      SetIsFastener
    End If
  End If
  
  If gIsDrawing Then
    CreateBaseDesignation
    WriteDrawingProperties
    SetSpeedformat
    ChangeLineStyles
  End If
  
  gDoc.ForceRebuild3 True
  
  If gIsDrawing Then
    TryRenameDraft MainForm.DrawNameBox.text
  End If
    
End Function

Sub ChangeMassEqual(Conf As String)

  If Not gIsUnnamed Then
    MainForm.MassBox.List(0) = Equal("SW-Mass", gItems(Conf)(pMass).fromAll, Conf, gNameModel)
  End If
    
End Sub

Sub ChangeSizeEqual(Conf As String)

  If Not gIsUnnamed And Not gIsAssembly Then
    MainForm.SizeBox.List(0) = GetEquationThickness(Conf, gItems(Conf)(pSize).fromAll, gNameModel)
  End If
  
End Sub

Function MaterialEqual(Conf As String) As String

  If Not gIsUnnamed And Not gIsAssembly Then
    MaterialEqual = Equal(pTrueMaterial, False, Conf, gNameModel)
  Else
    MaterialEqual = ""
  End If
  
End Function

Sub SetMaterial(Conf As String)

  Dim new_material As String
  
  new_material = gItems(Conf)(pMaterial).newValue
  If new_material <> sEmpty And new_material <> "" Then
    gModel.SetMaterialPropertyName2 Conf, materialdb, new_material  'it's method of PartDoc
  End If
    
End Sub

Function SetSpeedformat() 'hide

  If MainForm.RealFormatBox.text <> MainForm.RealFormatBox.List(0) Then
    ReloadSheet MainForm.RealFormatBox.text
  End If
    
End Function

Function OutputTypeAndName() 'hide

  MainForm.ModelNameBox.text = gFSO.GetBaseName(gNameModel)
  If gIsDrawing Then
    MainForm.DrawNameBox.Enabled = True
    MainForm.DrawNameLab.Enabled = True
    MainForm.DrawNameBox.text = gFSO.GetBaseName(gDoc.GetPathName)
  End If
  If gIsAssembly Then
    MainForm.Controls("ModelNameLab").Caption = "Файл сборки"
  Else
    MainForm.Controls("ModelNameLab").Caption = "Файл детали"
  End If
    
End Function

Function CreateCodeRegexPattern() As String

  Dim I As Integer
  Dim Codes() As String
  
  If UserDrawingTypes.Count > 0 Then
    ReDim Codes(UserDrawingTypes.Count - 1)
    For I = 0 To UserDrawingTypes.Count - 1
      Codes(I) = Replace(UserDrawingTypes.Keys(I), ".", "\.")
    Next
    CreateCodeRegexPattern = Join(Codes, "|")
  Else
    CreateCodeRegexPattern = "СБ|МЧ|УЧ|РСБ"
  End If
  'Debug.Print CreateCodeRegexPattern
  
End Function

Function GetConfNames() 'hide

  gModelConfNames = gModel.GetConfigurationNames
  QuickSort gModelConfNames, LBound(gModelConfNames), UBound(gModelConfNames) 'configurations list is not sorted
    
End Function

Function InitRealFormatBox() 'mask for button

  Dim filename As String
  Dim Names() As String
  Dim I As Long
  
  MainForm.RealFormatBox.AddItem ("<данная>")
  MainForm.RealFormatBox.text = MainForm.RealFormatBox.List(0)
  I = -1
  filename = Dir(gConfigPath & "*.SLDDRT")
  While filename <> ""
    I = I + 1
    ReDim Preserve Names(0 To I)
    Names(I) = gFSO.GetBaseName(filename)
    filename = Dir()
  Wend
  Names = SortSpeedFormats(Names)
  While I >= 0
    MainForm.RealFormatBox.AddItem Names(I)
    I = I - 1
  Wend
    
End Function

Function SortSpeedFormats(Names() As String) As String()

  Dim majorNames() As String
  Dim minorNames() As String
  Dim name_ As Variant
  Dim Name As String
  Dim n As Integer
  Dim j As Integer
  Dim I As Integer
  
  n = -1
  j = -1
  If Not IsArrayEmpty(Names) Then
    For Each name_ In Names
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
      Names(LBound(Names) + I) = minorNames(I)
    Next
    For I = j To 0 Step -1
      Names(UBound(Names) - j + I) = majorNames(I)
    Next
  End If
  SortSpeedFormats = Names
    
End Function

Sub SetModelProp(Conf As String, prop As String, Item As DataItem)

  Dim ConfManager As CustomPropertyManager
  
  Set ConfManager = gModelExt.CustomPropertyManager(Conf)
  If Conf <> commonSpace And Item.fromAll Then
    ConfManager.Delete (prop)
  Else
    SetProp2 ConfManager, prop, Item, Conf
  End If
    
End Sub

Function SetProp2(manager As CustomPropertyManager, prop As String, Item As DataItem, _
                  Optional Conf As String = commonSpace) As Boolean
                  
  Dim Result As Boolean
  
  Result = False
  If prop = pMaterial Then
    'If item.newValue <> sEmpty Then
      Result = SetProp(manager, prop, MaterialEqual(Conf))
    'Else
    '  gModelExt.CustomPropertyManager(conf).Delete2 pMaterial
    '  gModelManager.Delete2 pMaterial
    'End If
  Else
    Result = SetProp(manager, prop, Item.newValue)
  End If
  SetProp2 = Result
    
End Function

Function WriteModelProperties() 'hide

  Dim I As Variant
  Dim j As Variant
  Dim Conf As String
  Dim prop As String
  Dim Item As DataItem
  
  For Each I In gItems.Keys
    Conf = I
    For Each j In modelProps
      prop = j
      Set Item = gItems(Conf)(prop)
      
      Select Case prop
        Case pBlank, pSize, pLen, pWid
          If Not gIsAssembly Then
            SetModelProp Conf, prop, Item
          End If
        Case pMaterial
          If Not gIsAssembly Then
            If Not gIsUnnamed Then
              SetModelProp Conf, prop, Item
            End If
            SetMaterial Conf
          End If
        Case Else
          SetModelProp Conf, prop, Item
      End Select
      
    Next
  Next
    
End Function

'TODO: убрать чертежные свойства из gItems, читать их прямо из формы
Function WriteDrawingProperties() 'hide

  Dim toAll As Boolean: toAll = True
  Dim Item As Dictionary: Set Item = gItems(commonSpace)
  Dim DrawingCode As String

  DrawingCode = MainForm.MiniSignBox.text
  'см. массив drawProps
  SetProp gDrawManager, pShortDrawingType, DrawingCode
  SetProp2 gDrawManager, pOrganization, Item(pOrganization)
  SetProp2 gDrawManager, pDrafter, Item(pDrafter)
  SetProp2 gDrawManager, pChecking, Item(pChecking)  'before: userChecking(0)
  SetProp gDrawManager, pApprover, userApprover(0)
  SetProp gDrawManager, pNormControl, userNormControl(0)
  SetProp gDrawManager, pTechControl, userTechControl(0)
  SetProp gDrawManager, pLongDrawingType, IIf(DrawingCode = "", "", MainForm.CodeBox.text)
  SetProp gDrawManager, pBaseDesignation, gBaseDesignation
    
End Function

Sub SetValueInBox(Box As ComboBox, Index As Integer)

  If 0 <= Index And Index < Box.ListCount Then
    Box.text = Box.List(Index)
  End If
    
End Sub

Sub ExitByKey(KeyCode As MSForms.ReturnInteger, Shift As Integer)

  If Shift = 1 And KeyCode = vbKeyReturn Then
    Execute
    ExitApp
  End If
    
End Sub

Sub SetShiftStatus(Shift As Integer)
  
  isShiftPressed = Shift And 1
  
End Sub

Function SetPartCaptionIfEmptyDrawingCode() As Boolean

  SetPartCaptionIfEmptyDrawingCode = (MainForm.MiniSignBox.text = "")
  If SetPartCaptionIfEmptyDrawingCode Then
    MainForm.CodeBox.AddItem "Деталь"
    MainForm.CodeBox.text = MainForm.CodeBox.List(0)
    MainForm.CodeBox.Enabled = False
  End If

End Function
