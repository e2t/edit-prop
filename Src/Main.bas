Attribute VB_Name = "Main"
Option Explicit

Public Const MacroName = "EditProp"
Public Const MacroSection = "Main" 'registry
Public Const IniFileName = "Settings.ini"
Public Const MainSection = "Main" 'file.ini
Const KeyApplyAndExit = "ApplyAndExit"

'Свойства модели, записаны в массиве modelProps
Public Const pDesignation = "Обозначение"
Public Const pMaterial = "Материал"
Public Const pName = "Наименование"
Public Const pNameEN = "Наименование EN"
Public Const pNamePL = "Наименование PL"
Public Const pNameUA = "Наименование UA"
Public Const pBlank = "Заготовка"
Public Const pSize = "Типоразмер"
Public Const pNote = "Примечание"
Public Const pDesigner = "Разработал"
Public Const pFormat = "Формат"
Public Const pMass = "Масса"
Public Const pLen = "Длина"
Public Const pWid = "Ширина"
'Специальное свойство для получения материала
Public Const pTrueMaterial = "SW-Material"
'Специальное свойство деталей
Public Const pIsFastener = "IsFastener"
Public Const IsFastenerTrue = "1"
'Свойства чертежа, записаны в массиве drawProps
Public Const pDrafter = "Начертил"
Public Const pShortDrawingType = "Пометка"
Public Const pLongDrawingType = "Тип документа"
Public Const pOrganization = "Организация"
Public Const pChecking = "Проверил"
Public Const pApprover = "Утвердил"
Public Const pTechControl = "Техконтроль"
Public Const pNormControl = "Нормоконтроль"
Public Const pBaseDesignation = "Базовое обозначение"
Public Const MaterialDB = "Материалы"
Public Const CommonSpace = ""
Public Const Separator = ";"
Public Const Separator2 = "="
Public Const Separator3 = ","
Public Const SettingsFile = "Настройки.txt"
Public Const sEmpty = " "
Public Const CurrentChoice = "[текущ.]"
Public Const MaxNamingLen = 60
Public Const TagPaperSize = "PaperSize"

Enum ErrorCode
  Ok = 0
  EmptyView = 1
  EmptySheet = 2
End Enum

Public swApp As Object
Public gFSO As FileSystemObject

Public gDoc As ModelDoc2
Public gModel As ModelDoc2
Public gModelExt As ModelDocExtension
Public gCurConf As String 'выбранная в списке конфигурация
Public gIsAssembly As Boolean
Public gIsDrawing As Boolean
Public gModelManager As CustomPropertyManager
Public gDrawManager As CustomPropertyManager
Public UserDrawingTypes As Dictionary
Public ModelProps(13) As String
Public DrawProps(8) As String
Public PaperSizes As Dictionary
Public gItems As Dictionary
Public ReadOldAfterChecked As Boolean
Public gIsShiftPressed As Boolean
Public gIniFilePath As String

Dim gConfigPath As String
Dim gDrawExt As ModelDocExtension
Dim gSheet As Sheet
Dim gModelConfNames() As String
Dim gDrawing As DrawingDoc
Dim gSheetScale1 As Double
Dim gSheetScale2 As Double
Dim gIsFirstAngle As Double
Dim gMainConf As String 'основная конфигурация на чертеже
Dim gBaseDesignation As String
Dim gChangeNumber As Long
Dim gIsUnnamed As Boolean
Dim gNameModel As String
Dim UserName() As String
Dim UserBlank() As String
Dim UserSize() As String
Dim UserFormat() As String
Dim UserNote() As String
Dim UserMass() As String
Dim UserOrganization() As String
Dim UserDesigner() As String
Dim UserDrafter() As String
Dim UserChecking() As String
Dim UserApprover() As String
Dim UserTechControl() As String
Dim UserNormControl() As String
Dim UserLen() As String
Dim UserWid() As String
Dim UserMaterials() As String
Dim gCodeRegexPattern As String
Dim gRegexMaterial As RegExp
Dim gIsApplyAndExit As Boolean

Sub Main()

  Init
  Set gDoc = swApp.ActiveDoc
  If gDoc Is Nothing Then
    MsgBox "Нет открытых документов."
  Else
    EditorRun
  End If
    
End Sub

Function Init() As Boolean

  Set swApp = Application.SldWorks
  Set gFSO = New FileSystemObject
  
  ReDim UserName(0)
  UserName(0) = ""
  
  Set UserDrawingTypes = New Dictionary
  
  ReDim UserBlank(0)
  UserBlank(0) = ""
  ReDim UserSize(0)
  UserSize(0) = ""
  ReDim UserFormat(0)
  UserFormat(0) = ""
  ReDim UserNote(0)
  UserNote(0) = ""
  ReDim UserMass(0)
  UserMass(0) = ""
  ReDim UserOrganization(0)
  UserOrganization(0) = ""
  ReDim UserDesigner(0)
  UserDesigner(0) = ""
  ReDim UserDrafter(0)
  UserDrafter(0) = ""
  ReDim UserChecking(0)
  UserChecking(0) = ""
  ReDim UserApprover(0)
  UserApprover(0) = ""
  ReDim UserTechControl(0)
  UserTechControl(0) = ""
  ReDim UserNormControl(0)
  UserNormControl(0) = ""
  ReDim UserWid(0)
  UserWid(0) = ""
  ReDim UserLen(0)
  UserLen(0) = ""
  ReDim UserMaterials(0)
  UserMaterials(0) = ""
  ReDim UserPreExclude(0)
  UserPreExclude(0) = ""
  
  ModelProps(0) = pDesignation
  ModelProps(1) = pMaterial
  ModelProps(2) = pName
  ModelProps(3) = pBlank
  ModelProps(4) = pSize
  ModelProps(5) = pNote
  ModelProps(6) = pDesigner
  ModelProps(7) = pFormat
  ModelProps(8) = pMass
  ModelProps(9) = pLen
  ModelProps(10) = pWid
  ModelProps(11) = pNameEN
  ModelProps(12) = pNamePL
  ModelProps(13) = pNameUA
  
  DrawProps(0) = pDrafter
  DrawProps(1) = pShortDrawingType
  DrawProps(2) = pLongDrawingType
  DrawProps(3) = pOrganization
  DrawProps(4) = pChecking
  DrawProps(5) = pApprover
  DrawProps(6) = pTechControl
  DrawProps(7) = pNormControl
  DrawProps(8) = pBaseDesignation
  
  gConfigPath = swApp.GetCurrentMacroPathFolder() + "\config\"
  ReadSettings  'only after "gConfigPath = ..."
  
  Set gRegexMaterial = New RegExp
  gRegexMaterial.Global = True
  gRegexMaterial.MultiLine = True
  gRegexMaterial.IgnoreCase = True
  gRegexMaterial.Pattern = "material name=""([^""]+)"""
  
  Set PaperSizes = New Dictionary
  AppendPaperSize "A4", "A4", 0.21, 0.297
  AppendPaperSize "A3 гориз", "A3", 0.42, 0.297
  AppendPaperSize "A3 верт", "A3", 0.297, 0.42
  AppendPaperSize "A2 гориз", "A2", 0.594, 0.42
  AppendPaperSize "A2 верт", "A2", 0.42, 0.594
  AppendPaperSize "A1 гориз", "A1", 0.841, 0.594
  AppendPaperSize "A1 верт", "A1", 0.594, 0.841
  AppendPaperSize "A0 гориз", "A0", 1.189, 0.841
  AppendPaperSize "A0 верт", "A0", 0.841, 1.189
  
  AppendPaperSize "A4x3", "A4x3", 0.63, 0.297
  AppendPaperSize "A4x4", "A4x4", 0.841, 0.297
  AppendPaperSize "A4x5", "A4x5", 1.051, 0.297
  AppendPaperSize "A4x6", "A4x6", 1.261, 0.297
  
  AppendPaperSize "A3x3", "A3x3", 0.891, 0.42
  AppendPaperSize "A3x4", "A3x4", 1.189, 0.42
  AppendPaperSize "A3x5", "A3x5", 1.486, 0.42
  AppendPaperSize "A3x6", "A3x6", 1.783, 0.42
  
  AppendPaperSize "A2x3", "A2x3", 1.261, 0.594
  AppendPaperSize "A2x4", "A2x4", 1.682, 0.594
  AppendPaperSize "A2x5", "A2x5", 2.102, 0.594
  AppendPaperSize "A2x6", "A2x6", 2.52, 0.594
  
  AppendPaperSize "A1x3", "A1x3", 1.783, 0.841
  AppendPaperSize "A1x4", "A1x4", 2.378, 0.841
  AppendPaperSize "A1x5", "A1x5", 2.973, 0.841
  AppendPaperSize "A1x6", "A1x6", 3.568, 0.841
  
  AppendPaperSize "A0x3", "A0x3", 2.523, 1.189
  AppendPaperSize "A0x4", "A0x4", 3.36, 1.189
  AppendPaperSize "A0x5", "A0x5", 4.2, 1.189
  AppendPaperSize "A0x6", "A0x6", 5.04, 1.189
  
  AppendPaperSize "ANSI A гориз", "A", 0.28, 0.216
  AppendPaperSize "ANSI A верт", "A", 0.216, 0.28
  AppendPaperSize "ANSI B гориз", "B", 0.432, 0.279
  AppendPaperSize "ANSI B верт", "B", 0.279, 0.432
  AppendPaperSize "ANSI C гориз", "C", 0.559, 0.432
  AppendPaperSize "ANSI C верт", "C", 0.432, 0.559
  AppendPaperSize "ANSI D гориз", "D", 0.864, 0.559
  AppendPaperSize "ANSI D верт", "D", 0.559, 0.864
  AppendPaperSize "ANSI E гориз", "E", 1.121, 0.864
  AppendPaperSize "ANSI E верт", "E", 0.864, 1.121
  
  gIsShiftPressed = False
  gIniFilePath = gFSO.BuildPath(swApp.GetCurrentMacroPathFolder, IniFileName)
  If Not gFSO.FileExists(gIniFilePath) Then
    CreateDefaultIniFile
  End If
  gIsApplyAndExit = GetBooleanSetting(KeyApplyAndExit)
    
End Function
 
Function CreateDefaultIniFile() 'hide

  Const DefaultText = "[" + MainSection + "]" + vbNewLine _
    + KeyApplyAndExit + " = False" + vbNewLine

  Dim objStream As Stream
      
  Set objStream = New Stream
  objStream.Open
  objStream.WriteText DefaultText
  objStream.SaveToFile gIniFilePath
  objStream.Close

End Function

Function InitWidgets() 'hide

  Dim BaseMaterials As Dictionary
  Dim ResultMaterials() As String
  Dim I As Variant
  Dim K As Integer

  OutputTypeAndName
  GetConfNames  'set gModelConfNames
  InitWidgetFrom MainForm.ConfBox, gModelConfNames
  InitWidgetFrom MainForm.DevelBox, UserDesigner
  InitWidgetFrom MainForm.FormatBox, UserFormat
  InitWidgetFrom MainForm.NameBox, UserName
  InitWidgetFrom MainForm.NoteBox, UserNote
  
  If gIsUnnamed Then
    MainForm.MassLab.Enabled = False
    MainForm.MassBox.Enabled = False
    MainForm.MassChk.Enabled = False
  Else
    MainForm.MassBox.AddItem ("")
    InitWidgetFrom MainForm.MassBox, UserMass
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
    InitWidgetFrom MainForm.BlankBox, UserBlank
    InitWidgetFrom MainForm.lenBox, UserLen
    InitWidgetFrom MainForm.widBox, UserWid
  End If

  If gIsUnnamed Or gIsAssembly Then
    MainForm.SizeLab.Enabled = False
    MainForm.SizeBox.Enabled = False
    MainForm.SizeChk.Enabled = False
    MainForm.MaterialBox.Enabled = False
    MainForm.MaterialLab.Enabled = False
  Else
    MainForm.SizeBox.AddItem ("")  ' for Equation
    InitWidgetFrom MainForm.SizeBox, UserSize
    Set BaseMaterials = ReadMaterialNames("Материалы.sldmat")
    If BaseMaterials.Count > 0 Then
      ReDim ResultMaterials(BaseMaterials.Count - 1)
      K = 0
      For Each I In UserMaterials
        If BaseMaterials.Exists(I) Then
          ResultMaterials(K) = I
          K = K + 1
          BaseMaterials.Remove I
        End If
      Next
      
      For Each I In BaseMaterials.Keys
        ResultMaterials(K) = I
        K = K + 1
      Next
    End If
    MainForm.MaterialBox.AddItem sEmpty
    InitWidgetFrom MainForm.MaterialBox, ResultMaterials
  End If

  If gIsDrawing Then
    MainForm.MiniSignBox.AddItem ""
    InitWidgetFrom MainForm.MiniSignBox, UserDrawingTypes.Keys
    InitWidgetFrom MainForm.OrgBox, UserOrganization
    InitWidgetFrom MainForm.DraftBox, UserDrafter
    InitWidgetFrom MainForm.CheckingBox, UserChecking
    InitRealFormatBox '''установка основных надписей
    gCodeRegexPattern = CreateCodeRegexPattern
    If Not CheckIsFirstSheet(gDrawing, gSheet.GetName) Then
      MainForm.FormatLab.Enabled = False
      MainForm.FormatBox.Enabled = False
      MainForm.FormatChk.Enabled = False
    End If
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
  
  If gIsApplyAndExit Then
    MainForm.ApplyBut.Caption = "OK"
  End If
    
End Function

Function CheckScaleAndReport() 'hide

  Debug.Assert gIsDrawing
  
  Dim I As Variant
  Dim J As Variant
  Dim AView As View
  Dim AScale As Variant

  For Each I In gDrawing.GetViews
    For Each J In I
      Set AView = J
      AScale = AView.ScaleRatio
      If Not CheckIsStandardScale(AScale(0), AScale(1)) Then
        MainForm.labWarningSelected.Caption = "НЕСТАНДАРТНЫЙ МАСШТАБ """ & AView.Name & """ " & Str(AScale(0)) & ":" & Str(AScale(1))
        MainForm.labWarningSelected.ForeColor = &HFF&
        MainForm.labWarningSelected.Visible = True
        Exit Function
      End If
NextI:
    Next
  Next
  
End Function

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
          ReadHeaderValues UserName, I, Lines, EndLines
        Case HeaderInFile(pShortDrawingType)
          ReadDrawingTypes I, Lines, EndLines
        Case HeaderInFile(pSize)
          ReadHeaderValues UserSize, I, Lines, EndLines
        Case HeaderInFile(pBlank)
          ReadHeaderValues UserBlank, I, Lines, EndLines
        Case HeaderInFile(pDesigner)
          ReadHeaderValues UserDesigner, I, Lines, EndLines
        Case HeaderInFile(pDrafter)
          ReadHeaderValues UserDrafter, I, Lines, EndLines
        Case HeaderInFile(pFormat)
          ReadHeaderValues UserFormat, I, Lines, EndLines
        Case HeaderInFile(pOrganization)
          ReadHeaderValues UserOrganization, I, Lines, EndLines
        Case HeaderInFile(pMass)
          ReadHeaderValues UserMass, I, Lines, EndLines
        Case HeaderInFile(pNote)
          ReadHeaderValues UserNote, I, Lines, EndLines
        Case HeaderInFile(pChecking)
          ReadHeaderValues UserChecking, I, Lines, EndLines
        Case HeaderInFile(pApprover)
          ReadHeaderValues UserApprover, I, Lines, EndLines
        Case HeaderInFile(pTechControl)
          ReadHeaderValues UserTechControl, I, Lines, EndLines
        Case HeaderInFile(pNormControl)
          ReadHeaderValues UserNormControl, I, Lines, EndLines
        Case HeaderInFile(pLen)
          ReadHeaderValues UserLen, I, Lines, EndLines
        Case HeaderInFile(pWid)
          ReadHeaderValues UserWid, I, Lines, EndLines
        Case HeaderInFile(pMaterial)
          ReadHeaderValues UserMaterials, I, Lines, EndLines
      End Select
      I = I + 1
    Wend
  End If
    
End Function

Function OpenSettingsFile() As Boolean

  Dim Cmd As String
  Dim FileName As String
  Dim Text As String
  Dim Fst As Object
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
  
  FileName = gConfigPath + SettingsFile
  If Not gFSO.FileExists(FileName) Then
    If Not gFSO.FolderExists(gConfigPath) Then
      gFSO.CreateFolder gConfigPath
    End If
      
    Text = _
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
      HeaderInFile(pWid) + vbNewLine + ";;;" + vbNewLine
      
    Set Fst = CreateObject("ADODB.Stream")
    Fst.Type = 2  'we want to save text/string data.
    Fst.Charset = "utf-16le"
    Fst.Open
    Fst.WriteText Text
    Fst.SaveToFile FileName, 2
  End If
  
  Cmd = "notepad """ + FileName + """"
  Shell Cmd, vbNormalFocus
    
End Function

Function SetModelFromActiveDoc() 'hide

  Set gModel = gDoc
  gCurConf = gModel.GetActiveConfiguration.Name
    
End Function

Function EditorRun() As Boolean

  Dim IsSelectedComp As Boolean
  Dim Form As MainForm
  Dim HaveErrors As ErrorCode
  Dim AView As View
  Dim Selected As Component2
  Dim SheetProperties As Variant
  
  IsSelectedComp = False
  HaveErrors = ErrorCode.Ok
  gIsDrawing = IsDrawing(gDoc)
  If gIsDrawing Then
    Set gDrawing = gDoc
    Set gSheet = gDrawing.GetCurrentSheet
    If gSheet Is Nothing Then
      HaveErrors = ErrorCode.EmptySheet
    Else
      SheetProperties = gSheet.GetProperties
      gSheetScale1 = SheetProperties(2)
      gSheetScale2 = SheetProperties(3)
      gIsFirstAngle = SheetProperties(4)
      Set AView = SelectView
      If AView Is Nothing Then
        HaveErrors = ErrorCode.EmptySheet
      Else
        Set gModel = AView.ReferencedDocument
        If gModel Is Nothing Then
          HaveErrors = ErrorCode.EmptyView
        Else
          gCurConf = AView.ReferencedConfiguration
        End If
      End If
    End If
  Else
    If gDoc.GetType = swDocASSEMBLY Then
      Set Selected = GetSelectedComponent
      If Selected Is Nothing Then
        SetModelFromActiveDoc
      Else
        Set gModel = Selected.GetModelDoc2
        gCurConf = Selected.ReferencedConfiguration
        IsSelectedComp = True
      End If
    Else
      SetModelFromActiveDoc
    End If
  End If
  If HaveErrors = ErrorCode.Ok Then
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
  Select Case HaveErrors
    Case ErrorCode.Ok
      If gIsDrawing Then
        CheckScaleAndReport
      Else
        MainForm.labWarningSelected.Visible = IsSelectedComp
      End If
      MainForm.Show
    Case ErrorCode.EmptyView
      MsgBox ("Пустой вид. Нет ссылки на модель.")
    Case ErrorCode.EmptySheet
      MsgBox ("Пустой лист. Модель не обнаружена.")
  End Select
  
End Function

Function GetSelectedComponent() As Component2

  Set GetSelectedComponent = gDoc.SelectionManager.GetSelectedObjectsComponent3(1, -1)
  
End Function

Sub SetLineStyle(Object_type As swUserPreferenceIntegerValue_e, Value As Integer)

  gDrawExt.SetUserPreferenceInteger Object_type, swDetailingNoOptionSpecified, Value
  
End Sub

Function ReadMaterialNames(FileName As String) As Dictionary

  Dim Result As Dictionary
  Dim FullFilename As String
  Dim I As Variant
  Dim Matches As MatchCollection
  Dim FStream As TextStream
  Dim TextAll As String
  
  Set Result = New Dictionary
  FullFilename = gConfigPath + FileName
  If gFSO.FileExists(FullFilename) Then
    Set FStream = gFSO.OpenTextFile(FullFilename, ForReading, False, TristateTrue)
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

Function FindView() As View

  Dim PropViewName As String
  Dim AllViews As Variant  'array of array
  Dim SheetViews As Variant
  Dim J As Integer
  Dim AView As View
  
  PropViewName = gSheet.CustomPropertyView
  AllViews = gDrawing.GetViews
  For Each SheetViews In AllViews
    For J = 1 To UBound(SheetViews)
      Set AView = SheetViews(J)
      If AView.Name = PropViewName Then
        Set FindView = AView
        Exit Function
      End If
    Next
  Next
  If FindView Is Nothing Then
    Set FindView = AllViews(0)(1)  'first view
    'FIX IT: ERROR if 1st sheet don't have views
  End If
     
End Function

Function SelectView() As View

  Dim Selected As Object
  
  Set Selected = gDrawing.SelectionManager.GetSelectedObject5(1)
  If Selected Is Nothing Then
    Set SelectView = FindView
  ElseIf Not TypeOf Selected Is View Then
    Set SelectView = FindView
  Else
    Set SelectView = Selected
  End If
    
End Function

Sub SaveAsMy(NewName As String, OldName As String)

  Dim Error As Long, Warning As Long

  If gFSO.FileExists(NewName) Then
    If MsgBox("Файл с таким именем существует. Заменить?", vbOKCancel) = vbCancel Then
      Exit Sub
    End If
  End If
  If gDoc.Extension.SaveAs(NewName, swSaveAsCurrentVersion, _
                           swSaveAsOptions_AvoidRebuildOnSave, _
                           Nothing, Error, Warning) Then
      If OldName <> "" Then
        Kill (OldName)
      End If
  Else
    MsgBox ("Не удалось сохранить файл")
  End If
    
End Sub

'Only for drawings!
Function ZoomToSheet()  'hide

  Dim Width As Double
  Dim Height As Double

  gDrawing.GetCurrentSheet.GetSize Width, Height
  gDoc.ViewZoomTo2 0, 0, 0, Width, Height, 0
    
End Function

Function GetEquationThickness(Conf As String, ToAll As Boolean, NameModel As String) As String

  Const Temp As String = "__temp__"
  Dim Mgr As CustomPropertyManager
  Dim ThName_ As Variant
  Dim ThName As String
  Dim ResolvedValue As String
  Dim RawValue As String
  Dim VariantThicknessName(1) As String
  
  GetEquationThickness = ""
  Set Mgr = gModel.Extension.CustomPropertyManager(Conf)
  VariantThicknessName(0) = "Толщина"
  VariantThicknessName(1) = "Thickness"
  'variantThicknessName(2) = "Grubos'c'"
  For Each ThName_ In VariantThicknessName
    ThName = ThName_
    SetProp Mgr, Temp, Equal(ThName, ToAll, Conf, gNameModel)
    Mgr.Get5 Temp, False, RawValue, ResolvedValue, False
    If IsNumeric(StrConv(ResolvedValue, vbUnicode)) Then  'IsNumeric make error with raw 'resolvedValue' without StrConv
      GetEquationThickness = RawValue
      Exit For
    End If
  Next
  Mgr.Delete2 Temp
    
End Function

Function ExitApp() 'hide

  Unload MainForm
  End

End Function

Function GetIsFastener() As Boolean

  Dim Value As String
  Dim RawValue As String
  Dim WasResolved As Boolean

  GetIsFastener = False
  If gModelManager.Get5(pIsFastener, False, RawValue, Value, WasResolved) <> swCustomInfoGetResult_NotPresent Then
    GetIsFastener = (Value = IsFastenerTrue)
  End If

End Function

Function SetIsFastener() 'hide

  gModelManager.Delete2 pIsFastener
  gModelManager.Add3 pIsFastener, swCustomInfoNumber, IsFastenerTrue, swCustomPropertyDeleteAndAdd

End Function

' Без точек "." в наименовании
Sub SplitNameAndSign(Line As String, Conf As String, ByRef Designation As String, _
                     ByRef Name As String, ByRef Code As String)
                     
  Const Flat As String = "SM-FLAT-PATTERN"
  Dim RegexAsm As RegExp
  Dim RegexPrt As RegExp
  Dim Matches As Object
  Dim Z As Variant
  
  Designation = Line
  Name = Line
  Code = ""
  
  Set RegexAsm = New RegExp
  RegexAsm.Pattern = "(.*\..*[0-9] *)(" + gCodeRegexPattern + ") ([^.]+)"
  RegexAsm.IgnoreCase = True
  RegexAsm.Global = True
  
  Set RegexPrt = New RegExp
  RegexPrt.Pattern = "(.*\.[^ ]+) ([^.]+)"
  RegexPrt.IgnoreCase = True
  RegexPrt.Global = True
  
  If RegexAsm.Test(Line) Then
    Set Matches = RegexAsm.Execute(Line)
    Designation = Trim(Matches(0).SubMatches(0))
    Code = Matches(0).SubMatches(1)
    Name = Trim(Matches(0).SubMatches(2))
  ElseIf RegexPrt.Test(Line) Then
    Set Matches = RegexPrt.Execute(Line)
    Designation = Trim(Matches(0).SubMatches(0))
    Name = Trim(Matches(0).SubMatches(1))
  End If
  
  If Conf Like "*" & Flat Then
    Conf = Left(Conf, Len(Conf) - Len(Flat))
  End If
  If Not IsBaseConf(Conf) Then
    Conf = Split(Conf)(0)  '-00 откр.
    If Not IsBaseConf(Conf) Then
      MainForm.SignChk.Value = False ' running event
      Designation = Designation & "-" & Conf
    End If
  End If
  
End Sub

Function IsBaseConf(Conf As String) As Boolean

  Select Case Conf
    Case "00", "По умолчанию", "Default"
      IsBaseConf = True
    Case Else
      IsBaseConf = False
  End Select

End Function

' Устанавливает значения gItems из свойств, игнорируя существующие
Sub ReadProp(Manager As CustomPropertyManager, Conf As String, props() As String)

  Const UseCached = False
  Dim I As Variant
  Dim Prop As String
  Dim Item As DataItem
  Dim GetPropRes As swCustomInfoGetResult_e
  Dim WasResolved As Boolean
  Dim RawValue As String
  Dim Value As String
  
  If Not gItems.Exists(Conf) Then
    gItems.Add Conf, New Dictionary
  End If
  
  For Each I In props
    Prop = I
    
    If Not gItems(Conf).Exists(Prop) Then
      gItems(Conf).Add Prop, New DataItem
    End If
    
    Set Item = gItems(Conf)(Prop)
    GetPropRes = Manager.Get5(Prop, UseCached, RawValue, Value, WasResolved)
    Item.RawValue = Trim(RawValue)
    Item.Value = Trim(Value)
    
    If Conf = CommonSpace Then
      Item.FromAll = True
    Else
      Item.FromAll = (GetPropRes = swCustomInfoGetResult_NotPresent)
    End If
    
    If Prop = pMaterial Then
      Item.NewValue = Item.Value
    Else
      Item.NewValue = Item.RawValue
    End If
  Next
    
End Sub

Sub SetBoxValue2(Chk As CheckBox, Prop As String, Conf As String)

  Dim Item As DataItem
  
  Set Item = gItems(Conf)(Prop)

  If Not Chk Is Nothing Then
    If Chk.Value <> Item.FromAll Then
      Chk.Value = Item.FromAll
    Else
      ChangeChecked Prop
    End If
  Else
    ChangeChecked Prop
  End If
    
End Sub

Sub FromAllChecked(Chk As CheckBox, Box As Object, Prop As String, Conf As String, _
                   FromAll As Boolean, SetFirstItem As Boolean)
                   
  Dim Cmb As ComboBox
  Dim Value As String
  
  If ReadOldAfterChecked Then
    ReadBox Box, Chk, Conf, Prop, False
  End If
  If Prop = pSize Then
    ChangeSizeEqual (Conf)
  ElseIf Prop = pMass Then
    ChangeMassEqual (Conf)
  End If
  
  If FromAll Then
    Value = gItems(CommonSpace)(Prop).NewValue
  Else
    Value = gItems(Conf)(Prop).NewValue
  End If
  If SetFirstItem And Value = "" And TypeOf Box Is ComboBox Then
    Set Cmb = Box
    If Cmb.ListCount > 0 Then
      Value = Cmb.List(0)
    End If
  End If
  
  If Box.Enabled Then
    If TypeOf Box Is ComboBox Then
      Set Cmb = Box
      If Cmb.Style = fmStyleDropDownList Then
        If ExistsInCombo(Cmb, Value) Then
          SetComboInExistValue Box, Value
        ElseIf Cmb.ListCount > 0 Then
          Cmb.ListIndex = 0
        End If
      Else
        Cmb.Text = Value
      End If
    Else
      Box.Text = Value
    End If
  End If
  
End Sub

' Устанавливает значения gItems из формы
' conf - конфигурация ИЛИ элемент списка вырезов
Sub ReadBox(Box As Object, Chk As CheckBox, Conf As String, Prop As String, forward As Boolean)

  Dim TargetConf As String
  
  If Not gItems.Exists(Conf) Then
    gItems.Add Conf, New Dictionary
  End If
  If Not gItems(Conf).Exists(Prop) Then
    gItems(Conf).Add Prop, New DataItem
  End If
  
  If Chk Is Nothing And Conf = CommonSpace Then
    gItems(CommonSpace)(Prop).FromAll = True
    gItems(CommonSpace)(Prop).NewValue = Box.Text
  ElseIf Prop = pMaterial Then
    gItems(Conf)(Prop).FromAll = False
    gItems(Conf)(Prop).NewValue = Box.Text   'уравнение MaterialEqual устанавливается в SetProp2
  Else
    gItems(Conf)(Prop).FromAll = Chk.Value
    If forward Then
      If Chk.Value Then
        TargetConf = CommonSpace
      Else
        TargetConf = Conf
      End If
    Else
      If Chk.Value Then
        TargetConf = Conf
      Else
        TargetConf = CommonSpace
      End If
    End If
    gItems(TargetConf)(Prop).NewValue = Box.Text
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
  ReadBox Box, Nothing, CommonSpace, Property, True
    
End Sub

Sub ChangeChecked(Prop As String)

  Select Case Prop
    Case pDesignation
      FromAllChecked MainForm.SignChk, MainForm.SignBox, pDesignation, gCurConf, MainForm.SignChk.Value, False
    Case pName
      FromAllChecked MainForm.NameChk, MainForm.NameBox, pName, gCurConf, MainForm.NameChk.Value, False
    Case pNameEN
      FromAllChecked MainForm.NameChk, MainForm.NameBoxEN, pNameEN, gCurConf, MainForm.NameChk.Value, False
    Case pNamePL
      FromAllChecked MainForm.NameChk, MainForm.NameBoxPL, pNamePL, gCurConf, MainForm.NameChk.Value, False
    Case pNameUA
      FromAllChecked MainForm.NameChk, MainForm.NameBoxUA, pNameUA, gCurConf, MainForm.NameChk.Value, False
    Case pBlank
      FromAllChecked MainForm.BlankChk, MainForm.BlankBox, pBlank, gCurConf, MainForm.BlankChk.Value, False
    Case pFormat
      FromAllChecked MainForm.FormatChk, MainForm.FormatBox, pFormat, gCurConf, MainForm.FormatChk.Value, False
    Case pNote
      FromAllChecked MainForm.NoteChk, MainForm.NoteBox, pNote, gCurConf, MainForm.NoteChk.Value, False
    Case pDesigner
      FromAllChecked MainForm.DevelChk, MainForm.DevelBox, pDesigner, gCurConf, MainForm.DevelChk.Value, True
    Case pSize
      FromAllChecked MainForm.SizeChk, MainForm.SizeBox, pSize, gCurConf, MainForm.SizeChk.Value, False
    Case pMass
      FromAllChecked MainForm.MassChk, MainForm.MassBox, pMass, gCurConf, MainForm.MassChk.Value, True
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
      FromAllChecked MainForm.lenChk, MainForm.lenBox, pLen, gCurConf, MainForm.lenChk.Value, False
    Case pWid
      FromAllChecked MainForm.widChk, MainForm.widBox, pWid, gCurConf, MainForm.widChk.Value, False
  End Select
   
End Sub

Function CreateBaseDesignation() 'hide

  Dim MainDesignation As String
  Dim ResolvedValue As String
  Dim RawValue As String
  Dim WasResolved As Boolean
  
  If gModel.Extension.CustomPropertyManager(gMainConf).Get5(pDesignation, False, RawValue, ResolvedValue, WasResolved) = swCustomInfoGetResult_NotPresent Then
    gModel.Extension.CustomPropertyManager("").Get5 pDesignation, False, RawValue, ResolvedValue, WasResolved
  End If
  
  gBaseDesignation = GetBaseDesignation(ResolvedValue)
  
End Function

Function Execute() 'hide

  ReadForm gCurConf
  
  gModel.SetReadOnlyState False  'must be first!
  
  WriteModelProperties
  
  If Not gIsAssembly Then
    If MainForm.IsFastenerChk.Value Then
      SetIsFastener
    End If
  End If
  
  If gIsDrawing Then
    CreateBaseDesignation
    WriteDrawingProperties
    SetSpeedFormat
    ChangeLineStyles
  End If
  
  gDoc.ForceRebuild3 True
  
  If gIsDrawing Then
    TryRenameDraft MainForm.DrawNameBox.Text
  End If
    
End Function

Sub ChangeMassEqual(Conf As String)

  If Not gIsUnnamed Then
    MainForm.MassBox.List(0) = Equal("SW-Mass", gItems(Conf)(pMass).FromAll, Conf, gNameModel)
  End If
    
End Sub

Sub ChangeSizeEqual(Conf As String)

  If Not gIsUnnamed And Not gIsAssembly Then
    MainForm.SizeBox.List(0) = GetEquationThickness(Conf, gItems(Conf)(pSize).FromAll, gNameModel)
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

  Dim NewMaterial As String
  
  NewMaterial = gItems(Conf)(pMaterial).NewValue
  If NewMaterial <> sEmpty And NewMaterial <> "" Then
    gModel.SetMaterialPropertyName2 Conf, MaterialDB, NewMaterial  'it's method of PartDoc
  End If
    
End Sub

Function SetSpeedFormat() 'hide

  Dim TemplateName As String
  Dim Width As Double
  Dim Height As Double
  Dim OldWidth As Double
  Dim OldHeight As Double
  Dim SizeName As String
  Dim PaperChoice As String
  Dim FormatChoice As String
  
  FormatChoice = MainForm.RealFormatBox.Text
  PaperChoice = MainForm.PaperSizeBox.Text
  
  If (FormatChoice <> CurrentChoice) Or (PaperChoice <> CurrentChoice) Then
    GetSheetSizes Width, Height, SizeName, OldWidth, OldHeight, gSheet, PaperChoice
    
    If FormatChoice <> CurrentChoice Then
      TemplateName = gFSO.BuildPath(gConfigPath, FormatChoice & ".SLDDRT")
      If Not gFSO.FileExists(TemplateName) Then
        MsgBox ("Файл " + TemplateName + " не найден.")
        Exit Function
      End If
      
      gDrawing.SetupSheet5 _
        gSheet.GetName, swDwgPapersUserDefined, swDwgTemplateCustom, _
        gSheetScale1, gSheetScale2, gIsFirstAngle, TemplateName, 0, 0, _
        gSheet.CustomPropertyView, True
        
      gSheet.GetSize OldWidth, OldHeight
      
      gDrawing.SetupSheet5 _
        gSheet.GetName, swDwgPapersUserDefined, swDwgTemplateNone, _
        gSheetScale1, gSheetScale2, gIsFirstAngle, "", Width, Height, _
        gSheet.CustomPropertyView, False
    Else
      gSheet.SetSize swDwgPapersUserDefined, Width, Height
    End If
    ResizeSheetFormat Width, Height, gSheet, gDoc, gDrawing, OldWidth, OldHeight, SizeName
    gDoc.ForceRebuild3 True
    gDoc.ViewZoomTo2 0, 0, 0, Width, Height, 0
  End If
  
End Function

Function OutputTypeAndName() 'hide

  MainForm.ModelNameBox.Text = gFSO.GetBaseName(gNameModel)
  If gIsDrawing Then
    MainForm.DrawNameBox.Enabled = True
    MainForm.DrawNameLab.Enabled = True
    MainForm.DrawNameBox.Text = gFSO.GetBaseName(gDoc.GetPathName)
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

Function InitRealFormatBox() 'hide

  Dim FileName As String
  
  MainForm.RealFormatBox.AddItem (CurrentChoice)
  MainForm.RealFormatBox.Text = MainForm.RealFormatBox.List(0)
  
  FileName = Dir(gConfigPath & "*.SLDDRT")
  While FileName <> ""
    MainForm.RealFormatBox.AddItem gFSO.GetBaseName(FileName)
    FileName = Dir()
  Wend
    
End Function

Sub SetModelProp(Conf As String, Prop As String, Item As DataItem)

  Dim ConfManager As CustomPropertyManager
  
  Set ConfManager = gModelExt.CustomPropertyManager(Conf)
  If Conf <> CommonSpace And Item.FromAll Then
    ConfManager.Delete (Prop)
  Else
    SetProp2 ConfManager, Prop, Item, Conf
  End If
    
End Sub

Function WriteModelProperties() 'hide

  Dim I As Variant
  Dim J As Variant
  Dim Conf As String
  Dim Prop As String
  Dim Item As DataItem
  
  For Each I In gItems.Keys
    Conf = I
    For Each J In ModelProps
      Prop = J
      Set Item = gItems(Conf)(Prop)
      
      Select Case Prop
        Case pBlank, pSize, pLen, pWid
          If Not gIsAssembly Then
            SetModelProp Conf, Prop, Item
          End If
        Case pMaterial
          If Not gIsAssembly Then
            If Not gIsUnnamed Then
              SetModelProp Conf, Prop, Item
            End If
            SetMaterial Conf
          End If
        Case Else
          SetModelProp Conf, Prop, Item
      End Select
      
    Next
  Next
    
End Function

'TODO: убрать чертежные свойства из gItems, читать их прямо из формы
Function WriteDrawingProperties() 'hide

  Dim ToAll As Boolean
  Dim Item As Dictionary
  Dim DrawingCode As String

  ToAll = True
  Set Item = gItems(CommonSpace)
  DrawingCode = MainForm.MiniSignBox.Text
  'см. массив drawProps
  SetProp gDrawManager, pShortDrawingType, DrawingCode
  SetProp2 gDrawManager, pOrganization, Item(pOrganization)
  SetProp2 gDrawManager, pDrafter, Item(pDrafter)
  SetProp2 gDrawManager, pChecking, Item(pChecking)  'before: userChecking(0)
  SetProp gDrawManager, pApprover, UserApprover(0)
  SetProp gDrawManager, pNormControl, UserNormControl(0)
  SetProp gDrawManager, pTechControl, UserTechControl(0)
  SetProp gDrawManager, pLongDrawingType, IIf(DrawingCode = "", "", MainForm.CodeBox.Text)
  SetProp gDrawManager, pBaseDesignation, gBaseDesignation
    
End Function

Sub SetShiftStatus(Shift As Integer)
  
  gIsShiftPressed = Shift And 1
  
End Sub

Sub ApplyAndExitIfNeeded(IsShiftPressed As Boolean)

  Dim NeedExit As Boolean
  
  NeedExit = gIsApplyAndExit Xor IsShiftPressed

  Execute
  If NeedExit Then
    ExitApp
  End If

End Sub

Sub ExitByKey(KeyCode As MSForms.ReturnInteger, Shift As Integer)
  
  If KeyCode = vbKeyReturn Then
    ApplyAndExitIfNeeded Shift = 1
  End If
    
End Sub
