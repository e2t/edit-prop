Attribute VB_Name = "Main"
'Written in 2014-2017 by Eduard E. Tikhenko <aquaried@gmail.com>
'
'To the extent possible under law, the author(s) have dedicated all copyright
'and related and neighboring rights to this software to the public domain
'worldwide. This software is distributed without any warranty.
'You should have received a copy of the CC0 Public Domain Dedication along
'with this software.
'If not, see <http://creativecommons.org/publicdomain/zero/1.0/>

Public Const macroName As String = "EditProp"
Public Const macroSection As String = "Main"

'Свойства модели, записаны в массиве modelProps
Public Const pDesignation As String = "Обозначение"
Public Const pMaterial As String = "Материал"
Public Const pName As String = "Наименование"
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
Public Const separator As String = ";"
Public Const separator2 As String = "="
Public Const settingsFile As String = "Настройки.txt"
Public Const ppExclude As String = "Исключить"
Public Const sEmpty = " "

Private Enum ErrorCode
    ok = 0
    emptyView = 1
    emptySheet = 2
End Enum

Public gApp As Object
Public gDoc As ModelDoc2
Public gConfigPath As String
Public gModel As ModelDoc2
Public gModelConfNames() As String
Public gModelCutsNames() As String
Public gModelCutsCount As Integer
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
Private stdFile As String

Public userName() As String
Public userDrawingTypes As Dictionary
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
Public modelProps(11) As String
Public drawProps(8) As String

Public Const tabNumberConf As Integer = 0
Public Const tabNumberCuts As Integer = 1

Public indexLastConf As Integer
Public indexLastCut As Integer

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
    
    ReDim userName(0)
    userName(0) = ""
   
    Set userDrawingTypes = New Dictionary
    
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
    
    drawProps(0) = pDrafter
    drawProps(1) = pShortDrawingType
    drawProps(2) = pLongDrawingType
    drawProps(3) = pOrganization
    drawProps(4) = pChecking
    drawProps(5) = pApprover
    drawProps(6) = pTechControl
    drawProps(7) = pNormControl
    drawProps(8) = pBaseDesignation
    
    indexLastConf = 0
    indexLastCut = 0
    
    SetConfigFolder
    ReadSettings  'only after SetConfigFolder
End Function

Function SetConfigFolder() As Boolean
    gConfigPath = gApp.GetCurrentMacroPathFolder() + "\config\"
End Function

Function IsDrawing(doc As ModelDoc2) As Boolean
    IsDrawing = CBool(doc.GetType = swDocumentTypes_e.swDocDRAWING)
End Function

Function IsPropertyExist(manager As CustomPropertyManager, prop As String) As Boolean
    Dim names As Variant: names = manager.GetNames()
    IsPropertyExist = False
    If Not IsEmpty(names) Then
        If IndexInArray(prop, names) <> -1 Then
            IsPropertyExist = True
        End If
    End If
End Function

Function SetProp(manager As CustomPropertyManager, prop As String, value As String) As Boolean
    manager.Add2 prop, swCustomInfoText, ""
    SetProp = Not CBool(manager.Set(prop, value))
End Function

Function ShortFileNameExt(name As String) As String
    ShortFileNameExt = ""
    If name <> "" Then
        Dim list() As String: list = Split(name, "\")
        ShortFileNameExt = list(UBound(list))
    End If
End Function

Function IsArrayEmpty(anArray As Variant) As Boolean
    Dim i As Integer
  
    On Error Resume Next
        i = UBound(anArray, 1)
    IsArrayEmpty = Err.number <> 0
End Function

Sub InitWidgetFrom(widget As Object, values As Variant)
    If Not IsArrayEmpty(values) Then
        For Each i In values
            widget.AddItem (i)
        Next
    End If
End Sub

Function Equal(swProp As String, toAll As Boolean, conf As String, nameModel As String) As String
    Dim confText As String: confText = ""
    If conf <> commonSpace And Not toAll Then
        confText = "@@" + conf
    End If
    Equal = """" + swProp + confText + "@" + nameModel + """"
End Function

Sub ReadHeaderValues(ByRef userStr() As String, ByRef count As Long, lines() As String, endLines As Long)
    If count < endLines Then
        count = count + 1
        If lines(count) <> "" Then
            userStr = Split(lines(count), separator)
        End If
    End If
End Sub

Sub ReadDrawingTypes(ByRef count As Long, lines() As String, endLines As Long)
    Dim X As Variant
    Dim indexSeparator As Integer
    Dim names() As String
    Dim shortName As String
    Dim longName As String
    
    If count < endLines Then
        count = count + 1
        If lines(count) <> "" Then
            For Each X In Split(lines(count), separator)  'String()
                indexSeparator = InStr(X, separator2)  'равен нулю, если разделитель не найден
                If indexSeparator > 0 Then
                    names = Split(X, separator2)
                    shortName = names(0)
                    longName = names(1)
                Else
                    shortName = X
                    longName = ""
                End If
                If Not userDrawingTypes.Exists(shortName) Then
                    userDrawingTypes.Add shortName, longName
                End If
            Next
        End If
    End If
End Sub

Function ReadSettings() As Boolean
    Dim lines() As String: lines = ReadLinesFrom(settingsFile)
    Dim endLines As Long: endLines = UBound(lines)
    Dim i As Long: i = LBound(lines)
    While i <= endLines
        Select Case lines(i)
            Case HeaderInFile(pName)
                ReadHeaderValues userName, i, lines, endLines
            Case HeaderInFile(pShortDrawingType)
                ReadDrawingTypes i, lines, endLines
            Case HeaderInFile(pSize)
                ReadHeaderValues userSize, i, lines, endLines
            Case HeaderInFile(pBlank)
                ReadHeaderValues userBlank, i, lines, endLines
            Case HeaderInFile(pDesigner)
                ReadHeaderValues userDesigner, i, lines, endLines
            Case HeaderInFile(pDrafter)
                ReadHeaderValues userDrafter, i, lines, endLines
            Case HeaderInFile(pFormat)
                ReadHeaderValues userFormat, i, lines, endLines
            Case HeaderInFile(pOrganization)
                ReadHeaderValues userOrganization, i, lines, endLines
            Case HeaderInFile(pMass)
                ReadHeaderValues userMass, i, lines, endLines
            Case HeaderInFile(pNote)
                ReadHeaderValues userNote, i, lines, endLines
            Case HeaderInFile(pChecking)
                ReadHeaderValues userChecking, i, lines, endLines
            Case HeaderInFile(pApprover)
                ReadHeaderValues userApprover, i, lines, endLines
            Case HeaderInFile(pTechControl)
                ReadHeaderValues userTechControl, i, lines, endLines
            Case HeaderInFile(pNormControl)
                ReadHeaderValues userNormControl, i, lines, endLines
            Case HeaderInFile(pLen)
                ReadHeaderValues userLen, i, lines, endLines
            Case HeaderInFile(pWid)
                ReadHeaderValues userWid, i, lines, endLines
            Case HeaderInFile(pMaterial)
                ReadHeaderValues userMaterials, i, lines, endLines
            Case HeaderInFile(ppExclude)
                ReadHeaderValues userPreExclude, i, lines, endLines
        End Select
        i = i + 1
    Wend
End Function

Function HeaderInFile(prop As String) As String
    HeaderInFile = "[" + prop + "]"
End Function

Function OpenSettingsFile() As Boolean
    Dim filename As String: filename = gConfigPath + settingsFile
    If Not IsFileExists(filename) Then
        If Not IsFileExists(gConfigPath) Then
            FileIO.FileSystem.CreateDirectory (gConfigPath)
        End If
        
        Dim text As String: text = _
                HeaderInFile(pBlank) + vbNewLine + ";;;" + vbNewLine + vbNewLine + _
                HeaderInFile(pMass) + vbNewLine + "см. табл." + vbNewLine + vbNewLine + _
                HeaderInFile(pName) + vbNewLine + ";;;" + vbNewLine + vbNewLine + _
                HeaderInFile(pNormControl) + vbNewLine + "Юриков" + vbNewLine + vbNewLine + _
                HeaderInFile(pDrafter) + vbNewLine + ";;;" + vbNewLine + vbNewLine + _
                HeaderInFile(pOrganization) + vbNewLine + "ООО ""Эко-Инвест"";ЗАО НПФ ""Экотон""" + vbNewLine + vbNewLine + _
                HeaderInFile(pShortDrawingType) + vbNewLine + _
                "СБ=Сборочный чертеж;ВО=Чертеж общего вида;СП=Спецификация;МЧ=Монтажный чертеж;ГЧ=Габаритный чертеж;УЧ=Упаковочный чертеж;ТЧ=Теоретический чертеж;" + _
                "МЭ=Электромонтажный чертеж;ПЭ=Перечень элементов;ПЗ=Пояснительная записка;ТБ=Таблица;РР=Расчет;И=Инструкция;ТУ=Технические условия;" + _
                "ПМ=Программа и методика испытаний;ВС=Ведомость спецификаций;ВД=Ведомость ссылочных документов;ВП=Ведомость покупных изделий;" + _
                "ВИ=Ведомость разрешения применения покупных изделий;ДП=Ведомость держателей подлинников;ПТ=Ведомость технического предложения;" + _
                "ЭП=Ведомость эскизного проекта;ТП=Ведомость технического проекта;ВДЭ=Ведомость электронных документов" + vbNewLine + vbNewLine + _
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
        
        Dim fsT As Object
        Set fsT = CreateObject("ADODB.Stream")
        fsT.Type = 2  'we want to save text/string data.
        fsT.Charset = "utf-16le"
        fsT.Open
        fsT.WriteText text
        fsT.SaveToFile filename, 2
        
    End If
    Dim cmd As String: cmd = "notepad """ + filename + """"
    Shell cmd, vbNormalFocus
End Function

Function SetModelFromActiveDoc() 'mask for button
    Set gModel = gDoc
    gCurConf = gModel.GetActiveConfiguration.name
End Function

Function EditorRun() As Boolean
    Dim isSelectedComp As Boolean
    isSelectedComp = False
    
    Dim haveErrors As ErrorCode: haveErrors = ErrorCode.ok
    stdFile = gConfigPath + "Чертежный стандарт.sldstd"
    gIsDrawing = IsDrawing(gDoc)
    If gIsDrawing Then
        Set gDrawing = gDoc
        Set gSheet = gDrawing.GetCurrentSheet
        If gSheet Is Nothing Or IsEmpty(gSheet.GetViews) Then
            haveErrors = ErrorCode.emptySheet
        Else
            Dim aView As View: Set aView = SelectView
            Set gModel = aView.ReferencedDocument
            If gModel Is Nothing Then
                haveErrors = ErrorCode.emptyView
            Else
                gCurConf = aView.ReferencedConfiguration
            End If
        End If
    Else
        If gDoc.GetType = swDocASSEMBLY Then
            Dim selected As Component2
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
        gNameModel = ShortFileNameExt(gModel.GetPathName)
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
            Dim form As MainForm
            Set form = New MainForm
            form.labWarningSelected.Visible = isSelectedComp
            form.Show
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

Function ReloadStandard() As Boolean
    If IsFileExists(stdFile) Then
        gDrawExt.LoadDraftingStandard (stdFile)
    End If
End Function

Function ChangeMassUnits() As Boolean
    gModelExt.SetUserPreferenceInteger swUnitSystem, 0, swUnitSystem_Custom
    gModelExt.SetUserPreferenceInteger swUnitsMassPropMass, 0, 3
End Function

Sub ReloadSheet(format As String)
    Const isFirstAngle As Boolean = True
    Dim fullFormatName As String
    Dim oldSheetProp() As Double
    Dim width As Double, height As Double
    Dim scale1 As Double, scale2 As Double
        
    fullFormatName = gConfigPath + format + ".SLDDRT"
    If IsFileExists(fullFormatName) Then
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

Function ShortFileName(ByVal name As String) As String
    Dim ext As String: ext = Right(name, 7)
    If StrComp(ext, ".SLDDRW", vbTextCompare) = 0 _
            Or StrComp(ext, ".SLDPRT", vbTextCompare) = 0 _
            Or StrComp(ext, ".SLDASM", vbTextCompare) = 0 _
            Or StrComp(ext, ".SLDDRT", vbTextCompare) = 0 Then
        name = Mid(name, 1, Len(name) - 7)
    End If
    ShortFileName = name
End Function

Function ReadMaterialNames(filename As String) As String()
    Const aMaterial As String = "material name="
    Dim lenMeterial As Integer: lenMeterial = Len(aMaterial)
    Dim aStrings() As String: aStrings = ReadLinesFrom(filename)
    Dim pos As Integer, startquote As Integer, endquote As Long
    Dim str_v As Variant, name As String, result As String: result = ""
    For Each str_v In aStrings
        Dim str As String: str = str_v
        pos = InStr(1, str, aMaterial)
        If pos > 0 Then
            startquote = InStr(pos + lenMeterial, str, """")
            If startquote > 0 Then
                endquote = InStr(startquote + 1, str, """")
                If endquote > 0 Then
                    name = Mid(str, startquote + 1, endquote - startquote - 1)
                    If result = "" Then
                        result = name
                    Else
                        result = result + vbNewLine + name
                    End If
                End If
            End If
        End If
    Next
    ReadMaterialNames = Split(result, vbNewLine)
End Function

Function IsFileExists(fullname As String) As Boolean
    IsFileExists = False
    If Dir(fullname) <> "" Then
        IsFileExists = True
    End If
End Function

Function ReadLinesFrom(filename As String) As String()
    Dim aStrings() As String, fullFilename As String: fullFilename = gConfigPath + filename
    If IsFileExists(fullFilename) Then
        Open fullFilename For Binary As #1
        Dim skip As String: skip = InputB(2, #1) ' always FF FE (skip first 2 bytes)
        aStrings = Split(InputB(LOF(1), #1), vbNewLine)
        Close #1
    Else
        ReDim aStrings(0)
        aStrings(0) = ""
    End If
    ReadLinesFrom = aStrings
End Function

Function FirstItem(values As Variant) As String
    FirstItem = ""
    If Not IsEmpty(values) Then
        FirstItem = values(0)
    End If
End Function

Function FindView() As View
    Dim propView As String: propView = gSheet.CustomPropertyView
    Dim firstView As View: Set firstView = gDrawing.GetFirstView.GetNextView
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
    Dim selected As Object: Set selected = gDrawing.SelectionManager.GetSelectedObject5(1)
    If selected Is Nothing Then
        Set SelectView = FindView()
    ElseIf Not TypeOf selected Is View Then
        Set SelectView = FindView()
    Else
        Set SelectView = selected
    End If
End Function

Function IndexInArray(valueToFind As Variant, arr As Variant) As Integer
    Dim find As Boolean: find = False
    Dim i As Integer
    For i = LBound(arr) To UBound(arr)
        If arr(i) = valueToFind Then
            find = True
            Exit For
        End If
    Next
    If find Then
        IndexInArray = i
    Else
        IndexInArray = -1
    End If
End Function

Function GetPathName(fullname As String)
    GetPathName = Left(fullname, InStrRev(fullname, "\"))
End Function

Sub SaveAsMy(newname As String, oldname As String)
    If IsFileExists(newname) Then
        If MsgBox("Файл с таким именем существует. Заменить?", vbOKCancel) = vbCancel Then
            Exit Sub
        End If
    End If
    
    Dim error As Long, warning As Long
    If gDoc.Extension.SaveAs(newname, swSaveAsCurrentVersion, _
                             swSaveAsOptions_AvoidRebuildOnSave, _
                             Nothing, error, warning) Then
        If oldname <> "" Then
            Kill (oldname)
        End If
    Else
        MsgBox ("Не удалось сохранить файл")
    End If
End Sub

Sub TryRenameDraft(sname As String)
    If sname = "" Then
        Exit Sub
    End If
    
    Dim newname As String, oldname As String: oldname = gDoc.GetPathName
    If oldname = "" Then
        newname = GetPathName(gModel.GetPathName) + sname + ".SLDDRW"
        SaveAsMy newname, oldname
    Else
        newname = GetPathName(oldname) + sname + ".SLDDRW"
        If oldname <> newname Then
            SaveAsMy newname, oldname
        End If
    End If
End Sub

' Без точек "." в наименовании
Sub SplitNameAndSign(line As String, ByRef designation As String, ByRef name As String)
    Dim words As Variant
    Dim i, j As Integer
    
    designation = line
    name = line
    
    words = Split(line, " ")
    For i = UBound(words) To 0 Step -1
        If InStr(words(i), ".") <> 0 Then
            designation = words(0)
            For j = 1 To i
                designation = designation + " " + words(j)
            Next
            name = ""
            For j = i + 1 To UBound(words)
                If Not userDrawingTypes.Exists(words(j)) Then
                    name = name + " " + words(j)
                End If
            Next
            name = LTrim(name)
            Exit For
        End If
    Next
End Sub

'Only for drawings!
Function ZoomToSheet()  'mask for button
    Dim width As Double
    Dim height As Double

    gDrawing.GetCurrentSheet.GetSize width, height
    gDoc.ViewZoomTo2 0, 0, 0, width, height, 0
End Function

Function BubbleSort(ByVal arr As Variant) As Variant
    Dim i As Integer
    Dim j As Integer
    
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) > arr(j) Then
                tmp = arr(i)
                arr(i) = arr(j)
                arr(j) = tmp
            End If
        Next
    Next
    BubbleSort = arr
End Function

Function GetEquationThickness(conf As String, toAll As Boolean, nameModel As String) As String
    Const temp As String = "__temp__"
    Dim mgr As CustomPropertyManager
    Dim thName_ As Variant
    Dim thName As String
    Dim resolvedValue As String
    Dim rawValue As String
    Dim variantThicknessName(1) As String
    
    GetEquationThickness = ""
    Set mgr = gModel.Extension.CustomPropertyManager(conf)
    variantThicknessName(0) = "Толщина"
    variantThicknessName(1) = "Thickness"
    'variantThicknessName(2) = "Grubos'c'"
    For Each thName_ In variantThicknessName
        thName = thName_
        SetProp mgr, temp, Equal(thName, toAll, conf, gNameModel)
        mgr.Get5 temp, False, rawValue, resolvedValue, False
        If IsNumeric(StrConv(resolvedValue, vbUnicode)) Then  'IsNumeric make error with raw 'resolvedValue' without StrConv
            GetEquationThickness = rawValue
            Exit For
        End If
    Next
    mgr.Delete2 temp
End Function

Function GetBaseDesignation(designation As String) As String
    Dim lastFullstopPosition As Integer
    Dim firstHyphenPosition As Integer
    
    GetBaseDesignation = designation
    lastFullstopPosition = InStrRev(designation, ".")
    If lastFullstopPosition > 0 Then
        firstHyphenPosition = InStr(lastFullstopPosition, designation, "-")
        If firstHyphenPosition > 0 Then
            GetBaseDesignation = Left(designation, firstHyphenPosition - 1)
        End If
    End If
End Function
