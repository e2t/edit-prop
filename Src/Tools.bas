Attribute VB_Name = "Tools"
Option Explicit

Sub AppendPaperSize(ChoiceName As String, SizeName As String, Width As Double, Height As Double)

  Dim Item As TPaperSizeItem

  Set Item = New TPaperSizeItem
  Item.Name = SizeName
  Item.Width = Width
  Item.Height = Height
  PaperSizes.Add ChoiceName, Item

End Sub

Function IsDrawing(Doc As ModelDoc2) As Boolean

  IsDrawing = CBool(Doc.GetType = swDocDRAWING)
    
End Function

Function IsEqual(A As Double, B As Double, Optional Accuracy As Double = 0.000001) As Boolean

  IsEqual = Abs(A - B) <= Accuracy
  'If IsEqual Then MsgBox Str(A) & "x" & Str(B)

End Function

Function IsPropertyExist(Manager As CustomPropertyManager, Prop As String) As Boolean

  Dim Names As Variant
  
  Names = Manager.GetNames()
  IsPropertyExist = False
  If Not IsEmpty(Names) Then
    If IndexInArray(Prop, Names) <> -1 Then
      IsPropertyExist = True
    End If
  End If
  
End Function

Function SetProp(Manager As CustomPropertyManager, Prop As String, Value As String) As Boolean

  Manager.Add2 Prop, swCustomInfoText, ""
  SetProp = Not CBool(Manager.Set(Prop, Trim(Value)))
  
End Function

Function IsArrayEmpty(anArray As Variant) As Boolean

  Dim I As Integer

  On Error GoTo ArrayIsEmpty
  IsArrayEmpty = LBound(anArray) > UBound(anArray)
  Exit Function
  
ArrayIsEmpty:
  IsArrayEmpty = True

End Function

Sub InitWidgetFrom(Widget As Object, Values As Variant)

  Dim I As Variant
  
  If Not IsArrayEmpty(Values) Then
    For Each I In Values
      Widget.AddItem I
    Next
  End If
  
End Sub

Function CheckIsStandardScale(ByVal Scale1 As Double, ByVal Scale2 As Double) As Boolean

  NormalizeScales Scale1, Scale2
  If Scale1 = 1 Then
    Select Case Scale2
      Case 1, 2, 2.5, 4, 5, 10, 15, 20, 25, 40, 50, 75, 100, 200, 400, 500, 800, 1000, 2000, 5000, 10000, 20000, 25000, 50000
        CheckIsStandardScale = True
      Case Else
        CheckIsStandardScale = False
    End Select
  Else
    Select Case Scale1
      Case 2, 2.5, 4, 5, 10, 20, 40, 50, 100
        CheckIsStandardScale = True
      Case Else
        CheckIsStandardScale = (Scale1 Mod 100 = 0)
    End Select
  End If

End Function

Sub NormalizeScales(ByRef Scale1 As Double, ByRef Scale2 As Double)

  Dim MinNumber As Double
  
  If (Scale1 <> 1) And (Scale2 <> 1) Then
    MinNumber = IIf(Scale1 < Scale2, Scale1, Scale2)
    Scale1 = Scale1 / MinNumber
    Scale2 = Scale2 / MinNumber
  End If

End Sub

Sub ReadHeaderValues(ByRef UserStr() As String, ByRef Count As Long, Lines() As String, EndLines As Long)

  If Count < EndLines Then
    Count = Count + 1
    If Lines(Count) <> "" Then
      UserStr = Split(Lines(Count), Separator)
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

Sub MsgDebug(Msg As Variant)

  If MsgBox(Msg, vbOKCancel) = vbCancel Then ExitApp
  
End Sub

Function HeaderInFile(Prop As String) As String

  HeaderInFile = "[" + Prop + "]"
  
End Function

Function ChangeLineStyles() 'hide

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

Function IntOrNul(Str As String) As Long

  IntOrNul = 0
  If IsNumeric(Str) Then
    IntOrNul = CInt(Str)
  End If
    
End Function

Function FirstItem(Values As Variant) As String

  FirstItem = ""
  If Not IsEmpty(Values) Then
    FirstItem = Values(0)
  End If
    
End Function

Function IndexInArray(ValueToFind As Variant, Arr As Variant) As Integer

  Dim I As Integer
  
  IndexInArray = -1
  For I = LBound(Arr) To UBound(Arr)
    If Arr(I) = ValueToFind Then
      IndexInArray = I
      Exit Function
    End If
  Next
    
End Function

Function GetBaseDesignation(Designation As String) As String

  Dim LastFullstopPosition As Integer
  Dim FirstHyphenPosition As Integer
  
  GetBaseDesignation = Designation
  LastFullstopPosition = InStrRev(Designation, ".")
  If LastFullstopPosition > 0 Then
    FirstHyphenPosition = InStr(LastFullstopPosition, Designation, "-")
    If FirstHyphenPosition > 0 Then
      GetBaseDesignation = Left(Designation, FirstHyphenPosition - 1)
    End If
  End If
  
End Function

Sub RewriteNameAndSign(Source As String, Conf As String)

  Dim Designation As String
  Dim Name As String
  Dim Code As String
  Dim I As Integer
  Dim IsCodeFound As Boolean
  
  Designation = ""
  Name = ""
  SplitNameAndSign Source, Conf, Designation, Name, Code
  MainForm.SignBox.Text = Designation
  MainForm.NameBox.Text = Name
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

Function ExistsInCombo(Box As ComboBox, Value As String)

  Dim I As Variant

  ExistsInCombo = False
  If Box.ListCount > 0 Then
    For Each I In Box.List
      If I = Value Then
        ExistsInCombo = True
        Exit For
      End If
    Next
  End If
    
End Function

Sub SetComboInExistValue(ByRef Box As Object, Value As String)

  On Error Resume Next  ''''ПОДАВЛЕНИЕ ОШИБКИ для Гордиенко
  Box.Text = Value
    
End Sub

Sub TrySetPropToAll(Box As Object, Chk As CheckBox, Property As String)

  If IsShiftPressed And Not Chk.Value Then
    IsShiftPressed = False
    
    ReadOldAfterChecked = False
    Chk.Value = True
    ReadOldAfterChecked = True
    
    If MsgBox("Связать все конфигурации со значением?", vbYesNo) = vbYes Then
      SetPropToAll Box, Chk, Property
    End If
  Else
    ChangeChecked Property
  End If
  
End Sub

Sub SetValueInBox(Box As ComboBox, Index As Integer)

  If 0 <= Index And Index < Box.ListCount Then
    Box.Text = Box.List(Index)
  End If
    
End Sub

Sub ExitByKey(KeyCode As MSForms.ReturnInteger, Shift As Integer)

  If Shift = 1 And KeyCode = vbKeyReturn Then
    Execute
    ExitApp
  End If
    
End Sub

Function SetPartCaptionIfEmptyDrawingCode() As Boolean

  SetPartCaptionIfEmptyDrawingCode = (MainForm.MiniSignBox.Text = "")
  If SetPartCaptionIfEmptyDrawingCode Then
    MainForm.CodeBox.AddItem "Деталь"
    MainForm.CodeBox.Text = MainForm.CodeBox.List(0)
    MainForm.CodeBox.Enabled = False
  End If

End Function

Function InitPaperSizeBox() 'hide

  Dim I As Variant
  
  MainForm.PaperSizeBox.AddItem CurrentChoice
  MainForm.PaperSizeBox.Text = MainForm.PaperSizeBox.List(0)
  For Each I In PaperSizes.Keys
    MainForm.PaperSizeBox.AddItem I
  Next

End Function

Function Equal(swProp As String, ToAll As Boolean, Conf As String, NameModel As String) As String

  Dim ConfText As String
  
  ConfText = ""
  If Conf <> CommonSpace And Not ToAll Then
    ConfText = "@@" + Conf
  End If
  Equal = """" + swProp + ConfText + "@" + NameModel + """"
    
End Function

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

Sub ReloadForm(Conf As String)

  ReadOldAfterChecked = False

  If gIsDrawing Then
    SetBoxValue2 Nothing, pShortDrawingType, CommonSpace
    SetBoxValue2 Nothing, pLongDrawingType, CommonSpace
    SetBoxValue2 Nothing, pOrganization, CommonSpace
    SetBoxValue2 Nothing, pDrafter, CommonSpace
    SetBoxValue2 Nothing, pChecking, CommonSpace
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
  ReadOldAfterChecked = True
    
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
    ReadBox MainForm.MiniSignBox, Nothing, CommonSpace, pShortDrawingType, True
    ReadBox MainForm.CodeBox, Nothing, CommonSpace, pLongDrawingType, True
    ReadBox MainForm.OrgBox, Nothing, CommonSpace, pOrganization, True
    ReadBox MainForm.DraftBox, Nothing, CommonSpace, pDrafter, True
    ReadBox MainForm.CheckingBox, Nothing, CommonSpace, pChecking, True
  End If
  
  If Not gIsAssembly Then
    ReadBox MainForm.BlankBox, MainForm.BlankChk, Conf, pBlank, True
    ReadBox MainForm.SizeBox, MainForm.SizeChk, Conf, pSize, True
    ReadBox MainForm.MaterialBox, Nothing, Conf, pMaterial, True
    ReadBox MainForm.lenBox, MainForm.lenChk, Conf, pLen, True
    ReadBox MainForm.widBox, MainForm.widChk, Conf, pWid, True
  End If
  
End Sub

Function SetProp2(Manager As CustomPropertyManager, Prop As String, Item As DataItem, _
                  Optional Conf As String = CommonSpace) As Boolean
                  
  Dim Result As Boolean
  
  Result = False
  If Prop = pMaterial Then
    Result = SetProp(Manager, Prop, MaterialEqual(Conf))
  Else
    Result = SetProp(Manager, Prop, Item.NewValue)
  End If
  SetProp2 = Result
    
End Function
