VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "EditProp 23.4"
   ClientHeight    =   6720
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

Dim PreviousNameXX As String
Dim IsInitialized As Boolean

Private Sub ApplyBut_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetShiftStatus Shift
End Sub

Private Sub CodeBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ExitByKey KeyCode, Shift
End Sub

Private Sub lenLab_Click()
    Me.lenBox.Text = ""
End Sub

Private Sub MiniSignBox_Change()
    Dim Key As String
    Dim I As Variant
    
    Me.CodeBox.Enabled = True
    Me.CodeBox.Clear
    
    If Not SetPartCaptionIfEmptyDrawingCode Then
        Key = Me.MiniSignBox.Text
        If UserDrawingTypes.Exists(Key) Then
            For Each I In UserDrawingTypes(Key)
                Me.CodeBox.AddItem I
            Next
        End If
        If Me.CodeBox.ListCount > 0 Then
            Me.CodeBox.Text = Me.CodeBox.List(0)
        End If
    End If
End Sub

Private Sub NameBox_Change()
    CheckComboBoxAndWarnIfNeeded Me.NameBox
End Sub

Private Sub NameBoxTranslate_Change()
    CheckTextBoxAndWarnIfNeeded Me.NameBoxTranslate
End Sub

Private Sub RealFormatLab_Click()
    SetValueInBox RealFormatBox, 1
End Sub

Private Sub widLab_Click()
    Me.widBox.Text = ""
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

Private Sub MaterialLab_Click()
    SetValueInBox MaterialBox, 1
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

Private Sub NameBoxTranslate_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
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

Private Sub RealFormatBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ExitByKey KeyCode, Shift
End Sub

Private Sub PaperSizeBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
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
    FillNameAndSign ModelNameBox.Text, ConfBox.Text
End Sub

Private Sub DrawNameLab_Click()
    If gIsDrawing Then
        FillNameAndSign DrawNameBox.Text, ConfBox.Text
    End If
End Sub

Private Sub NameLang_Change()
    Dim pNameXX  As String
    
    pNameXX = GetNameXX
    If IsInitialized Then
        ReadNameLang gCurConf, PreviousNameXX
        ChangeChecked pNameXX  'because ConfBox is set after "IsInitialized = True"
    End If
    SaveLangTranslate Me.NameLang.ListIndex
    
    'It should be last
    PreviousNameXX = pNameXX
End Sub

Private Sub UserForm_Initialize()
    Dim I As Variant
    Dim LangIndex As Integer
    
    IsInitialized = False
    
    Set gItems = New Dictionary
    ReadOldAfterChecked = True
    InitWidgets
    
    ReadProp gModelManager, CommonSpace, ModelProps
    
    If gIsDrawing Then
        ReadProp gDrawManager, CommonSpace, DrawProps
        SetPartCaptionIfEmptyDrawingCode
        InitPaperSizeBox
    End If
    If Not gIsAssembly Then
        Me.IsFastenerChk.Value = GetIsFastener
    End If
    
    PreviousNameXX = ""
    For Each I In NameTranslateLangs.Keys
        Me.NameLang.AddItem I
    Next
    LangIndex = GetLangTranslateSetting
    If LangIndex >= 0 And LangIndex < Me.NameLang.ListCount Then
        Me.NameLang.ListIndex = LangIndex
    Else
        Me.NameLang.ListIndex = 0
    End If
    
    IsInitialized = True
    Me.ConfBox.Text = CreateConfItem(gCurConf)
End Sub

Private Sub ConfBox_Change()
    Dim Part As PartDoc
    
    If ConfBox.Text = "" Then Exit Sub
        
    If gItems.Exists(gCurConf) Then 'запись старой конфигурации
        ReadForm gCurConf
    End If
    
    gCurConf = gModelConfNames(ConfBox.Text)    'до этого в gCurConf записана стара€ конфигураци€
    
    If Not gItems.Exists(gCurConf) Then
        gModel.ShowConfiguration2 gCurConf 'ускор€ет чтение свойств
        ReadProp gModelExt.CustomPropertyManager(gCurConf), gCurConf, ModelProps
    End If
    ReloadForm gCurConf
End Sub

Private Sub SignChk_Change()
    TrySetPropToAll SignBox, SignChk, pDesignation
    Me.ConfBox.SetFocus
End Sub

Private Sub NameChk_Change()
    TrySetPropToAll NameBox, NameChk, pName
    TrySetPropToAll NameBoxTranslate, NameChk, GetNameXX
    Me.ConfBox.SetFocus
End Sub

Private Sub BlankChk_Change()
    TrySetPropToAll BlankBox, BlankChk, pBlank
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

Private Sub lenChk_Change()
    TrySetPropToAll lenBox, lenChk, pLen
    Me.ConfBox.SetFocus
End Sub

Private Sub widChk_Change()
    TrySetPropToAll widBox, widChk, pWid
    Me.ConfBox.SetFocus
End Sub

Private Sub CloseBut_Click()
    ExitApp
End Sub

Private Sub ApplyBut_Click()
    ApplyAndExitIfNeeded gIsShiftPressed
End Sub
