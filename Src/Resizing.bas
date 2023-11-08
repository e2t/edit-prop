Attribute VB_Name = "Resizing"
Option Explicit

Sub GetSheetSizes( _
    ByRef Width As Double, ByRef Height As Double, ByRef SizeName As String, _
    ByRef OldWidth As Double, ByRef OldHeight As Double, _
    CurrentSheet As Sheet, PaperChoice As String)
    
    Const Accuracy = 0.002
    Dim I As Variant
    Dim Item As TPaperSizeItem
    
    CurrentSheet.GetSize OldWidth, OldHeight
    If PaperChoice = CurrentChoice Then
        Width = OldWidth
        Height = OldHeight
        SizeName = Str(Round(Width * 1000#)) & "x" & Str(Round(Height * 1000#))
        For Each I In PaperSizes.items
            Set Item = I
            If IsEqual(Item.Width, Width, Accuracy) And IsEqual(Item.Height, Height, Accuracy) Then
                SizeName = Item.Name
                Exit For
            End If
        Next
    Else
        Set Item = PaperSizes(PaperChoice)
        Width = Item.Width
        Height = Item.Height
        SizeName = Item.Name
    End If
End Sub

Sub SetPaperSizeToSheetFormat(SizeName As String, CurrentSheetName As String)
    Dim I As Variant
    Dim AView As View
    Dim ANote As Note
    Dim CurrentDrawing As DrawingDoc
        
    Set CurrentDrawing = gDoc
    For Each I In CurrentDrawing.GetViews 'array of array
        Set AView = I(0)
        If AView.Name = CurrentSheetName Then
            Exit For
        End If
    Next
    For Each I In AView.GetNotes
        Set ANote = I
        If ANote.TagName = TagPaperSize Then
            ANote.SetText "SIZE " + SizeName
            Exit For
        End If
    Next
End Sub

Sub ResizeSheetFormat( _
    Width As Double, Height As Double, CurrentSheet As Sheet, CurrentDoc As ModelDoc2, _
    CurrentDraw As DrawingDoc, OldWidth As Double, OldHeight As Double, SizeName As String)
    
    Const UnusedZ = 0
    Dim Sk As Sketch
    Dim I As Variant
    Dim P As SketchPoint
    
    SetPaperSizeToSheetFormat SizeName, CurrentSheet.GetName
    CurrentDraw.EditTemplate
    
    Set Sk = CurrentSheet.GetTemplateSketch
    For Each I In Sk.GetSketchPoints2
        Set P = I
        If IsEqual(P.X, OldWidth) And IsEqual(P.Y, OldHeight) Then
            Exit For
        End If
    Next
    
    If OldHeight <> Height Then
        CurrentDoc.ClearSelection2 True
        CurrentDoc.Extension.SketchBoxSelect OldWidth, OldHeight, UnusedZ, OldWidth - 0.1, OldHeight - 0.1, UnusedZ
        gDoc.Extension.MoveOrCopy False, 0, True, 0, OldHeight, UnusedZ, 0, Height, UnusedZ
    End If
    
    If OldWidth <> Width Then
        CurrentDoc.ClearSelection2 True
        CurrentDoc.Extension.SketchBoxSelect OldWidth, 0, UnusedZ, OldWidth - 0.19, Height, UnusedZ
        gDoc.Extension.MoveOrCopy False, 0, True, OldWidth, 0, UnusedZ, Width, 0, UnusedZ
        'Иногда рамка не перемещается. Нижний блок заставляет управляющую точку смещаться,
        'пока это не случится. Есть риск зависания.
        While IsEqual(P.X, OldWidth)
            P.SetCoords Width, P.Y, P.Z
        Wend
    End If
    
    CurrentDraw.EditSheet
End Sub
