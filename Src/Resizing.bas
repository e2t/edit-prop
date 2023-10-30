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

Sub SetPaperSizeToSheetFormat(SizeName As String, CurrentSheet As Sheet)
    Dim I As Variant
    Dim J As Variant
    Dim AView As View
    Dim ANote As Note
    Dim CurrentDrawing As DrawingDoc
    
    Set CurrentDrawing = gDoc
    For Each I In CurrentDrawing.GetViews 'array of array
        Set AView = I(0)
        If AView.Name = CurrentSheet.GetName Then
            For Each J In AView.GetNotes
                Set ANote = J
                If ANote.TagName = TagPaperSize Then
                    ANote.SetText SizeName
                End If
            Next
        End If
    Next
End Sub

Sub ResizeSheetFormat( _
    Width As Double, Height As Double, CurrentSheet As Sheet, CurrentDoc As ModelDoc2, CurrentDraw As DrawingDoc, _
    OldWidth As Double, OldHeight As Double, SizeName As String)
    
    Const RightBottomBorderWidth = 0.19
    Const RightBottomBorderHeight = 0.065
    Const UnusedZ = 0
    
    Dim SelMgr As SelectionMgr
    Dim Sk As Sketch
    Dim I As Variant
    Dim P As SketchPoint
    
    Set SelMgr = CurrentDoc.SelectionManager
    Set Sk = CurrentSheet.GetTemplateSketch
    For Each I In Sk.GetSketchPoints2
        Set P = I
        If IsEqual(P.X, OldWidth) And IsEqual(P.Y, OldHeight) Then
            CurrentDraw.EditTemplate
            
            SetPaperSizeToSheetFormat SizeName, CurrentSheet
            
            CurrentDoc.ClearSelection2 True
            CurrentDoc.Extension.SketchBoxSelect OldWidth, 0, UnusedZ, OldWidth - RightBottomBorderWidth, Height, UnusedZ
            gDoc.Extension.MoveOrCopy False, 0, True, 0, 0, 0, Width - OldWidth, 0, 0
            P.SetCoords P.X, Height, P.Z
            
            CurrentDraw.EditSheet
            Exit For
        End If
    Next
End Sub
