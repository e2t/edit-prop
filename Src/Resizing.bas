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

Sub ResizeSheetFormat( _
  Width As Double, Height As Double, CurrentSheet As Sheet, CurrentDoc As ModelDoc2, CurrentDraw As DrawingDoc, _
  OldWidth As Double, OldHeight As Double, SizeName As String)

  Const RightBottomBorderWidth = 0.19
  Const RightBottomBorderHeight = 0.06
  Const RightBottomNameWidth = 0.055
  Const RightBottomNameHeight = 0.004
  Const UnusedZ = 0

  Dim Sk As Sketch
  Dim I As Variant
  Dim J As Integer
  Dim P As SketchPoint
  Dim SelMgr As SelectionMgr
  Dim ANote As Note
  Dim ANoteCoord As Variant
  Dim SizeNameArray As Variant
  
  Set Sk = CurrentSheet.GetTemplateSketch
  For Each I In Sk.GetSketchPoints2
    Set P = I
    If IsEqual(P.X, OldWidth) And IsEqual(P.Y, OldHeight) Then
      CurrentDraw.EditTemplate
      P.SetCoords Width, Height, UnusedZ
      
      Set SelMgr = CurrentDoc.SelectionManager
      
      CurrentDoc.ClearSelection2 True
      CurrentDoc.Extension.SketchBoxSelect OldWidth, 0, UnusedZ, OldWidth - RightBottomNameWidth, RightBottomNameHeight, UnusedZ
      If SelMgr.GetSelectedObjectCount2(-1) > 0 Then
        If SelMgr.GetSelectedObjectType3(1, -1) = swSelNOTES Then
          Set ANote = SelMgr.GetSelectedObject6(1, -1)
          SizeNameArray = Split(ANote.GetText, " ")
          SizeNameArray(UBound(SizeNameArray)) = SizeName
          ANote.SetText Join(SizeNameArray, " ")
        End If
      End If
      
      CurrentDoc.ClearSelection2 True
      CurrentDoc.Extension.SketchBoxSelect OldWidth, 0, UnusedZ, OldWidth - RightBottomBorderWidth, RightBottomBorderHeight, UnusedZ
      For J = 1 To SelMgr.GetSelectedObjectCount2(-1)
        If SelMgr.GetSelectedObjectType3(J, -1) = swSelNOTES Then
          Set ANote = SelMgr.GetSelectedObject6(J, -1)
          ANoteCoord = ANote.GetTextPoint2
          ANote.SetTextPoint ANoteCoord(0) + Width - OldWidth, ANoteCoord(1), ANoteCoord(2)
        End If
      Next
      
      CurrentDraw.EditSheet
      Exit For
    End If
  Next
  
End Sub
