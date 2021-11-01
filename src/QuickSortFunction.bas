Attribute VB_Name = "QuickSortFunction"
Option Explicit

'https://stackoverflow.com/questions/152319/vba-array-sort-function
Public Sub QuickSort(vArray As Variant, inLow As Long, inHi As Long)

  Dim Pivot As Variant
  Dim TmpSwap As Variant
  Dim TmpLow As Long
  Dim TmpHi As Long
  
  TmpLow = inLow
  TmpHi = inHi
  
  Pivot = vArray((inLow + inHi) \ 2)
  
  While (TmpLow <= TmpHi)
    While (vArray(TmpLow) < Pivot And TmpLow < inHi)
      TmpLow = TmpLow + 1
    Wend
    
    While (Pivot < vArray(TmpHi) And TmpHi > inLow)
      TmpHi = TmpHi - 1
    Wend
    
    If (TmpLow <= TmpHi) Then
      TmpSwap = vArray(TmpLow)
      vArray(TmpLow) = vArray(TmpHi)
      vArray(TmpHi) = TmpSwap
      TmpLow = TmpLow + 1
      TmpHi = TmpHi - 1
    End If
  Wend
  
  If (inLow < TmpHi) Then QuickSort vArray, inLow, TmpHi
  If (TmpLow < inHi) Then QuickSort vArray, TmpLow, inHi
    
End Sub
