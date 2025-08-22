Sub Swapcells()
    Dim rng1 As Range, rng2 As Range
    Dim arr1, arr2
    
    If Selection.Areas.Count <> 2 Then Exit Sub
    
    Set rng1 = Selection.Areas(1)
    Set rng2 = Selection.Areas(2)
    
    arr1 = rng1.FormulaR1C1
    arr2 = rng2.FormulaR1C1
    
    rng1.FormulaR1C1 = arr2
    rng2.FormulaR1C1 = arr1
End Sub
