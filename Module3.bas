Sub TestCombinationSum()
    Dim candidates As Variant
    Dim target As Long
    Dim result As Variant
    Dim i As Long
    
    candidates = Array(2, 3, 6, 7)
    target = 7
    
    result = CombinationSum(candidates, target)
    
    For i = LBound(result) To UBound(result)
        Debug.Print "Combination: " & Join(result(i), ", ")
    Next i
End Sub
