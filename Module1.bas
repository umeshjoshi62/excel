Option Explicit

Function CombinationSum(candidates As Variant, target As Long) As Variant
    Dim result As Collection
    Set result = New Collection
    
    Dim current As Collection
    Set current = New Collection
    
    Call DFS(1, candidates, current, 0, target, result)
    
    Dim output() As Variant
    Dim i As Long
    ReDim output(0 To result.count - 1)
    
    For i = 1 To result.count
        output(i - 1) = result(i)
    Next i
    
    CombinationSum = output
End Function

Sub DFS(i As Long, candidates As Variant, current As Collection, total As Long, target As Long, result As Collection)
    If total = target Then
        Dim combination As Variant
        combination = CollectionToArray(current)
        result.Add combination
        Exit Sub
    ElseIf i >= UBound(candidates) + 1 Or total > target Then
        Exit Sub
    End If
    
    ' Include the current candidate
    current.Add candidates(i)
    Call DFS(i, candidates, current, total + candidates(i), target, result)
    
    ' Backtrack
    current.Remove current.count
    
    ' Exclude the current candidate and move to the next
    Call DFS(i + 1, candidates, current, total, target, result)
End Sub

Function CollectionToArray(col As Collection) As Variant
    Dim arr() As Variant
    Dim i As Long
    ReDim arr(0 To col.count - 1)
    
    For i = 1 To col.count
        arr(i - 1) = col(i)
    Next i
    
    CollectionToArray = arr
End Function
