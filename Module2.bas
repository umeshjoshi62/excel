Sub FindCombination()
Dim candidates As Variant
target = InputBox("enter target number")
candidates = ReadNumbersToArray()
result = CombinationSum(candidates, CLng(target))


strMessage = ""
For i = LBound(result) To UBound(result)
strMessage = strMessage & Join(result(i), ", ") & vbCrLf
Next


MsgBox strMessage


End Sub

Function ReadNumbersToArray()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim numbers() As Double
    Dim count As Long
    
    ' Set the worksheet (you can change "Sheet1" to your sheet name)
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' Find the last row in column A
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    
    ' Initialize the array with a size equal to the number of numeric entries
    ReDim numbers(1 To lastRow) ' Assuming all entries are numeric for now
    
    ' Loop through each cell in column A
    count = 0
    For i = 1 To lastRow
        If IsNumeric(ws.Cells(i, 1).value) Then
            count = count + 1
            numbers(count) = ws.Cells(i, 1).value
        End If
    Next i
    
    ' Resize the array to the actual number of numeric entries
    If count > 0 Then
        ReDim Preserve numbers(1 To count)
    Else
        MsgBox "No numeric values found in column A."
        Exit Function
    End If
    
    ' Output the numbers to the Immediate Window (Ctrl + G to view)
    For i = LBound(numbers) To UBound(numbers)
        Debug.Print numbers(i)
    Next i
     
    ReadNumbersToArray = numbers
    ' Optional: You can return the array or use it for further processing
End Function

