Attribute VB_Name = "mPoNumberGenerator"
Option Explicit

Public Function GeneratePoNumber(companyCode) As String

    Dim A As String:                        A = "891"   ' Company A
    Dim B As String:                        B = "121"   ' Company B
    Dim currentCompany As String:           currentCompany = companyCode
    
    Dim lastPoNumberPartial As String:      lastPoNumberPartial = max_number(currentCompany)
    Dim newPoNumber As String:              newPoNumber = Format(Val(lastPoNumberPartial) + 1, "0000")

    newPoNumber = currentCompany & newPoNumber
    GeneratePoNumber = newPoNumber

End Function

Private Function max_number(currentCompany As String) As String

    Dim dataRng As Range: Set dataRng = posTracker.Range("B6").CurrentRegion
    
    Dim i As Long
    Dim poNumbersArr() As Variant
    Dim maxValue As String: maxValue = "0000"
    Dim companyCodeLength As Integer: companyCodeLength = Len(currentCompany)
    
    If dataRng.Rows.Count = 1 Then
        max_number = ""
        Exit Function
    End If
    
    Set dataRng = dataRng.Resize(dataRng.Rows.Count - 1, 1).Offset(1, 0)
    
    If dataRng.Rows.Count = 1 Then
    
        maxValue = Format(Val(Right(dataRng.Value, 4)), "0000")
        
    ElseIf dataRng.Rows.Count > 1 Then
    
        ReDim poNumbersArr(1 To dataRng.Rows.Count, 0) As Variant
        poNumbersArr = dataRng.Value
    
        For i = 1 To UBound(poNumbersArr, 1)
            If Left(poNumbersArr(i, 1), companyCodeLength) = currentCompany Then
                If Format(Val(Right(poNumbersArr(i, 1), 4)), "0000") > Format(Val(maxValue), "0000") Then
                    maxValue = Right(poNumbersArr(i, 1), 4)
                End If
            End If
        Next i
        
    End If
    
    max_number = CStr(maxValue)
    
End Function

