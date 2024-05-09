Attribute VB_Name = "mdl_noise"
Sub Noise()

Dim numLayers As Integer
Dim i As Integer

Dim arr     As clsArray:     Set arr = New clsArray
Dim combArr As clsArray:     Set combArr = New clsArray


numLayers = 6

Dim layerArr(1 To 6) As clsArray


For i = 1 To numLayers
    Set layerArr(i) = New clsArray
Next


Dim maxR As Long
Dim maxC As Long
Dim maxH As Long
Dim minH As Long
Dim depth As Long
Dim multi As Long

maxR = 100
maxC = 100
maxH = 1
minH = -1

depth = 3
multi = 3


arr.SetArr2D maxR, maxC
combArr.SetArr2D maxR, maxC

Dim r As Long
Dim c As Long



For i = 1 To numLayers
    
    Debug.Print "Depth: " & depth & ", Layer: " & i
    For r = 1 To arr.ArrSize(1)
        For c = 1 To arr.ArrSize(2)
            arr.Member2D(r, c) = WorksheetFunction.RandBetween(minH, maxH)
        Next
    Next

    arr.Init smoothing_grid(arr, 1, depth)
    
    layerArr(i).Init arr.Members
        
    minH = minH * multi
    maxH = maxH * multi
    depth = depth * multi
Next



For i = 1 To numLayers
    For r = 1 To arr.ArrSize(1)
        For c = 1 To arr.ArrSize(2)
            combArr.AddToMember2D r, c, layerArr(i).Member2D(r, c)
        Next
    Next
Next

combArr.SetRange ActiveWorkbook.Sheets("Combined"), "A1", True

Set arr = Nothing
Set combArr = Nothing

End Sub



Private Function smoothing_grid(arr As clsArray, Optional currDepth As Long = 1, Optional depth As Long = 1) As Variant()

Dim arr2 As clsArray: Set arr2 = New clsArray

Dim firstRow As Boolean: firstRow = False
Dim firstCol As Boolean: firstCol = False
Dim lastRow As Boolean: lastRow = False
Dim lastCol As Boolean: lastCol = False

Dim up As Variant
Dim down As Variant
Dim left As Variant
Dim right As Variant
Dim result As Variant

Dim maxR As Long
Dim maxC As Long

Dim r As Long
Dim c As Long

maxR = arr.ArrSize(1)
maxC = arr.ArrSize(2)

arr2.SetArr2D maxR, maxC


For r = 1 To maxR

    If r = 1 Then firstRow = True
    If r = maxR Then lastRow = True
    
    For c = 1 To maxC
    
        If c = 1 Then firstCol = True
        If c = maxC Then lastCol = True
        
        If firstRow And firstCol Then
            right = arr.Member2D(r, c + 1)
            down = arr.Member2D(r + 1, c)
            
            result = WorksheetFunction.Average(arr.Member2D(r, c), right, down)
        ElseIf lastRow And lastCol Then
            left = arr.Member2D(r, c - 1)
            up = arr.Member2D(r - 1, c)
            
            result = WorksheetFunction.Average(arr.Member2D(r, c), left, up)
        ElseIf firstRow And lastCol Then
            left = arr.Member2D(r, c - 1)
            down = arr.Member2D(r + 1, c)
            
            result = WorksheetFunction.Average(arr.Member2D(r, c), left, down)
        ElseIf lastRow And firstCol Then
            right = arr.Member2D(r, c + 1)
            up = arr.Member2D(r - 1, c)
            
            result = WorksheetFunction.Average(arr.Member2D(r, c), right, up)
        ElseIf firstRow Then
            right = arr.Member2D(r, c + 1)
            down = arr.Member2D(r + 1, c)
            left = arr.Member2D(r, c - 1)
            
            result = WorksheetFunction.Average(arr.Member2D(r, c), right, down, left)
        ElseIf firstCol Then
            up = arr.Member2D(r - 1, c)
            down = arr.Member2D(r + 1, c)
            right = arr.Member2D(r, c + 1)
            
            result = WorksheetFunction.Average(arr.Member2D(r, c), up, down, right)
        ElseIf lastRow Then
            up = arr.Member2D(r - 1, c)
            right = arr.Member2D(r, c + 1)
            left = arr.Member2D(r, c - 1)
            
            result = WorksheetFunction.Average(arr.Member2D(r, c), right, up, left)
        ElseIf lastCol Then
            up = arr.Member2D(r - 1, c)
            down = arr.Member2D(r + 1, c)
            left = arr.Member2D(r, c - 1)
            
            result = WorksheetFunction.Average(arr.Member2D(r, c), right, up, left)
        Else
            up = arr.Member2D(r - 1, c)
            down = arr.Member2D(r + 1, c)
            left = arr.Member2D(r, c - 1)
            right = arr.Member2D(r, c + 1)
            
            result = WorksheetFunction.Average(arr.Member2D(r, c), right, up, left, down)
        End If
        
        arr2.Member2D(r, c) = result
        
        firstCol = False
        lastCol = False
        
    Next
    
    firstRow = False
    lastRow = False
Next
    If currDepth = depth Then
        smoothing_grid = arr2.Members
    Else
        currDepth = currDepth + 1
        smoothing_grid = smoothing_grid(arr2, currDepth, depth)
    End If
    
    Set arr2 = Nothing
    
End Function


