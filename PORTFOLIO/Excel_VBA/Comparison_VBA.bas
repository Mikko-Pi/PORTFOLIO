Sub Comparison()
'VERY VERY USEFUL!!! THIS CODE IS AN EXAMPLE THAT INCLUDES VSTACK, APPEARANCE MODIFICATION, DUBLICTE SEARCH AND COMPARISON BETWEEN OLD AND NEW CHARTS
'Lets first erase the old data including all filters
Call turnoffunctionality
shBoth.Cells.Clear
'Selection.AutoFilter

'Then we declare data type variables and build our stacked table

Dim arr As Variant, arr2 As Variant
Dim i As Long, j As Long

arr = shNew.Range("A1").CurrentRegion
arr2 = shOld.Range("A1").CurrentRegion

'LET´S ADD A TIMER (JUST FOR FUN)
Dim start, finish, totaltime
start = Timer

'LET´S PUT OLD AND NEW CHARTS ON TOP OF EACH OTHER

For i = LBound(arr, 1) To UBound(arr, 1)
    For j = LBound(arr, 2) To UBound(arr, 2)
    
    shBoth.Cells(i, j) = arr(i, j)
    
    Next j
    'Debug.Print i
    Next i
    
    
For i = LBound(arr2, 1) + 1 To UBound(arr2, 1)
    For j = LBound(arr2, 2) To UBound(arr2, 2)
    
    
    shBoth.Cells(UBound(arr) + i - 1, j) = arr2(i, j)
    
    Next j
    
    Next i
    
   
'NOW LETS CREATE FILTERS AND PERFORM FILTERING TO ASSENDING ORDER:


Call filter_and_order


 'LETS GIVE SOME GREY COLOR FOR THE HEADER AREA:
 shBoth.Range("A1:D1").Interior.Color = 14277081


'DOUBLE BOTTOM LINE FOR DECORATING THE CHART AND IMPROVE READABILITY:

Call bottom_line

'LET´S FIND THE SAME POSITIONS BY FINDING THE DUNBLICATES FROM THE COLUMN A:A:

Dim arrVstack As Variant

arrVstack = shBoth.Range("A1").CurrentRegion

For i = LBound(arrVstack, 1) + 1 To UBound(arrVstack, 1)
    For j = LBound(arrVstack, 2) To UBound(arrVstack, 2)
    
    Application.WorksheetFunction.Clean (arrVstack(i, j))
    
    If Application.WorksheetFunction.CountIf(shBoth.Range("A2:A" & UBound(arrVstack, 1)), arrVstack(i, 1)) > 1 Then
    
    
    shBoth.Cells(i, j).Interior.Color = 16115392
    


    With shBoth.Cells(i, j).Borders(xlEdgeBottom)
        .LineStyle = xlDot
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

    
    End If
    
    Next j
    Next i
    
    'LET´S COMPARE VALUES BETWEEN THE OLD AND THE NEW TABLE:
    
For i = LBound(arrVstack) + 1 To UBound(arrVstack)
    For j = LBound(arrVstack, 2) To UBound(arrVstack, 2)
    
    If shBoth.Cells(i - 1, 1).Value = shBoth.Cells(i, 1).Value And shBoth.Cells(i - 1, j).Interior.Color = 16115392 Then
    
    If shBoth.Cells(i, j) <> shBoth.Cells(i - 1, j) Then
    
    shBoth.Cells(i, j).Font.Color = vbRed
    shBoth.Cells(i - 1, j).Font.Color = vbRed
   
    End If
    End If
    
    Next j
    Next i
    
'LET´S FINISH THE TIMER FOR FUN :)
finish = Timer
totaltime = finish - start
Call turnonunctionality

Debug.Print "Time is:" & totaltime
End Sub

'FROM THIS LINE ALL THE WAY DOWN THE SUBS ARE JUST SOME HELPER SUBS THAT I CALL

Public Sub turnoffunctionality()
'THESE CODES HELPS YOU TO MAKE YOUR CODE FASTER AND POTENTIONALLY PREVENTS YOUR EXCEL FROM CRUSHING :)


Application.EnableEvents = False
Application.ScreenUpdating = False



End Sub
Public Sub turnonunctionality()


Application.EnableEvents = True
Application.ScreenUpdating = True


End Sub
Sub Underline()
'
' Underline Macro ()
'

'
    Range("A4").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    
    Range("A7").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
End Sub
Public Sub bottom_line()
'

    
    With shBoth.Range("A1:D1").Borders(xlEdgeBottom)
        .LineStyle = xlDouble
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
    

End Sub

Public Sub filter_and_order()

Dim arr As Variant
arr = shBoth.Range("A1").CurrentRegion

With shBoth.Range("A1:D1")
'Debug.Print shBoth.Range("A2:A" & UBound(arr, 2)).Address
'Debug.Print UBound(arr, 2)
.AutoFilter
End With

    shBoth.AutoFilter.Sort.SortFields.Clear
    shBoth.AutoFilter.Sort.SortFields.Add2 Key _
        :=Range("A2"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With shBoth.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
