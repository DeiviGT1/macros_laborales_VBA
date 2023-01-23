Attribute VB_Name = "VxP"
Sub VxP()
Dim i As Integer
Dim pregunta As Integer
On Error Resume Next

    i = 1
    WS_Count = ActiveWorkbook.Worksheets.Count

    Do Until i > WS_Count
        Sheets(i).Select
1
    'Funciones acá
        Columns("A:P").Select
    Range("P1").Activate
    Selection.AutoFilter
    ActiveSheet.Range("A:P").AutoFilter Field:=11, Criteria1:=">3", _
        Operator:=xlAnd
    ActiveSheet.Range("A:P").AutoFilter Field:=14, Criteria1:="<0.2", _
        Operator:=xlAnd
    Columns("N:P").Select
    Range("P1").Activate
    Selection.FormatConditions.AddColorScale ColorScaleType:=3
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).ColorScaleCriteria(1).Type = _
        xlConditionValueLowestValue
    With Selection.FormatConditions(1).ColorScaleCriteria(1).FormatColor
        .Color = 7039480
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).ColorScaleCriteria(2).Type = _
        xlConditionValuePercentile
    Selection.FormatConditions(1).ColorScaleCriteria(2).Value = 50
    With Selection.FormatConditions(1).ColorScaleCriteria(2).FormatColor
        .Color = 8711167
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).ColorScaleCriteria(3).Type = _
        xlConditionValueHighestValue
    With Selection.FormatConditions(1).ColorScaleCriteria(3).FormatColor
        .Color = 8109667
        .TintAndShade = 0
    End With
        
        
        i = i + 1
        
        
    Loop
End Sub




Sub VxP_EC()
Dim i As Integer
Dim pregunta As Integer
On Error Resume Next

    i = 1
    WS_Count = ActiveWorkbook.Worksheets.Count

    Do Until i > WS_Count
        Sheets(i).Select
        
    'Funciones acá
        Columns("A:S").Select
    Range("S1").Activate
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$S$31").AutoFilter Field:=11, Criteria1:=">3", _
        Operator:=xlAnd
    ActiveSheet.Range("$A$1:$S$31").AutoFilter Field:=14, Criteria1:="<0.2", _
        Operator:=xlAnd
    Columns("N:S").Select
    Range("S1").Activate
    Selection.FormatConditions.AddColorScale ColorScaleType:=3
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).ColorScaleCriteria(1).Type = _
        xlConditionValueLowestValue
    With Selection.FormatConditions(1).ColorScaleCriteria(1).FormatColor
        .Color = 7039480
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).ColorScaleCriteria(2).Type = _
        xlConditionValuePercentile
    Selection.FormatConditions(1).ColorScaleCriteria(2).Value = 50
    With Selection.FormatConditions(1).ColorScaleCriteria(2).FormatColor
        .Color = 8711167
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).ColorScaleCriteria(3).Type = _
        xlConditionValueHighestValue
    With Selection.FormatConditions(1).ColorScaleCriteria(3).FormatColor
        .Color = 8109667
        .TintAndShade = 0
    End With
        
        
        i = i + 1
        
        
    Loop
End Sub

